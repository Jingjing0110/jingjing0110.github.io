#!/usr/bin/env python3
"""Export the site's research diagrams as fully editable one-slide PowerPoints.

This intentionally converts SVG primitives to PowerPoint objects; it never embeds
an SVG or a screenshot. Text stays in textboxes, circles/rectangles remain native
shapes, and SVG paths become DrawingML freeforms with editable Bézier points.

Dependencies: python-pptx, svgpathtools, Pillow.
Example (using isolated dependencies):
    PYTHONPATH=/tmp/jingjing-ppt-tools python scripts/svg_to_editable_pptx.py
    PYTHONPATH=/tmp/jingjing-ppt-tools python scripts/svg_to_editable_pptx.py \
        --only surveillance

The renderer deliberately rejects unsupported visible SVG elements rather than
silently rasterizing or dropping them. It supports the primitives in these two
source figures, including inherited styles, transforms, and open arrow markers.
"""

from __future__ import annotations

import argparse
import math
from pathlib import Path
import re
import subprocess
import xml.etree.ElementTree as ET
from zipfile import ZipFile

from lxml import etree
from PIL import ImageFont
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Pt
from svgpathtools import Arc, CubicBezier, Line, QuadraticBezier, parse_path


REPO = Path(__file__).resolve().parents[1]
EMU_PER_PX = 914400 / 96
SVG_NS = "http://www.w3.org/2000/svg"
IDENTITY = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)
STYLE_KEYS = {
    "fill", "stroke", "stroke-width", "stroke-linecap", "stroke-linejoin",
    "font-family", "font-size", "font-style", "font-weight", "text-anchor",
    "letter-spacing", "opacity", "fill-opacity", "stroke-opacity", "marker-end",
}
JOBS = {
    "surveillance": ("research-surveillance.svg", "Wastewater_Surveillance_Editable.pptx"),
    "recovery": ("research-resource-recovery.svg", "Resource_Recovery_Editable.pptx"),
}


def px(value: float) -> int:
    return round(value * EMU_PER_PX)


def element(tag: str, **attrs):
    node = OxmlElement(tag)
    for key, value in attrs.items():
        node.set(key, str(value))
    return node


def multiply(a, b):
    """Affine a * b, using SVG's six-number matrix convention."""
    a0, a1, a2, a3, a4, a5 = a
    b0, b1, b2, b3, b4, b5 = b
    return (
        a0 * b0 + a2 * b1, a1 * b0 + a3 * b1,
        a0 * b2 + a2 * b3, a1 * b2 + a3 * b3,
        a0 * b4 + a2 * b5 + a4, a1 * b4 + a3 * b5 + a5,
    )


def transform_point(z: complex, matrix) -> complex:
    a, b, c, d, e, f = matrix
    return complex(a * z.real + c * z.imag + e, b * z.real + d * z.imag + f)


def matrix_for(value: str):
    matrix = IDENTITY
    for name, values in re.findall(r"([A-Za-z]+)\(([^)]*)\)", value or ""):
        numbers = [float(n) for n in re.findall(r"[-+]?(?:\d*\.\d+|\d+\.?\d*)(?:[eE][-+]?\d+)?", values)]
        if name == "translate":
            local = (1, 0, 0, 1, numbers[0], numbers[1] if len(numbers) > 1 else 0)
        elif name == "scale":
            local = (numbers[0], 0, 0, numbers[-1], 0, 0)
        elif name == "matrix":
            local = tuple(numbers)
        elif name == "rotate":
            angle = math.radians(numbers[0])
            local = (math.cos(angle), math.sin(angle), -math.sin(angle), math.cos(angle), 0, 0)
            if len(numbers) == 3:
                x, y = numbers[1:]
                local = multiply(multiply((1, 0, 0, 1, x, y), local), (1, 0, 0, 1, -x, -y))
        else:
            raise ValueError(f"Unsupported SVG transform: {name}")
        matrix = multiply(matrix, local)
    return matrix


def style_for(node, inherited):
    style = dict(inherited)
    style.update({k: v for k, v in node.attrib.items() if k in STYLE_KEYS})
    for declaration in node.get("style", "").split(";"):
        if ":" in declaration:
            key, value = declaration.split(":", 1)
            if key.strip() in STYLE_KEYS:
                style[key.strip()] = value.strip()
    return style


def color_value(value):
    if value in (None, "none", "transparent"):
        return None
    if re.fullmatch(r"#[0-9a-fA-F]{6}", value):
        return value[1:].upper()
    if re.fullmatch(r"#[0-9a-fA-F]{3}", value):
        return "".join(c * 2 for c in value[1:]).upper()
    raise ValueError(f"Unsupported SVG color: {value}")


def opacity_for(color_node, opacity):
    if opacity < 1:
        color_node.append(element("a:alpha", val=round(opacity * 100000)))


def apply_appearance(shape, style, matrix=IDENTITY):
    # The default PowerPoint shape style includes a shadow. Every appearance
    # property here is explicit, so remove that theme style and retain the
    # source SVG's clean, flat vector treatment.
    theme_style = shape._element.find(qn("p:style"))
    if theme_style is not None:
        shape._element.remove(theme_style)
    shape._element.spPr.append(element("a:effectLst"))
    fill = color_value(style.get("fill", "#000000"))
    stroke = color_value(style.get("stroke"))
    opacity = float(style.get("opacity", 1))
    if fill is None:
        shape.fill.background()
    else:
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor.from_string(fill)
        opacity_for(shape.fill._xPr.solidFill[0], opacity * float(style.get("fill-opacity", 1)))
    if stroke is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = RGBColor.from_string(stroke)
        scale = math.sqrt(abs(matrix[0] * matrix[3] - matrix[1] * matrix[2]))
        shape.line.width = Pt(float(style.get("stroke-width", 1)) * scale * .75)
        ln = shape.line._get_or_add_ln()
        ln.set("cap", {"round": "rnd", "square": "sq"}.get(style.get("stroke-linecap"), "flat"))
        ln.append(element("a:round" if style.get("stroke-linejoin") == "round" else "a:miter"))
        opacity_for(ln.solidFill[0], opacity * float(style.get("stroke-opacity", 1)))


def font_name(style):
    # These are the source's intended desktop fonts rather than a rasterized font.
    return "Georgia" if "Georgia" in style.get("font-family", "") else "Segoe UI"


def font_metrics(text, font_size, style):
    """Conservative text width and ascent, with a local open-font fallback."""
    family = font_name(style)
    try:
        font_file = subprocess.check_output(
            ["fc-match", "-f", "%{file}", family], text=True, stderr=subprocess.DEVNULL
        ).strip()
        font = ImageFont.truetype(font_file, round(font_size * 10))
        width = font.getlength(text) / 10
        ascent = font.getmetrics()[0] / 10
    except (FileNotFoundError, OSError, subprocess.CalledProcessError):
        width, ascent = len(text) * font_size * .58, font_size * .94
    width += max(0, len(text) - 1) * float(style.get("letter-spacing", 0))
    return width, ascent


class EditableSvg:
    def __init__(self, source: Path):
        self.source = source
        self.root = ET.parse(source, parser=ET.XMLParser(target=ET.TreeBuilder(insert_comments=True))).getroot()
        view_box = [float(n) for n in self.root.get("viewBox").split()]
        if view_box[:2] != [0, 0]:
            raise ValueError("The research figures require a zero-origin viewBox")
        self.width, self.height = view_box[2:]
        self.markers = {n.get("id"): n for n in self.root.iter(f"{{{SVG_NS}}}marker")}
        self.prs = Presentation()
        self.prs.slide_width = px(self.width)
        self.prs.slide_height = px(self.height)
        self.slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        self.slide.background.fill.solid()
        self.slide.background.fill.fore_color.rgb = RGBColor.from_string("F8F7F4")
        title = self.root.find(f"{{{SVG_NS}}}title")
        desc = self.root.find(f"{{{SVG_NS}}}desc")
        self.prs.core_properties.title = title.text if title is not None else source.stem
        self.prs.core_properties.subject = "Fully editable research figure"
        self.prs.core_properties.author = "Jingjing Wu"
        self.prs.core_properties.keywords = "environmental engineering, research, editable vector figure"
        self.prs.core_properties.comments = "Text and vectors are editable PowerPoint objects; no flattened figure images."
        self.slide.notes_slide.notes_text_frame.text = (
            (desc.text if desc is not None else "") + "\n\n"
            f"Source: images/{source.name}\n"
            "Every visible element is editable. Use Home > Arrange > Selection Pane to select objects, "
            "and Shape Format > Edit Shape > Edit Points for freeform icon paths. "
            "Colors and study scope follow the website figure."
        )
        self.text_records = []
        self.serial = 0

    def name(self, shape, label):
        self.serial += 1
        shape.name = f"{self.serial:02d} | {label[:110]}"
        return shape

    def freeform(self, target, commands, style, matrix, label):
        points = [p for _, pts in commands for p in pts]
        if not points:
            return
        transformed = [(kind, [transform_point(p, matrix) for p in pts]) for kind, pts in commands]
        points = [p for _, pts in transformed for p in pts]
        left, top = min(p.real for p in points), min(p.imag for p in points)
        width = max(.001, max(p.real for p in points) - left)
        height = max(.001, max(p.imag for p in points) - top)
        shape = self.name(target.add_shape(MSO_SHAPE.RECTANGLE, px(left), px(top), px(width), px(height)), label)
        sp_pr = shape._element.spPr
        sp_pr.remove(sp_pr.prstGeom)
        geometry = element("a:custGeom")
        for child in ("a:avLst", "a:gdLst", "a:ahLst", "a:cxnLst"):
            geometry.append(element(child))
        geometry.append(element("a:rect", l="l", t="t", r="r", b="b"))
        paths = element("a:pathLst")
        path = element("a:path", w=round(width * 1000), h=round(height * 1000),
                       fill="none" if style.get("fill") == "none" else "norm", stroke="1", extrusionOk="0")
        for kind, pts in transformed:
            command = element(f"a:{kind}")
            for p in pts:
                command.append(element("a:pt", x=round((p.real - left) * 1000), y=round((p.imag - top) * 1000)))
            path.append(command)
        paths.append(path)
        geometry.append(paths)
        sp_pr.insert(1, geometry)
        apply_appearance(shape, style, matrix)
        return shape

    def path(self, target, d, style, matrix, label):
        parsed = parse_path(d)
        commands = []
        for subpath in parsed.continuous_subpaths():
            commands.append(("moveTo", [subpath.start]))
            for segment in subpath:
                if isinstance(segment, Line):
                    commands.append(("lnTo", [segment.end]))
                elif isinstance(segment, QuadraticBezier):
                    commands.append(("quadBezTo", [segment.control, segment.end]))
                elif isinstance(segment, CubicBezier):
                    commands.append(("cubicBezTo", [segment.control1, segment.control2, segment.end]))
                elif isinstance(segment, Arc):
                    for curve in segment.as_cubic_curves(curves=max(1, math.ceil(abs(segment.delta) / 45))):
                        commands.append(("cubicBezTo", [curve.control1, curve.control2, curve.end]))
                else:
                    raise ValueError(f"Unknown path segment {type(segment).__name__}")
            if subpath.isclosed():
                commands.append(("close", []))
        self.freeform(target, commands, style, matrix, label)
        marker_url = style.get("marker-end")
        if marker_url and marker_url != "none":
            marker_id = re.fullmatch(r"url\(#([^)]*)\)", marker_url).group(1)
            marker = self.markers[marker_id]
            last = parsed[-1]
            tangent = last.derivative(1)
            if abs(tangent) == 0:
                tangent = last.end - last.start
            angle = math.atan2(tangent.imag, tangent.real)
            sw = float(style.get("stroke-width", 1)) if marker.get("markerUnits", "strokeWidth") == "strokeWidth" else 1
            vbox = [float(n) for n in marker.get("viewBox").split()]
            sx = float(marker.get("markerWidth", 3)) * sw / vbox[2]
            sy = float(marker.get("markerHeight", 3)) * sw / vbox[3]
            local = multiply((1, 0, 0, 1, last.end.real, last.end.imag),
                             (math.cos(angle), math.sin(angle), -math.sin(angle), math.cos(angle), 0, 0))
            local = multiply(local, (sx, 0, 0, sy, 0, 0))
            local = multiply(local, (1, 0, 0, 1, -float(marker.get("refX", 0)), -float(marker.get("refY", 0))))
            for child in marker:
                self.visit(child, target, {"fill": "#000000", "stroke": "none"}, multiply(matrix, local), label + " arrowhead")

    def text(self, target, node, style, matrix):
        text = "".join(node.itertext()).strip()
        if list(node):
            raise ValueError("SVG text spans need explicit handling rather than silently losing formatting")
        origin = transform_point(complex(float(node.get("x", 0)), float(node.get("y", 0))), matrix)
        x, y = origin.real, origin.imag
        scale = math.sqrt(abs(matrix[0] * matrix[3] - matrix[1] * matrix[2]))
        size = float(style.get("font-size", 16)) * scale
        measured, ascent = font_metrics(text, size, style)
        anchor = style.get("text-anchor", "start")
        # Generous one-line boxes accommodate font substitution; anchoring remains
        # exactly at the source x coordinate, so centered labels stay centered.
        width = measured * 1.04 + 6
        if anchor == "middle":
            width = min(width, 2 * min(x, self.width - x) - 2)
            left = x - width / 2
        elif anchor == "end":
            width = min(width, x - 1)
            left = x - width
        else:
            width = min(width, self.width - x - 1)
            left = x
        top = y - ascent
        height = size * 1.4
        # The last resource footnote sits close to the slide bottom. PPT's text
        # body can use available space without cropping the editable glyphs.
        height = min(height, self.height - top)
        box = self.name(target.add_textbox(px(left), px(top), px(width), px(height)), "Text · " + text)
        frame = box.text_frame
        frame.clear()
        frame.margin_left = frame.margin_top = frame.margin_right = frame.margin_bottom = 0
        frame.word_wrap = False
        frame.auto_size = MSO_AUTO_SIZE.NONE
        frame.vertical_anchor = MSO_ANCHOR.TOP
        frame._txBody.bodyPr.set("anchorCtr", "0")
        frame._txBody.bodyPr.set("horzOverflow", "overflow")
        frame._txBody.bodyPr.set("vertOverflow", "overflow")
        p = frame.paragraphs[0]
        p.alignment = {"middle": PP_ALIGN.CENTER, "end": PP_ALIGN.RIGHT}.get(anchor, PP_ALIGN.LEFT)
        p.space_before = p.space_after = Pt(0)
        p.line_spacing = Pt(size * .75)
        run = p.add_run()
        run.text = text
        run.font.name = font_name(style)
        run.font.size = Pt(size * .75)
        run.font.bold = style.get("font-weight") in ("bold", "600", "700")
        run.font.italic = style.get("font-style") == "italic"
        run.font.color.rgb = RGBColor.from_string(color_value(style.get("fill", "#000000")))
        rpr = run._r.get_or_add_rPr()
        rpr.set("lang", "en-US")
        if style.get("letter-spacing"):
            rpr.set("spc", str(round(float(style["letter-spacing"]) * .75 * 100)))
        # Keep font selection explicit for each script if opened on macOS/Windows.
        rpr.append(element("a:ea", typeface=font_name(style)))
        rpr.append(element("a:cs", typeface=font_name(style)))
        self.text_records.append((text, left, top, width, height, measured))

    def visit(self, node, target, inherited, parent_matrix, label="Figure"):
        if node.tag is ET.Comment:
            return
        tag = node.tag.rsplit("}", 1)[-1]
        if tag in {"defs", "title", "desc", "metadata"}:
            return
        style = style_for(node, inherited)
        matrix = multiply(parent_matrix, matrix_for(node.get("transform", "")))
        label = node.get("id", label)
        if tag in {"svg", "g"}:
            current_label = label
            group = target if tag == "svg" else self.name(target.add_group_shape(), label)
            for child in node:
                if child.tag is ET.Comment:
                    current_label = re.sub(r"\s+", " ", child.text or "").strip().rstrip(".")
                else:
                    self.visit(child, group.shapes if tag != "svg" else target, style, matrix, current_label)
            return
        if tag == "text":
            self.text(target, node, style, matrix)
            return
        if tag == "path":
            self.path(target, node.get("d"), style, matrix, label + " · path")
            return
        if tag == "line":
            d = f"M{node.get('x1', 0)} {node.get('y1', 0)} L{node.get('x2', 0)} {node.get('y2', 0)}"
            self.path(target, d, style, matrix, label + " · line")
            return
        if tag in {"polyline", "polygon"}:
            points = [float(n) for n in re.findall(r"[-+]?(?:\d*\.\d+|\d+\.?\d*)", node.get("points", ""))]
            commands = [("moveTo" if n == 0 else "lnTo", [complex(*points[n:n+2])]) for n in range(0, len(points), 2)]
            if tag == "polygon":
                commands.append(("close", []))
            self.freeform(target, commands, style, matrix, label + f" · {tag}")
            return
        if tag in {"circle", "ellipse", "rect"}:
            # These research figures use axis-aligned primitives. Keep them as
            # ordinary PowerPoint shapes instead of converting to point paths.
            if abs(matrix[1]) > 1e-9 or abs(matrix[2]) > 1e-9:
                raise ValueError("Rotated primitives need explicit handling")
            if tag in {"circle", "ellipse"}:
                cx, cy = float(node.get("cx", 0)), float(node.get("cy", 0))
                rx = float(node.get("r", node.get("rx", 0)))
                ry = float(node.get("r", node.get("ry", 0)))
                p1 = transform_point(complex(cx-rx, cy-ry), matrix)
                p2 = transform_point(complex(cx+rx, cy+ry), matrix)
                shape_type = MSO_SHAPE.OVAL
            else:
                x, y = float(node.get("x", 0)), float(node.get("y", 0))
                w, h = float(node.get("width")), float(node.get("height"))
                p1, p2 = transform_point(complex(x, y), matrix), transform_point(complex(x+w, y+h), matrix)
                shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if node.get("rx") else MSO_SHAPE.RECTANGLE
            shape = self.name(target.add_shape(shape_type, px(min(p1.real, p2.real)), px(min(p1.imag, p2.imag)),
                                              px(abs(p2.real-p1.real)), px(abs(p2.imag-p1.imag))), label + f" · {tag}")
            if tag == "rect" and node.get("rx"):
                shape.adjustments[0] = min(.5, float(node.get("rx")) / min(w, h))
            apply_appearance(shape, style, matrix)
            return
        raise ValueError(f"Unsupported visible SVG element: {tag}")

    def export(self, destination: Path):
        self.visit(self.root, self.slide.shapes, {"fill": "#000000", "stroke": "none"}, IDENTITY)
        # Group bounds must be recalculated after nested child objects exist.
        def update_groups(shapes):
            for shape in shapes:
                if shape.shape_type == 6:
                    update_groups(shape.shapes)
                    shape._element.recalculate_extents()
        update_groups(self.slide.shapes)
        destination.parent.mkdir(parents=True, exist_ok=True)
        self.prs.save(destination)
        return self.validate(destination)

    def validate(self, destination):
        reopened = Presentation(destination)
        assert len(reopened.slides) == 1, "Expected exactly one slide"
        with ZipFile(destination) as package:
            assert not any(n.startswith("ppt/media/") for n in package.namelist()), "Unexpected embedded image"
            slide_xml = etree.fromstring(package.read("ppt/slides/slide1.xml"))
            ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main",
                  "p": "http://schemas.openxmlformats.org/presentationml/2006/main"}
            actual = slide_xml.xpath("//a:t/text()", namespaces=ns)
            expected = [text for text, *_ in self.text_records]
            assert actual == expected, "Text changed during export"
            assert not slide_xml.xpath("//p:pic", namespaces=ns), "Raster fallback found"
            counts = {
                "textboxes": len(actual),
                "native_primitives": len(slide_xml.xpath("//p:sp[not(p:nvSpPr/p:cNvSpPr/@txBox='1')]/p:spPr/a:prstGeom", namespaces=ns)),
                "editable_freeforms": len(slide_xml.xpath("//a:custGeom", namespaces=ns)),
                "groups": len(slide_xml.xpath("//p:grpSp", namespaces=ns)),
                "embedded_images": 0,
            }
        bounds = []
        for text, left, top, width, height, measured in self.text_records:
            if min(left, top) < -.1 or left + width > self.width + .1 or top + height > self.height + .1:
                bounds.append(text)
            if measured > width + .1:
                raise AssertionError(f"Text box narrower than conservative glyph estimate: {text}")
        assert not bounds, f"Text boxes outside slide: {bounds}"
        return counts


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--only", choices=JOBS)
    args = parser.parse_args()
    for key, (source, destination) in JOBS.items():
        if args.only and args.only != key:
            continue
        source, destination = REPO / "images" / source, REPO / "files" / destination
        counts = EditableSvg(source).export(destination)
        print(f"{destination.relative_to(REPO)}: " + ", ".join(f"{v} {k}" for k, v in counts.items()))


if __name__ == "__main__":
    main()
