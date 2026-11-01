#!/usr/bin/env python3
"""Copy a built Jekyll site into an ignored preview that opens directly from disk.

Usage: python3 scripts/build_portable_preview.py /path/to/jekyll-output
This never modifies Jekyll source pages or publishes the site.
"""
from pathlib import Path
import argparse
import os
import re
import shutil
from urllib.parse import unquote, urlsplit

ROOT = Path(__file__).resolve().parents[1]
DESTINATION = ROOT / 'local' / 'homepage-preview' / 'site'


def local_url(value, page):
    if not value.startswith('/') or value.startswith('//'):
        return value
    parsed = urlsplit(value)
    target = DESTINATION / unquote(parsed.path.lstrip('/'))
    if target.is_dir():
        target /= 'index.html'
    relative = os.path.relpath(target, page.parent).replace(os.sep, '/')
    return relative + ('?' + parsed.query if parsed.query else '') + ('#' + parsed.fragment if parsed.fragment else '')


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('built_site', type=Path)
    args = parser.parse_args()
    source = args.built_site.resolve()
    if not (source / 'index.html').is_file() or source == ROOT or source == DESTINATION:
        parser.error('Pass a built Jekyll output folder with index.html, outside the source/preview folders.')
    if DESTINATION.is_symlink():
        parser.error("The generated preview folder must not be a symbolic link.")
    if DESTINATION.exists():
        shutil.rmtree(DESTINATION)
    DESTINATION.parent.mkdir(parents=True, exist_ok=True)
    shutil.copytree(source, DESTINATION)
    count = 0
    pattern = re.compile(r'(\b(?:href|src|action|srcset)\s*=\s*)([\'\"])(.*?)(\2)', re.I | re.S)
    for page in DESTINATION.rglob('*.html'):
        html = page.read_text()
        def replace(match):
            return match[1] + match[2] + local_url(match[3], page) + match[4]
        html = pattern.sub(replace, html)
        html = html.replace('</head>', '<meta name="robots" content="noindex,nofollow"></head>', 1)
        page.write_text(html)
        count += 1
    print(f'Created {count} portable preview pages: {DESTINATION / "index.html"}')


if __name__ == '__main__':
    main()
