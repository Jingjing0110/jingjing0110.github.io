(() => {
  const root = document.documentElement;
  const button = document.querySelector('.theme-toggle');
  if (!button) return;
  const update = () => {
    const dark = root.dataset.theme === 'dark';
    button.setAttribute('aria-pressed', String(dark));
    button.setAttribute('aria-label', `Switch to ${dark ? 'light' : 'dark'} theme`);
    document.querySelector('meta[name="theme-color"]')?.setAttribute('content', dark ? '#202b33' : '#f7f7f3');
  };
  button.hidden = false;
  update();
  button.addEventListener('click', () => {
    const theme = root.dataset.theme === 'dark' ? 'light' : 'dark';
    root.dataset.theme = theme;
    try { localStorage.setItem('theme', theme); } catch (_) { /* Theme still works without storage. */ }
    update();
  });
})();

// Retain the template's optional Markdown charts without loading chart tools on academic pages.
window.addEventListener('load', () => {
  if (!window.Plotly) return;
  document.querySelectorAll('pre > code.language-plotly').forEach((code) => {
    try {
      const spec = JSON.parse(code.textContent);
      const chart = document.createElement('div');
      code.parentElement.after(chart);
      Plotly.newPlot(chart, spec.data, spec.layout || {}, {responsive: true}).then(() => {
        code.parentElement.hidden = true;
      });
    } catch (_) { /* Leave readable source visible if a chart cannot be rendered. */ }
  });
});
