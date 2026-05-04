# xlsx-preview

Browser-side XLSX → HTML renderer. Sibling project to
[docxjs](../docxjs).

Current scope is intentionally minimal: opens an `.xlsx`, reads the first
sheet, and renders the cells as an HTML table with column-letter headers
and a row-number gutter. Plain values only — styles, number formats,
formulas, merged cells, and images are follow-up work.

## Usage

```html
<script src="https://unpkg.com/jszip/dist/jszip.min.js"></script>
<script src="dist/xlsx-preview.js"></script>
<script>
  fileInput.addEventListener('change', async () => {
    await xlsx.renderAsync(fileInput.files[0], document.getElementById('out'));
  });
</script>
```

API surface is `parseAsync(data, options)`, `renderWorkbook(wb, options)`,
and `renderAsync(data, bodyContainer, styleContainer?, options)`.

## Develop

```bash
npm install
npm run build
npm run test:render   # jsdom harness against the `basic` fixture
npm run dev           # static server at :8766
```

Apache-2.0.
