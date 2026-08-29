# Test fixtures

**These are synthetic.** They are hand-written to match the *shape* of Microsoft
OneNote's HTML/ZIP export and this app's own JSON export, so the importers can be
exercised without a Microsoft 365 account. They are not captured from a real
OneNote client.

When real exports become available, drop them in here alongside these and point
new tests at them — the synthetic files can stay as minimal, readable cases.

| File | Represents | Deliberately contains |
| --- | --- | --- |
| `onenote-page.html` | A single-page OneNote "Export as HTML" file | `<meta name="created">`, one `position:absolute` outline, `data-tag` paragraphs, a table, an inline `<img>` data URL, an `<object data-attachment>` data URL, **and** hostile markup outside the outline (`<script>`, `onclick`, `onerror`) that must not survive import |
| `onenote-multi-outline.html` | A page with several OneNote outlines | Two `position:absolute` outlines (one in `px`, one in `pt`), no stray body content |
| `plain.html` | An HTML file with no positioned outlines | Just a heading and a paragraph — the single-block fallback path |
| `onenote-tags.html` | A page using OneNote `data-tag` markup | `important`, `question`, a comma list, an unknown tag, `to-do` and `to-do:completed` |
| `onenote-graph-object.html` | A page whose attachment lives behind Microsoft Graph | An `<object data-attachment>` whose `data` is a `graph.microsoft.com` resource URL (can't be fetched at import) |
| `onenote-section.html` | A whole-section "Export as HTML" file | Three `<h1>`-headed pages in one document, an inline `<object>` on the first, `data-tag` markup, `<meta name="created">` |
| `webapp-notebook.json` | This app's own "Export notebook" JSON (full-backup shape) | `parentPageId` chains for subpage-level inference; a block with `<script>`/`onclick` to prove JSON import sanitizes |
| `zip-src/` | Unpacked tree that `import.test.js` zips at runtime to stand in for an "OneNote ZIP" | A `Subpages/` folder, a `Report.html` + `Report 1.html` / `Report 2.html` subpage group, a sibling-resolved `<object data-attachment>`, an `<img>` referenced by a path that only resolves via basename fallback, and one `<img>` with no matching file |

`zip-src/**/*.png` is a 1×1 transparent PNG.
