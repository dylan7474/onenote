# OneNote Compatibility Review

Scope: interoperability with Microsoft OneNote — reading and writing what real
OneNote clients and the Microsoft Graph OneNote API actually produce. The
`README.md` roadmap already covers broad feature parity; this document is
narrower and every point is tied to current code.

Reviewed: `index.html` (~2,300 lines), `server.js`, `deploy.sh`, `README.md`.

---

## 1. The only supported interop path: Microsoft Graph OneNote API

There is currently no connection to Microsoft 365, and the README frames
`.one`/`.onepkg` as the compatibility story. The realistic, supported route is
**Graph**: `GET /me/onenote/notebooks`, `.../sections`, `.../pages`,
`POST .../pages` with `multipart/form-data` HTML.

- Add an MSAL browser sign-in and a "Connect OneNote" action alongside the
  existing Import modal (`index.html:171`).
- Import: pull the notebook/section/page tree from Graph, and fetch each page's
  `content` (`?includeIDs=true`) — HTML this app already ingests.
- Export/push: `POST` a page as Graph's "supported input HTML" subset.
- `server.js` needs a token-exchange/proxy endpoint (Graph blocks CORS for some
  calls, and the client secret must not sit in the browser).

This does more for OneNote compatibility than everything below combined. The
rest is about not losing data on the file-based path in the meantime.

## 2. HTML import silently drops the most common content

`parseHtmlImport` (`index.html:2105`) and `parseZipImport` (`index.html:2144`)
keep `doc.body.innerHTML` verbatim in a single block and only special-case
`<object data-attachment>`.

- **Inline images are lost.** OneNote HTML export emits
  `<img data-fullres-src="..." data-render-src="..." src="...">`, and in a ZIP
  those point at sibling files. Nothing resolves them, so images render broken
  or not at all. `extractOneNoteAttachments` already has a `resolveFile`
  callback for the ZIP path — extend it to walk `img[src]`, pull bytes from the
  ZIP (or the `data:` URL), and rewrite `src` to a data URL or a stored
  attachment. Biggest visible fidelity loss today.
  _Status: done in Phase 1 (#25) — `inlineImages()` embeds ZIP-resident and
  `data:` images; remote/missing refs are left untouched. Plain-HTML relative
  refs still can't be resolved (no sibling files)._
- **Positioned outlines are flattened.** OneNote wraps each note container in
  `<div style="position:absolute;left:Npx;top:Mpx;width:Wpx">`. The data model
  already has per-block `x`/`y` (`index.html:2131`), but the importer forces
  everything into one block at `0,0`. Parse each top-level absolutely-positioned
  `<div>` into its own block with `x`/`y`/`width`.
- **Timestamps are fabricated.** `createdAt`/`updatedAt` are set to
  `new Date()` (`index.html:2122-2123`). OneNote HTML carries
  `<meta name="created">` / `<meta name="lastModified">`; Graph gives
  `createdDateTime`/`lastModifiedDateTime`. Read them.
- **Page title.** `doc.title` is right for single-page export, but OneNote's
  "export whole section as one HTML file" produces one document with many
  `<h1>`-delimited pages. Split on the page separators rather than making one
  giant page.

## 3. No HTML export = no round-trip back into OneNote

`exportCurrentNotebook` (`index.html:2259`) only emits this app's own JSON.
There is no way to get content into real OneNote. Add:

- **"Copy page as HTML"** to the clipboard using `text/html` — desktop OneNote
  accepts pasted HTML directly, no API needed.
- **"Export page/section as OneNote-compatible HTML"** — the Graph
  supported-input subset: `<html><head><title>` plus a body with `data-tag`
  attributes, tables, `<img>` as data URLs, `<object data-attachment>` for
  files. Same shape the importer reads, so it also gives lossless self
  round-trip.

## 4. Align the data model with Graph's schema

Small renames now avoid a painful adapter later:

| This project | Graph |
| --- | --- |
| `notebook.name`, `section.name` (`index.html:855`) | `displayName` |
| `page.title` | `title` |
| `block.content` | page-level `content` (single HTML doc) |
| `page.level` 0–2 (`index.html:2031`) | no subpage nesting via API; subpages are `<h1>`/indent level in content |

The last row matters: the 3-level `level` model maps to nothing Graph exposes,
and OneNote's own ZIP export represents subpages by **filename** (`Page.html`,
`Page 1.html`, `Page 2.html` for a subpage group), **not nested folders**. The
ZIP heuristic `filename.split('/').length - 1 - commonDepth`
(`index.html:2175`) mis-nests almost every real OneNote export, which uses a
flat `Notebook/Section/Page.html` layout. Test against an actual export and
rework it.

## 5. Tag / checkbox model mismatch

Checkboxes are `<input type="checkbox" onclick="this.setAttribute('checked', this.checked)">`
(`index.html:740`, `index.html:2235`). OneNote and Graph use
`<p data-tag="to-do">` / `data-tag="to-do:completed"`, and likewise
`data-tag="important"`, `data-tag="question"`, `data-tag="critical"`, etc. —
which also line up with the three page tags (Important / Question / Idea,
`index.html:331`).

- Import: map `data-tag` on paragraphs to the checkbox UI and tag chips.
- Export: emit `data-tag` instead of `<input>`.
- Bonus: removes the inline `onclick` handlers, which are an XSS vector and
  won't survive sanitization.

## 6. Sanitize before ingesting real OneNote HTML

`workspace.innerHTML = ...block.content` (`index.html:1088-1094`) and
`template.innerHTML = block.content` (`index.html:1128`) render imported HTML
with no filtering. "Import an HTML file someone sent me" is the untrusted path.
Run imported and `contenteditable` output through DOMPurify with an allowlist
that **keeps** `style`, `data-tag`, `data-id`, `data-render-src`,
`width`/`height`, and table attributes — otherwise sanitization strips the
OneNote semantics too. Pin/self-host the library (JSZip is already self-hosted
via cdnjs; `unpkg` for Lucide at `index.html:30` is a supply-chain soft spot).

## 7. `.one` / `.onepkg` — README framing is inaccurate

`.one`/`.onepkg` is widely assumed to be an undocumented binary blob. It is
not: Microsoft publishes **[MS-ONESTORE]** (the revision-store file) and
**[MS-ONE]** (the property set / content schema) as open specifications, and
`.onepkg` is a CAB archive of `.one` files. A **read-only** importer is
feasible (parse CAB → parse ONESTORE object space → walk [MS-ONE] properties →
emit the block model), though it is a multi-week effort and the output still
has to be lowered to HTML. Recommend Graph first; the README should describe it
as a "documented but complex binary format" rather than simply unsupported. An
intermediate `.onepkg` CAB-extraction step that reports "N .one files found,
native parsing not yet supported" would be honest.

_Status: README updated with the [MS-ONESTORE]/[MS-ONE] pointers in Phase 0._

## 8. Section colors

`SECTION_COLORS` is an arbitrary hex list assigned at random
(`index.html:485`, `index.html:2196`). OneNote has a fixed set of ~16 named
section colors. Map to/from that named palette on import/export so a round-trip
preserves the user's section colors instead of randomizing them.

## 9. Smaller importer bugs

- `decodeDataUrl` non-base64 branch (`index.html:2049`):
  `btoa(unescape(encodeURIComponent(decodeURIComponent(match[3]))))`
  double-decodes and corrupts any `data:text/...` URL containing `%`. Use
  `TextEncoder` plus proper base64.
- ZIP attachment resolution (`index.html:2164`) only tries
  `directory + cleanSource` and bare `cleanSource`; OneNote often stores assets
  in a sibling `*_files/` or `attachments/` folder — add a basename fallback
  search across `zip.files`.
- `extractOneNoteAttachments` size calc `payload.data.length * 3 / 4`
  (`index.html:2063`) ignores base64 padding — off by up to 2 bytes; harmless
  but `formatBytes` then displays it.
- Graph-sourced `<object data-attachment>` has
  `data="https://graph.microsoft.com/.../resources/..."` needing an auth
  header — `resolveFile` cannot fetch it. Detect the Graph host and surface
  "attachment requires sign-in" rather than silently producing a chip with no
  payload.

---

## Priority order

1. Sanitize imported/edited HTML (safety gate for everything else).
2. Resolve `<img>` on HTML/ZIP import (biggest fidelity win; importer is already
   80% there).
3. `data-tag` import/export, and drop inline `onclick`.
4. HTML/clipboard export (first real path back into OneNote).
5. Fix/verify the ZIP subpage heuristic against a real export.
6. Microsoft Graph connect (largest effort, but the actual compatibility
   answer).
