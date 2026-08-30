# OneNote Web

A lightweight, self-hosted notebook inspired by Microsoft OneNote. It runs in a
browser and automatically persists the complete notebook state to the server.

> [!IMPORTANT]
> This is an independent project, not a Microsoft product. “OneNote” is a
> Microsoft trademark. The application does **not** currently read native
> `.one` files or connect to Microsoft 365/OneDrive. “Parity” below means a
> similar user workflow, not file-format or service compatibility.

## Run locally

The application has no build step and uses only Node.js built-in modules on the
server:

```bash
PORT=3020 DATA_FILE=./data/state.json node server.js
```

Open <http://localhost:3020>. Do not open `index.html` directly: the application
uses the server's `/api/state` endpoint to save and restore data, and loads its
browser libraries from `/vendor/`. `DATA_FILE` defaults to `./data/state.json`,
and request bodies are limited to 50 MiB.

DOMPurify, Lucide and JSZip are checked into `vendor/` at pinned versions and
served by `server.js`; imported and edited HTML is sanitised through that
DOMPurify build before it is stored or rendered. See `vendor/README.md` for
provenance. Tailwind and the Inter web font are still loaded from a CDN.

## Tests

```bash
npm ci        # installs jsdom, the only dev dependency
npm test      # node --test
```

The suite (`test/`) loads the inline application script into jsdom and pins the
current behaviour of the importers, exporter, sanitiser and notebook/section/page
operations, so later refactors are caught. Cases that document a known gap from
`REVIEW.md` are marked `{ todo: true }` and become real assertions when that gap
is closed. Fixtures under `test/fixtures/` are synthetic — see the README there.
CI runs `npm test` on pushes to `main` and on pull requests.

## Using file attachments

Select a page and place the caret in a note, then choose **Insert → File
attachment**, use **Attach File** on the editor toolbar, or drop one or more
files onto the page. Each accepted file is inserted at the caret; selecting its
attachment chip opens or downloads it. If there is no usable caret, the file is
inserted at the end of the first note container.

Attachments are limited to 10 MiB per file and 30 MiB in total per page. They
are Base64-encoded inside the notebook state and JSON exports rather than stored
as separate files, so attachments increase the size of `DATA_FILE`, exports,
and every save request. Back up the data file or Docker volume before importing
large archives.

HTML imports can recover an attachment only when its OneNote `<object
data-attachment>` contains an embedded data URL. ZIP imports can also resolve a
referenced payload stored in the ZIP alongside the HTML page, and inline any
`<img>` whose file is present in the ZIP as a data URL (remote and missing image
references are left as-is). Each OneNote `position:absolute` outline is imported
as its own note block (keeping its canvas coordinates), and page
created/modified times are taken from the document's `<meta name="created">` /
`<meta name="lastModified">` when present. A single HTML file that holds a whole
section (several `<h1>`-headed pages) is split into one page each, and a ZIP
whose subpages are named `Foo 1.html` / `Foo 2.html` alongside `Foo.html` keeps
that nesting. `data-tag` markup is mapped on
import: `to-do` becomes a checkbox, and the other tag values (`important`,
`question`, …) become page tag chips. An `<object>` attachment that points at a
remote URL (e.g. a Microsoft Graph resource) is shown as an "unavailable"
placeholder, since loading it needs an authenticated request. Native `.one` and
`.onepkg` files are **inspected but not parsed**: dropping a `.onepkg` reads its
MS-CAB directory and lists the `.one` sections it contains (with sizes), a
`.one` file is recognised by name, and either way the Import dialog explains
that native binary parsing is unavailable and points at OneNote's own HTML/ZIP
export or the Microsoft account connection. The formats themselves are
documented by Microsoft ([MS-ONESTORE] for the revision-store file, [MS-ONE] for
the content schema; `.onepkg` is a CAB archive of `.one` files), so a full
read-only importer is feasible but out of scope here — use Microsoft 365 / the
Graph OneNote API for live interoperability.

[MS-ONESTORE]: https://learn.microsoft.com/openspecs/office_file_formats/ms-onestore/
[MS-ONE]: https://learn.microsoft.com/openspecs/office_file_formats/ms-one/

**File → Export page as HTML** writes the active page in the "supported input
HTML" shape the Microsoft Graph OneNote API accepts: a `<head>` with
created/modified meta, one absolutely-positioned `<div>` per note block,
checkboxes as `<p data-tag="to-do">`, and attachments as `<object
data-attachment>` with embedded data URLs. **Export section as HTML** writes the
whole section as one document — an `<h1>` per page (with `data-level` for
subpages) — and **Copy page as HTML** puts the same page markup on the clipboard
(`text/html`) to paste straight into desktop OneNote. Re-importing an exported
file reproduces the page(s), subpage nesting included. (Page-level tag chips
other than the `to-do` checkbox are not carried — OneNote's content HTML has no
page-tag slot.)

## Connect to Microsoft OneNote (optional)

The server can talk to the **Microsoft Graph OneNote API** on behalf of a
signed-in user. It is **off unless `GRAPH_CLIENT_ID` is set** — a default deploy
is unaffected and needs no Microsoft account.

To enable it you need a free [Entra ID app
registration](https://learn.microsoft.com/graph/auth-register-app-v2) (no paid
Microsoft 365 subscription — a personal Microsoft account works):

- **Redirect URI:** `http://localhost:3020/api/graph/callback` (match your host/port)
- **Delegated permissions:** `Notes.ReadWrite`, `Notes.Create`, `User.Read`, `offline_access`
- Supported account types: personal + work/school

Then run the server with:

```bash
GRAPH_CLIENT_ID=<app-client-id> \
GRAPH_CLIENT_SECRET=<secret>            # optional; PKCE is used if omitted \
GRAPH_TENANT=common                     # or consumers | organizations | <tenant-id> \
GRAPH_REDIRECT_URI=http://localhost:3020/api/graph/callback \
node server.js
```

The server exposes `/api/graph/config`, `/api/graph/login` (→ Microsoft
sign-in), `/api/graph/callback`, `/api/graph/logout`, and a
`/api/graph/v1.0/*` passthrough to `https://graph.microsoft.com/v1.0/*` that
attaches the user's bearer token (refreshed on expiry). Tokens live only in the
server's memory, keyed by the `onenote_gsid` session cookie.

Once enabled, the **Import OneNote** dialog shows a **Microsoft OneNote (live)**
panel: connect, pick a notebook, and **Import selected notebook** pulls its
sections and pages (following `@odata.nextLink` pagination and backing off on
`429`) straight into the app. The same panel can **send the open page to a
OneNote section** — it POSTs the page's OneNote-compatible HTML and opens the
new page in OneNote on the web.

## Deploy with Docker

```bash
./deploy.sh 3020
```

The deployment stores application state in the `onenote-data` Docker volume, so
rebuilding or replacing the application container does not delete saved notes.

## Feature status

This inventory was checked against the implementation on **7 August 2026**.
Microsoft offers several OneNote clients whose feature sets differ; the target
for this comparison is the common modern OneNote experience rather than a
specific legacy desktop release.

### Available now

- Notebook → colored section → page hierarchy, with create/delete/navigation,
  two levels of subpages, and collapsible subpage groups. Section colors come
  from a fixed 16-colour named palette (new sections cycle through it; imported
  colors snap to the nearest name).
- Vertical or horizontal section navigation and light/dark themes.
- Page title, created/updated timestamps, page filtering, and the Important,
  Question, and Idea page tags.
- Multiple editable note containers with headings, bold, italic, underline,
  strikethrough, font color, highlighting, bullets, numbering, and checkboxes.
- Inline file attachments via the Insert menu, editor toolbar, or drag and drop.
  Files are inserted at the text caret, persist with the page, and can be
  downloaded/launched in place (up to 10 MiB each and 30 MiB per page). OneNote
  HTML `<object data-attachment>` elements are retained at their original body
  position when their payload is embedded or present in an imported ZIP.
- Table insertion (up to 6 × 6 from the picker), row/column editing, rectangular
  multi-cell selection, merge/split, drag resizing, even distribution, sorting,
  header and banded-row styles, cell shading, keyboard navigation, and deletion.
- A pen layer with selectable color and width, saved with the page.
- JSON notebook export and per-page OneNote-compatible HTML export; JSON,
  HTML/HTM, and ZIP-of-HTML import, preserving supported OneNote subpage levels,
  outline geometry, `<meta>` timestamps, `data-tag` markup, and inlined images.
- Debounced server-side JSON persistence with a versioned state schema:
  `migrateState()` upgrades older saved data on load, and the server refuses to
  overwrite the data file with a state stamped newer than it understands. Also a
  one-time migration from the older browser-local state.

### Comparison with Microsoft OneNote

Legend: **Available** = usable today; **Partial** = a narrower local
implementation exists; **Missing** = no implementation yet.

| Capability | This project | Microsoft OneNote comparison / remaining gap |
| --- | --- | --- |
| Notebooks, sections, pages | **Available** | Core hierarchy, two-level page nesting, hierarchy-aware import, and collapsing/expanding subpage groups exist. Add section groups, reorder/move/copy, rename, colors, recycle bins, and closed-notebook management. |
| Free-form page canvas | **Partial** | Notes can contain multiple blocks and ink, but blocks are vertically stacked. Add true click-anywhere positioning, drag/resize, z-order, canvas zoom, and paper size/background/rule-line controls. |
| Rich text and lists | **Partial** | Basic formatting, headings, lists, highlighting, and checkboxes exist. Add font family/size, styles, indentation, alignment, line spacing, clear formatting, format painter, symbols, equations, code formatting, and robust paste. |
| Tables | **Partial** | Creation, resizing, Shift-click rectangular selection, merge/split, distribution, sorting, headers, banding, shading, row/column operations, and Tab/Shift+Tab navigation exist. Remaining OneNote gaps include nested/irregular merge edge cases, conversion to/from text or Excel, formulas, advanced style galleries, repeated headers, and a complete assistive-technology/browser test matrix. |
| Tags and tasks | **Partial** | Three page-level labels and inline checkboxes exist. OneNote applies built-in/custom tags to individual content and can summarize/find tags; add that model plus Outlook task integration where supported. |
| Search | **Partial** | The “global” search currently filters only titles/tags in the active section. Add indexed full-text and OCR search across all notebooks, result snippets, scopes, filters, recent searches, and tag search. |
| Ink and Draw | **Partial** | One bitmap pen layer exists. Add stroke/vector storage, eraser/lasso, selection/transform, highlighters, pressure/touch support, shapes, ruler, ink replay, and ink-to-shape/text/math. |
| Insert content | **Partial** | File attachments can be added at the caret, persisted, and downloaded/launched; OneNote HTML attachment objects retain their body position on import. Add inline images, attachment printouts/previews, camera/scans, links, audio/video recordings, online video, date/time, equations/symbols, stickers, and reusable page templates. |
| Capture and integrations | **Partial** | With `GRAPH_CLIENT_ID` set, sign in from the Import dialog to pull a Microsoft 365 / personal notebook and push individual pages back via the Graph OneNote API. Still missing: a web clipper/share target, email-to-note workflow, meeting details, and Outlook tasks. |
| Collaboration and sync | **Missing** | State is one server-wide JSON document. Add accounts, private notebooks, invitations/links, permissions, real-time coauthoring, presence, comments/@mentions, conflict handling, offline cache, and multi-device sync. |
| History and recovery | **Missing** | The History menu is informational only. Add undo/redo across editing, autosaved revisions, page versions/diff/restore, author attribution, recent edits, deleted-notes recycle bin, and backup/restore. |
| Import/export/print | **Partial** | Project JSON and HTML-based imports exist; “OneNote ZIP” means ZIP files containing HTML, not native OneNote packages. Add sanitized, asset-aware import; PDF/HTML/Markdown export; print/preview; and document native `.one`/`.onepkg` limitations clearly. |
| Accessibility and language | **Partial** | Semantic controls are limited and no audit has been completed. Add complete keyboard navigation, focus management, screen-reader labels, contrast/reflow testing, accessibility checker, spell/grammar checking, translation, dictation, and Immersive Reader-style reading tools. |
| Security and administration | **Missing** | Add authentication/authorization, encrypted transport guidance, password-protected sections, per-user storage, audit logs, retention, quotas, validated uploads, HTML sanitization, CSP, CSRF protection, and rate limiting. |
| Cross-platform experience | **Partial** | It is browser-based but desktop-oriented. Add responsive/mobile layouts, touch gestures, installable PWA support, offline editing, and tested browser/device compatibility. |

## Roadmap to closer parity

The order below deliberately puts data safety and editor foundations ahead of
surface-level ribbon imitation. A visual clone without reliable storage,
security, or accessible interaction would not be a useful OneNote replacement.

### Phase 0 — Define and protect the contract

- [ ] Add automated tests for notebook CRUD, editor serialization, import/export,
  drawing persistence, API validation, and atomic saves.
- [ ] Version the state schema and implement migrations, validation, stable IDs,
  backup/restore, and recovery from a corrupt or interrupted write.
- [ ] Split the single HTML file into testable editor, data, persistence, import,
  and UI modules; pin or self-host third-party browser dependencies.
- [ ] Sanitize edited/imported HTML and attachments; add CSP, security headers,
  CSRF defenses, upload type/size checks, and safe filenames.
- [ ] Correct UI claims: distinguish local HTML/JSON interchange from native
  OneNote `.one`/`.onepkg` compatibility and label search by its actual scope.

**Exit criteria:** repeatable tests protect all current features; old data
migrates without loss; malformed or hostile imports cannot execute scripts.

### Phase 1 — Match the core notebook and editor workflow

- [ ] Add rename, drag-to-reorder, move/copy, section groups,
  duplicate, archive, and recycle-bin restore for every hierarchy level.
- [ ] Implement truly free-form, draggable/resizable note containers with
  click-anywhere creation, selection, layering, zoom, and page backgrounds.
- [ ] Replace deprecated browser editing commands with a structured editor model
  supporting undo/redo, keyboard shortcuts, robust paste, links, images,
  attachments, advanced text formatting, and mature table editing.
- [ ] Store ink as editable strokes and add eraser, lasso, highlighter, shapes,
  pressure input, and selection transforms.
- [ ] Build indexed full-content search across notebooks, including scoped
  results, snippets, tag queries, and (after image support) OCR text.

**Exit criteria:** everyday notebook organization, typing, pasting, attaching,
drawing, finding, undoing, deleting, and restoring work without data loss.

### Phase 2 — Multi-user sync, sharing, and history

- [ ] Move from one JSON blob to a transactional per-user/per-object data model
  with attachment/blob storage and incremental synchronization.
- [ ] Add authentication, notebook ownership, share invitations/links, read/edit
  roles, revocation, password-protected sections, and an audit trail.
- [ ] Add real-time coauthoring with presence, cursors, attribution, comments,
  @mentions, deterministic conflict resolution, and reconnect handling.
- [ ] Add immutable page revisions, diffs, restore, author history, recent edits,
  deleted-item retention, and scheduled backups.
- [ ] Add a service worker, local operation queue, offline status, and tested
  offline-to-online conflict recovery.

**Exit criteria:** two users can safely edit and recover a shared notebook across
devices, including disconnects and concurrent changes.

### Phase 3 — Capture, intelligence, and interoperability

- [ ] Add audio/video capture with timestamped notes, dictation, transcription,
  image/document OCR, equations, ink conversion, and accessibility reading tools.
- [ ] Add camera/document scan, a browser clipper/share target, templates, meeting
  details, task workflows, and optional calendar/email integrations.
- [ ] Add high-fidelity print/PDF plus HTML/Markdown export and asset-preserving
  round trips; investigate native OneNote conversion only through documented,
  legally supportable APIs or user-controlled desktop tooling.
- [ ] Offer opt-in Microsoft Graph integration for Microsoft 365 notebooks,
  respecting Graph permissions, throttling, supported content, and service terms.
- [ ] Complete WCAG-oriented keyboard, screen-reader, zoom/reflow, contrast,
  reduced-motion, touch, localization, and mobile/PWA test matrices.

**Exit criteria:** capture, retrieval, accessibility, and interchange cover the
major modern OneNote workflows, with every unsupported edge case documented.

## Near-term candidate issues

These are intentionally small slices that can be implemented and reviewed
independently:

1. Add schema versioning, runtime validation, and fixture-based migration tests.
2. Sanitize HTML import and `contenteditable` output; introduce CSP headers.
3. Rename “global search” in the UI, then index titles, tags, and block text
   across all notebooks.
4. Add rename/reorder/move and a recoverable recycle bin before more destructive
   actions are introduced.
5. Add undo/redo and keyboard/focus tests before expanding the editor toolbar.
6. Replace bitmap ink with a stroke model plus pen, highlighter, and eraser.
7. Move attachments out of the state JSON into dedicated blob storage, then add
   inline images and attachment previews/printouts.
8. Add export/restore tests and an explicit data-backup command to the server.

## Microsoft feature references

The comparison and roadmap use Microsoft's documentation as the product
baseline. Because capabilities differ between Windows, macOS, web, iOS, and
Android, verify a feature against the intended client before implementing it.

- [Microsoft OneNote help & learning](https://support.microsoft.com/onenote)
- [Differences between OneNote versions](https://support.microsoft.com/office/what-s-the-difference-between-the-onenote-versions-a624e692-b78b-4c09-b07f-46181958118f)
- [Microsoft Graph OneNote API overview](https://learn.microsoft.com/graph/integrate-with-onenote)
- [OneNote API content model](https://learn.microsoft.com/graph/onenote-concept-overview)

This is a living comparison: update the audit date, matrix, tests, and roadmap
whenever either this project or the reference OneNote clients change.
