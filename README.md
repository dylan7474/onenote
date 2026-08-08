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
uses the server's `/api/state` endpoint to save and restore data. `DATA_FILE`
defaults to `./data/state.json`, and request bodies are limited to 50 MiB.

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

MHT imports support Microsoft OneNote's **Single File Web Page (`.mht`)** export
and retain resources embedded in its MIME archive, including inline images.
The import dialog can also import an entire folder of MHT or HTML exports as a
new notebook: first-level subfolders become sections, deeper folders preserve up
to two subpage levels, and an unreadable page does not prevent the remaining
pages from importing. Browser folder import cannot read native `.onepkg` files.
HTML imports can recover an attachment only when its OneNote `<object
data-attachment>` contains an embedded data URL. ZIP imports can also resolve a
referenced payload stored in the ZIP alongside the HTML page. Native `.one` and
`.onepkg` files remain unsupported.

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
  two levels of subpages, and collapsible subpage groups.
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
- JSON notebook export; JSON, HTML/HTM, MHT/MHTML, and ZIP-of-HTML import, preserving
  supported OneNote subpage levels from JSON metadata and ZIP folders.
- Debounced server-side JSON persistence, including a one-time migration from
  the older browser-local state.

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
| Capture and integrations | **Missing** | Add a web clipper/share target, email-to-note workflow, meeting details, Outlook tasks, and optional Microsoft Graph interoperability. |
| Collaboration and sync | **Missing** | State is one server-wide JSON document. Add accounts, private notebooks, invitations/links, permissions, real-time coauthoring, presence, comments/@mentions, conflict handling, offline cache, and multi-device sync. |
| History and recovery | **Missing** | The History menu is informational only. Add undo/redo across editing, autosaved revisions, page versions/diff/restore, author attribution, recent edits, deleted-notes recycle bin, and backup/restore. |
| Import/export/print | **Partial** | Project JSON, HTML-based, and OneNote Single File Web Page (`.mht`) imports exist; “OneNote ZIP” means ZIP files containing HTML, not native OneNote packages. Add sanitized import, PDF/HTML/Markdown export, print/preview, and document native `.one`/`.onepkg` limitations clearly. |
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
