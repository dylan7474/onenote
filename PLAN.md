# OneNote Compatibility — Phased Implementation Plan

Companion to [`REVIEW.md`](./REVIEW.md). Section references (§) point at that
document. Sequenced so each phase unblocks the next: sanitization before
ingesting untrusted HTML, import fidelity before export (so there is round-trip
test material), and Microsoft Graph last since it is the largest lift and
depends on the aligned data model.

Each table row is deliberately a standalone, independently reviewable
PR-sized slice, matching how the repo already works.

---

## Phase 0 — Safety & test foundations

**Why first:** real OneNote HTML cannot be safely ingested without
sanitization, and no import/export change can be verified without fixtures.

| Task | REVIEW ref | Status | Code touchpoints |
| --- | --- | --- | --- |
| Add DOMPurify (pinned, self-hosted) with an allowlist that keeps `style`, `data-tag`, `data-id`, `data-render-src`, `data-fullres-src`, `width`/`height`, table attrs | §6 | done (#23) | `sanitizeHtml()` / `sanitizeImportedTree()`; called on every import path, `saveBlockContent`, `insertInlineAttachments`, and defensively in `renderActivePage` / `documentAttachmentIds` |
| Self-host Lucide instead of `unpkg@latest` (plus DOMPurify + JSZip) | §6 | done (#23) | `vendor/` + `vendor/README.md`; `server.js` `/vendor/*`; `deploy.sh` |
| OneNote export fixtures: single-page HTML, ZIP export, `data-tag` paragraphs, attachments | §2, §5 | done (#24, synthetic) | `test/fixtures/` — synthetic, to be supplemented with real captures |
| Test harness (node:test + jsdom) with characterization tests over the current importers/exporter/sanitiser/CRUD as-is | — | done (#24) | `test/` (27 tests; 2 `todo` placeholders for Phase 1); CI in `.github/workflows/test.yml` |
| README correction: `.one`/`.onepkg` is documented ([MS-ONESTORE]/[MS-ONE]) | §7 | done (#23) | `README.md`, `REVIEW.md` §7 |

**Done when:** malformed or hostile imported HTML cannot execute script; existing
import/export behavior is pinned by tests. — **met.**

---

## Phase 1 — HTML/ZIP import fidelity

**Why:** highest-value data-loss fixes; the importer is already ~80% structured
for this.

| Task | REVIEW ref | Status | Code touchpoints |
| --- | --- | --- | --- |
| Resolve `<img src>` on import — inline ZIP-resident / `data:` images as data URLs; leave remote and missing refs alone | §2 | done (#25) | `inlineImages()` in `extractOneNoteAttachments`; `parseZipImport` `resolveFile` now types by extension |
| Parse top-level `position:absolute` outline `<div>`s into separate blocks with `x`/`y`/`width` instead of one block at `0,0` | §2 | done (#26) | `extractPositionedBlocks()` / `blocksFromImportedDoc()`; `px`/`pt` parsed; stray body content kept as a trailing block. Geometry is stored + round-trips in JSON export; the editor still stacks vertically (free-form canvas is later). |
| Read `<meta name="created">` / `lastModified` for `createdAt`/`updatedAt` | §2 | done (#27) | `readDocDates()`; case-insensitive meta lookup, offset-aware parse, falls back to import time |
| Import `<p data-tag="to-do｜important｜question｜…">` → checkbox UI + tag chips | §5 | done (#28) | `applyDataTags()` — `to-do`/`to-do:completed` → checkbox (attribute kept for round-trip); other values → page chips (known set → friendly labels, unknown → title-cased); comma lists supported |
| Split whole-section HTML export (one doc, `<h1>`-delimited) into multiple pages | §2 | done (#30) | `splitSectionHtml()` — only when ≥2 top-level `<h1>` and no positioned outline; inline attachments routed to the page that references them; first page keeps the doc's `<meta>` dates |
| Rework ZIP subpage-level detection — subpages named by filename (`Foo 1.html` next to `Foo.html`) — beside the folder-depth heuristic | §4 | done (#30) | `subpageInfo()` — numbered members → level 1, sorted after their group's own page; folder depth kept as the fallback |
| Fix importer bugs: `decodeDataUrl` malformed-escape crash (#25); ZIP asset basename fallback across `*_files/`; base64 size math padding | §9 | done (#25, #29) | `resolveFile` now falls back to a unique same-basename zip entry; `base64ByteLength()` accounts for `=` padding |
| Detect Graph-host `<object data>` URLs and surface "requires sign-in" instead of an empty chip | §9 | done (#29) | remote `<object data>` → `span.inline-attachment.attachment-unresolved` with `data-attachment-source` and a "(unavailable)" label, no dangling id |

**Done when:** a real OneNote HTML page and ZIP export import with images intact,
correct timestamps, checkboxes/tags mapped, and layout approximating the
original. — **met** (all rows landed in #25–#30; `test/import.test.js` covers
each against synthetic fixtures).

---

## Phase 2 — Export & round-trip

**Why:** first real path *back into* OneNote; Phase 1 fixtures become round-trip
test cases.

| Task | REVIEW ref | Status | Code touchpoints |
| --- | --- | --- | --- |
| Export `data-tag` attributes instead of `<input type=checkbox>` | §5 | done (#31) | `blockContentToOneNoteHtml()` — checkbox → `data-tag="to-do"` / `to-do:completed` on its block-level host, `<input>` removed |
| "Export page as OneNote-compatible HTML" — `<head>` meta + one `position:absolute` `<div>` per block, tables, `<img>` data URLs, `<object data-attachment>` | §3 | done (#31) | `pageToOneNoteHtml()` / `exportActivePageHtml()`; File menu "Export page as HTML" |
| "Copy page as HTML" to clipboard (`text/html`) for direct paste into desktop OneNote | §3 | done (#32) | `copyActivePageHtml()` — `navigator.clipboard.write` with `text/html` + `text/plain`, `writeText` fallback; File menu item |
| Export whole section as one HTML file | §2 | done (#32) | `sectionToOneNoteHtml()` — `<h1>` per page with `data-level` for subpages; re-imported by the section-split path (subpage level survives) |
| Self round-trip test: export → re-import → assert structural equality | — | done (#31, #32) | `test/export.test.js` — page and section round-trips |

**Done when:** a page exported as HTML pastes/imports into real OneNote with
formatting, checkboxes, tables, and images preserved. — **met** (#31–#32).

_Not serialized: non-`to-do` page tag chips (Important, Question, …) — OneNote's
content HTML has no page-level tag slot; the `to-do` checkbox does round-trip._

---

## Phase 3 — Data-model alignment & section colors

**Why:** small now, avoids a painful adapter when Graph lands.

| Task | REVIEW ref | Status | Code touchpoints |
| --- | --- | --- | --- |
| Introduce a Graph⇄internal adapter: `name`→`displayName`, keep `title`, page-level `content` vs `block.content` | §4 | done (#35) | `notebookToGraph()` / `sectionToGraph()` / `pageToGraphContent()` and `pageFromGraph()` / `notebookFromGraph()` — pure shape conversion, no network; Phase 4 wires them to Graph calls |
| Map `section.color` to a named palette on import; deterministic assignment for new sections | §8 | done (#33) | `SECTION_COLOR_PALETTE` + `nearestSectionColor()` / `snapSectionColor()` / `sectionColorName()`; `Math.random()` picks replaced with rotation; JSON import snaps foreign hex |
| `page.level` handling: keep the 3-level UI model, serialize subpages as `<h1 data-level>` in exported content | §4 | done (#32) | `sectionToOneNoteHtml()` / `splitSectionHtml()` — landed with Phase 2 |
| Add a state schema version + migration hook | §4 | done (#34) | `SCHEMA_VERSION` in `index.html` + `server.js`; `migrateState()` runs on every load (v0→v1: snap colours, fill arrays, clamp `level`); `writeState` rejects a state stamped newer (409). `server.js` now exports `{ server, SCHEMA_VERSION }` for tests |

**Done when:** internal objects convert to/from Graph shapes without loss;
section colors survive a round-trip. — **met** (#33–#35).

---

## Phase 4 — Microsoft Graph OneNote API

**Why:** the actual supported interop path; depends on Phases 1–3 for the ingest
pipeline and data model.

| Task | REVIEW ref | Status | Code touchpoints |
| --- | --- | --- | --- |
| Server-side OAuth (auth-code + PKCE) and a `/api/graph/*` Graph proxy, config-gated by `GRAPH_CLIENT_ID`, keeping the client secret off the browser | §1 | done (#37) | `graph.js` — `/config`, `/login`, `/callback`, `/logout`, `/v1.0/*` proxy; in-memory session keyed by `onenote_gsid` cookie; token refresh on expiry |
| Browser sign-in UI — a "Connect OneNote" control near Import that drives `/api/graph/login` and reflects `/api/graph/config` | §1 | done (#38) | Import modal `#graphImportSection`; `refreshGraphStatus()` / `renderGraphImport()` / `connectGraph()` / `disconnectGraph()` |
| Import: walk `GET /me/onenote/notebooks` → sections → pages via the proxy; `content?includeIDs=true` through `notebookFromGraph()` | §1 | done (#38) | `importSelectedGraphNotebook()` — notebook picker, section/page walk, per-page content fetch |
| Handle Graph throttling (429 / `Retry-After`), pagination (`@odata.nextLink`) | §1, §9 | done (#38) | `graphGet()` (429 back-off) / `graphGetAll()` (follows `@odata.nextLink`) |
| Push: `POST /me/onenote/sections/{id}/pages` (`text/html`) using `pageToGraphContent()` | §1, §3 | done (#39) | `pushActivePageToGraph()` — section picker in the Graph panel, `graphSend()` POST, opens the created page's `oneNoteWebUrl`; proxy body limit raised to 25 MiB |
| Resolve a page's Graph-hosted images / file attachments on import (they arrive as `graph.microsoft.com/.../resources/{id}/$value` links, unloadable without the bearer token) | §1, §2, §9 | done (#41) | `graph.js` `/api/graph/resource?url=` (`httpsBuffer()`, host-allowlisted, redirect-following, token not forwarded off-host) → `{ type, data:<base64> }`; client `graphResolveResource()` passed into `extractOneNoteAttachments()` / `inlineImages()` from `pageFromGraph()`; fetch failure keeps the "(unavailable)" placeholder |

**Done when:** a signed-in user can import a real Microsoft 365 notebook — with
its images and attachments intact — and push a page back that appears correctly
in OneNote. — **met** (#37–#39, #41).

---

## Phase 5 — `.onepkg` / `.one` (stretch)

Low priority; do only if Graph does not cover the need.

- `.onepkg` = CAB archive → extract, list contained `.one` files, report
  "native parsing not yet supported" as an honest intermediate. (§7)
  — **done.** `parseOnePkgImport()` reads just the MS-CAB directory
  (`listCabinetFiles()`: CFHEADER → CFFILE walk, no decompression), lists the
  contained `.one` sections with sizes, and shows a persistent Import-dialog
  notice (`showImportNotice()` / `#importNotice`) pointing at the HTML/ZIP
  export and Microsoft Graph paths. Bare `.one` files are recognised and
  reported the same way. No binary content is parsed and no state is mutated.
  `test/import.test.js` covers the cabinet inventory, the non-cabinet and bare
  `.one` cases, and the notice rendering/escaping.
- Full read-only `.one` parser (CAB → [MS-ONESTORE] object space → [MS-ONE]
  properties → block model → HTML). Multi-week; only with real demand for
  offline/legacy files. — not started.

---

## Sequencing & effort (relative)

```
Phase 0 ─┬─> Phase 1 ──> Phase 2 ─┐
         └─> Phase 3 ─────────────┴─> Phase 4 ──> Phase 5 (optional)
```

| Phase | Size | Shape |
| --- | --- | --- |
| 0 | S | 1 focused PR |
| 1 | L | 4–6 PRs, one per row |
| 2 | M | 2–3 PRs |
| 3 | M | 2 PRs |
| 4 | L | 4+ PRs, plus Azure app registration |
| 5 | XL | optional |
