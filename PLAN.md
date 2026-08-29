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
| Split whole-section HTML export (one doc, `<h1>`-delimited) into multiple pages | §2 | todo | `parseHtmlImport` |
| Rework ZIP subpage-level detection — flat `Notebook/Section/Page.html`, subpages by filename — replace the folder-depth heuristic | §4 | todo | `parseZipImport` |
| Fix importer bugs: `decodeDataUrl` malformed-escape crash (**done #25**); ZIP asset basename fallback across `*_files/`; base64 size math padding | §9 | partial (#25) | `decodeDataUrl` done; `resolveFile` basename fallback + size math todo |
| Detect Graph-host `<object data>` URLs and surface "requires sign-in" instead of an empty chip | §9 | todo | `extractOneNoteAttachments` |

**Done when:** a real OneNote HTML page and ZIP export import with images intact,
correct timestamps, checkboxes/tags mapped, and layout approximating the
original.

---

## Phase 2 — Export & round-trip

**Why:** first real path *back into* OneNote; Phase 1 fixtures become round-trip
test cases.

| Task | REVIEW ref | Code touchpoints |
| --- | --- | --- |
| Export `data-tag` attributes instead of `<input type=checkbox onclick=…>` (also removes inline handlers) | §5 | sample data `index.html:740`, `index.html:2235`; new serializer |
| "Export page/section as OneNote-compatible HTML" — Graph supported-input subset: `<html><head><title>`, `data-tag`, tables, `<img>` as data URLs, `<object data-attachment>` | §3 | new `exportPageHtml()` alongside `exportCurrentNotebook` (`index.html:2259`); File menu `index.html:187` |
| "Copy page as HTML" to clipboard (`text/html`) for direct paste into desktop OneNote | §3 | new toolbar/menu action |
| Self round-trip test: export → re-import → assert structural equality | — | `test/` |

**Done when:** a page exported as HTML pastes/imports into real OneNote with
formatting, checkboxes, tables, and images preserved.

---

## Phase 3 — Data-model alignment & section colors

**Why:** small now, avoids a painful adapter when Graph lands.

| Task | REVIEW ref | Code touchpoints |
| --- | --- | --- |
| Introduce a Graph⇄internal adapter (or rename): `name`→`displayName`, keep `title`, page-level `content` vs `block.content` | §4 | `server.js` validation (`server.js:49`), `index.html:855`, state schema |
| Map `section.color` to OneNote's ~16 named colors on import/export; stop random hex assignment | §8 | `SECTION_COLORS` (`index.html:485`, `index.html:2196`) |
| Decide `page.level` handling: keep the 3-level UI model, but serialize subpages as `<h1>`/indent depth in exported `content` | §4 | `normalizeImportedPages` (`index.html:2016-2033`), exporters |
| Add a state schema version + migration hook (supports Phase 4 sync and safe reshaping) | §4 | `server.js` `writeState`, `initializeState` (`index.html:687`) |

**Done when:** internal objects convert to/from Graph shapes without loss;
section colors survive a round-trip.

---

## Phase 4 — Microsoft Graph OneNote API

**Why:** the actual supported interop path; depends on Phases 1–3 for the ingest
pipeline and data model.

| Task | REVIEW ref | Code touchpoints |
| --- | --- | --- |
| Register an Azure AD app; add MSAL.js browser sign-in ("Connect OneNote" near Import, `index.html:171`) | §1 | new auth module |
| Server-side token exchange / Graph proxy endpoint (keep the client secret off the browser, handle CORS) | §1 | `server.js` — new `/api/graph/*` routes |
| Import: walk `GET /me/onenote/notebooks` → sections → pages; fetch each page `content?includeIDs=true` and run it through the Phase 1 importer | §1 | reuse `extractOneNoteAttachments`, block parser |
| Push: `POST /me/onenote/sections/{id}/pages` as `multipart/form-data` using the Phase 2 HTML serializer | §1, §3 | reuse `exportPageHtml()` |
| Handle Graph throttling (429 / `Retry-After`), pagination (`@odata.nextLink`), and resource URLs needing auth headers | §1, §9 | proxy layer |
| Map Graph `createdDateTime`/`lastModifiedDateTime`, `displayName` via the Phase 3 adapter | §4 | adapter |

**Done when:** a signed-in user can import a real Microsoft 365 notebook and push
a page back that appears correctly in OneNote.

---

## Phase 5 — `.onepkg` / `.one` (stretch)

Low priority; do only if Graph does not cover the need.

- `.onepkg` = CAB archive → extract, list contained `.one` files, report
  "native parsing not yet supported" as an honest intermediate. (§7)
- Full read-only `.one` parser (CAB → [MS-ONESTORE] object space → [MS-ONE]
  properties → block model → HTML). Multi-week; only with real demand for
  offline/legacy files.

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
