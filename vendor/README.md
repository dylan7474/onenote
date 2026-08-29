# Vendored browser libraries

These files are checked in and served by `server.js` from `/vendor/` so the app
has no runtime dependency on third-party CDNs and always sanitises imported
OneNote HTML with a known, pinned DOMPurify build. See `PLAN.md` Phase 0.

Do not edit these files. To upgrade, replace a file with an official release of
the same library, update the version and SHA-256 below, and re-check the
`sanitizeHtml` behaviour in `index.html`.

| File | Library | Version | Source | SHA-256 |
| --- | --- | --- | --- | --- |
| `dompurify.min.js` | [DOMPurify](https://github.com/cure53/DOMPurify) | 3.4.14 | `https://cdnjs.cloudflare.com/ajax/libs/dompurify/3.4.14/purify.min.js` | `c2f26ea4fc0d88141c9aa430eb515ac86fce59418ceebd85fa475b87a8d6c3e6` |
| `lucide.min.js` | [Lucide](https://github.com/lucide-icons/lucide) | 1.37.0 | `https://unpkg.com/lucide@1.37.0/dist/umd/lucide.min.js` | `970650887f4992a6882dcd3f3b3d71a12bfcf89ec513a3e6404a12031a34f670` |
| `jszip.min.js` | [JSZip](https://github.com/Stuk/jszip) | 3.10.1 | `https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js` | `acc7e41455a80765b5fd9c7ee1b8078a6d160bbbca455aeae854de65c947d59e` |

Verify with `sha256sum vendor/*.js` from the repo root.

## Licenses

- DOMPurify — Apache-2.0 OR MPL-2.0 (see header in `dompurify.min.js`)
- Lucide — ISC
- JSZip — MIT OR GPL-3.0-or-later

## Not vendored

Tailwind (`cdn.tailwindcss.com`) and the Inter web font
(`fonts.googleapis.com`) are still loaded from a CDN by `index.html`; the
Tailwind play CDN has no drop-in static build. Self-hosting those is tracked
separately from Phase 0.
