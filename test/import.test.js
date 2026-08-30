'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const {
    createApp, seedNotebook, getState, fixtureFile, importZipFixture,
} = require('./harness');

// These pin CURRENT importer behaviour. Gaps that REVIEW.md / PLAN.md Phase 1
// will close are marked `{ todo: true }` so they flip to real assertions then.

// Arrays/objects that cross back from the jsdom realm have a different prototype,
// which trips assert/strict's deep equality. Normalise primitives-only arrays.
const plain = (arr) => Array.from(arr);

test('HTML import: page shape and sanitisation', async (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);
    seedNotebook(window);

    await window.parseHtmlImport(fixtureFile(window, 'onenote-page.html', 'text/html', 'onenote-page.html'));

    const pages = getState(window).notebooks[0].sections[0].pages;
    const page = pages[0]; // parseHtmlImport unshifts
    assert.equal(page.title, 'Quarterly Planning', 'title from <title>');
    assert.deepEqual(plain(page.tags), ['Imported OneNote', 'Important'], 'data-tag="important" -> chip');
    // One positioned outline + stray body content outside it -> two blocks.
    assert.equal(page.blocks.length, 2);

    const outline = page.blocks[0].content;
    const html = page.blocks.map((b) => b.content).join('\n');
    assert.doesNotMatch(html, /<script/i);
    assert.doesNotMatch(html, /onclick/i);
    assert.doesNotMatch(html, /onerror/i);
    assert.match(outline, /data-tag="to-do"/, 'data-tag preserved');
    assert.match(outline, /data-tag="to-do:completed"/);
    // to-do paragraphs gain a checkbox; the completed one is checked.
    assert.match(outline, /<p data-tag="to-do"><input type="checkbox"[^>]*>\s*Draft the departmental budget/);
    assert.match(outline, /<p data-tag="to-do:completed"><input type="checkbox" checked[^>]*>\s*Book the offsite venue/);
    assert.match(outline, /<table/i);
    assert.match(outline, /<img[^>]+src="data:image\/png/i, 'inline data-URL image kept');
    assert.doesNotMatch(html, /<object/i, '<object> replaced by a span');
    assert.match(outline, /class="inline-attachment"[^>]+data-attachment-id=/);
    // The wrapping position:absolute <div> is consumed, not kept in content.
    assert.doesNotMatch(outline, /position:\s*absolute/);
    assert.equal(page.blocks[0].x, 48);
    assert.equal(page.blocks[0].y, 115);
    assert.equal(page.blocks[0].width, 624);
    assert.match(page.blocks[1].content, /Do not run me/, 'stray body content kept as a trailing block');
});

test('HTML import: positioned outlines become separate blocks with geometry (REVIEW §2)', async (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);
    seedNotebook(window);

    await window.parseHtmlImport(fixtureFile(window, 'onenote-multi-outline.html', 'text/html', 'onenote-multi-outline.html'));
    const page = getState(window).notebooks[0].sections[0].pages[0];

    assert.equal(page.blocks.length, 2, 'one block per outline, no trailing block');
    assert.deepEqual(
        plain(page.blocks).map((b) => ({ x: b.x, y: b.y, width: b.width })),
        [{ x: 48, y: 115, width: 624 }, { x: 64, y: 520, width: 300 }], // second outline is in pt
    );
    assert.match(page.blocks[0].content, /data-tag="important"/);
    assert.match(page.blocks[1].content, /<table/i);
    assert.doesNotMatch(page.blocks[0].content, /Second outline/);
});

test('HTML import: a document with no positioned outlines stays one block', async (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);
    seedNotebook(window);

    await window.parseHtmlImport(fixtureFile(window, 'plain.html', 'text/html', 'plain.html'));
    const page = getState(window).notebooks[0].sections[0].pages[0];

    assert.equal(page.blocks.length, 1);
    assert.equal(page.blocks[0].x, 0);
    assert.equal(page.blocks[0].y, 0);
    assert.equal(page.blocks[0].width, undefined);
    assert.match(page.blocks[0].content, /data-tag="to-do"><input type="checkbox"/);
    assert.deepEqual(plain(page.tags), ['Imported OneNote'], 'a bare to-do adds no chip');
});

test('HTML import: data-tag maps to checkboxes and page chips (REVIEW §5)', async (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);
    seedNotebook(window);

    await window.parseHtmlImport(fixtureFile(window, 'onenote-tags.html', 'text/html', 'onenote-tags.html'));
    const page = getState(window).notebooks[0].sections[0].pages[0];

    assert.deepEqual(
        plain(page.tags).slice().sort(),
        ['Critical', 'Important', 'Imported OneNote', 'Question', 'Web Site To Visit'],
        'known tags mapped to friendly labels, unknown tag title-cased, deduped',
    );

    const html = page.blocks[0].content;
    assert.match(html, /<p data-tag="to-do"><input type="checkbox"[^>]*>\s*Open task/);
    assert.match(html, /<p data-tag="to-do:completed"><input type="checkbox" checked[^>]*>\s*Finished task/);
    // comma-separated list on one element yields both chips, element keeps its attr
    assert.match(html, /<p data-tag="important, critical">Both at once\.<\/p>/);
});

test('HTML import: a whole-section export splits into one page per top-level <h1> (REVIEW §2)', async (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);
    seedNotebook(window);
    const nowYear = new Date().getUTCFullYear();

    await window.parseHtmlImport(fixtureFile(window, 'onenote-section.html', 'text/html', 'onenote-section.html'));
    // Imported pages are unshifted ahead of the seed page.
    const pages = plain(getState(window).notebooks[0].sections[0].pages).slice(0, 3);

    assert.deepEqual(pages.map((p) => p.title), ['Trip Notes', 'Packing List', 'Budget']);

    // First page keeps the document's created meta; the rest fall back to now.
    assert.equal(pages[0].createdAt, '2021-05-01T12:00:00.000Z');
    assert.equal(new Date(pages[1].createdAt).getUTCFullYear(), nowYear);

    // Section-wide data-tag chips are on every page.
    for (const p of pages) assert.ok(plain(p.tags).includes('Important'));

    // The <object> lives in the first page's fragment -> its attachment lands there.
    assert.equal(pages[0].attachments.length, 1);
    assert.equal(pages[0].attachments[0].name, 'itinerary.pdf');
    assert.equal(pages[0].attachments[0].size, window.atob(pages[0].attachments[0].data).length);
    assert.equal(pages[1].attachments.length, 0);

    // Content is routed to the right page.
    assert.match(pages[1].blocks[0].content, /Passport/);
    assert.match(pages[1].blocks[0].content, /<input type="checkbox" checked[^>]*>\s*Charger/);
    assert.match(pages[2].blocks[0].content, /<table/i);
    assert.doesNotMatch(pages[0].blocks[0].content, /Passport/);
});

test('HTML import: a single positioned-outline page is not split on its inner <h1>s', async (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);
    seedNotebook(window);

    await window.parseHtmlImport(fixtureFile(window, 'onenote-multi-outline.html', 'text/html', 'onenote-multi-outline.html'));
    const pages = getState(window).notebooks[0].sections[0].pages;
    assert.equal(pages[0].title, 'Two Outlines');
    assert.equal(pages[1].title, 'Seed Page', 'only one page imported, not one per inner <h1>');
    assert.equal(pages[0].blocks.length, 2, 'two outline blocks');
});

test('HTML import: embedded <object data-attachment> becomes a stored attachment', async (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);
    seedNotebook(window);

    await window.parseHtmlImport(fixtureFile(window, 'onenote-page.html', 'text/html', 'onenote-page.html'));
    const page = getState(window).notebooks[0].sections[0].pages[0];

    assert.equal(page.attachments.length, 1);
    assert.equal(page.attachments[0].name, 'agenda.pdf');
    assert.equal(page.attachments[0].type, 'application/pdf');
    assert.ok(page.attachments[0].data, 'base64 payload stored');
    // size is the real decoded byte length, not base64 length * 3/4 (REVIEW §9)
    assert.equal(page.attachments[0].size, window.atob(page.attachments[0].data).length);
});

test('HTML import: an unresolvable remote <object> (Graph resource) is flagged, not stored (REVIEW §9)', async (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);
    seedNotebook(window);

    await window.parseHtmlImport(fixtureFile(window, 'onenote-graph-object.html', 'text/html', 'onenote-graph-object.html'));
    const page = getState(window).notebooks[0].sections[0].pages[0];

    assert.equal(page.attachments.length, 0, 'nothing to store without an authenticated fetch');
    const html = page.blocks[0].content;
    assert.match(html, /class="inline-attachment attachment-unresolved"/);
    assert.match(html, /data-attachment-source="https:\/\/graph\.microsoft\.com\//);
    assert.match(html, /report\.docx \(unavailable\)/);
    assert.doesNotMatch(html, /data-attachment-id=/, 'no dangling attachment id');
});

test('HTML import: reads <meta name="created"> / "lastModified" (REVIEW §2)', async (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);
    seedNotebook(window);

    await window.parseHtmlImport(fixtureFile(window, 'onenote-page.html', 'text/html', 'onenote-page.html'));
    const page = getState(window).notebooks[0].sections[0].pages[0];

    assert.equal(page.createdAt, '2019-06-14T17:30:00.000Z', 'created meta (offset applied)');
    assert.equal(page.updatedAt, '2019-07-02T23:05:12.000Z', 'lastModified meta');
});

test('HTML import: missing / partial date meta falls back to import time', async (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);
    const nowYear = new Date().getUTCFullYear();

    seedNotebook(window);
    await window.parseHtmlImport(fixtureFile(window, 'onenote-multi-outline.html', 'text/html', 'onenote-multi-outline.html'));
    let page = getState(window).notebooks[0].sections[0].pages[0];
    assert.equal(page.createdAt, '2020-02-20T09:00:00.000Z', 'created meta honoured');
    assert.equal(new Date(page.updatedAt).getUTCFullYear(), nowYear, 'no lastModified -> now');

    await window.parseHtmlImport(fixtureFile(window, 'plain.html', 'text/html', 'plain.html'));
    page = getState(window).notebooks[0].sections[0].pages[0];
    assert.equal(new Date(page.createdAt).getUTCFullYear(), nowYear, 'no meta at all -> now');
});

test('JSON import: single-notebook backup is appended, subpage levels inferred', async (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);
    seedNotebook(window);
    const before = getState(window).notebooks.length;

    await window.parseJsonImport(
        fixtureFile(window, 'webapp-notebook.json', 'application/json', 'webapp-notebook.json')
    );

    const notebooks = getState(window).notebooks;
    assert.equal(notebooks.length, before + 1);
    const imported = notebooks.at(-1);
    assert.equal(imported.name, 'Imported Notebook');

    const levels = plain(imported.sections[0].pages.map((p) => p.level));
    assert.deepEqual(levels, [0, 1, 2], 'parentPageId chain -> levels, capped at 2');

    const root = imported.sections[0].pages[0].blocks[0].content;
    assert.doesNotMatch(root, /<script/i, 'JSON import sanitises block bodies');
    assert.doesNotMatch(root, /onclick/i);
});

test('normalizeImportedPages: explicit level wins, inferred chain caps at 2', (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);

    const out = window.normalizeImportedPages([
        { id: 'a' },
        { id: 'b', parentPageId: 'a' },
        { id: 'c', parentPageId: 'b' },
        { id: 'd', parentPageId: 'c' },
        { id: 'e', level: 5 },
        { id: 'f', level: -3 },
    ]);
    assert.deepEqual(plain(out.map((p) => p.level)), [0, 1, 2, 2, 2, 0]);
});

test('ZIP import: section name, sort order, and subpage levels (REVIEW §4)', async (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);
    seedNotebook(window);

    await importZipFixture(window, 'My Section.zip', 'zip-src');

    const nb = getState(window).notebooks[0];
    assert.equal(nb.sections.length, 2, 'seed section + imported section');
    const section = nb.sections.at(-1);
    assert.equal(section.name, 'My Section', 'name from the zip filename');
    assert.deepEqual(
        plain(section.pages.map((p) => p.title)),
        ['Page One', 'Page Two', 'Report', 'Report — Appendix A', 'Report — Appendix B', 'Deep Dive'],
    );

    // "Report 1/2.html" next to "Report.html" -> subpages (level 1);
    // "Subpages/Deep Dive.html" -> level 1 by folder depth.
    assert.deepEqual(plain(section.pages.map((p) => p.level)), [0, 0, 0, 1, 1, 1]);

    const pageOne = section.pages[0];
    assert.equal(pageOne.createdAt, '2022-01-10T08:00:00.000Z', 'created meta read from the zipped HTML');
    assert.equal(new Date(section.pages[1].createdAt).getUTCFullYear(), new Date().getUTCFullYear(), 'Page Two has no meta -> now');
    assert.equal(pageOne.attachments.length, 1, 'sibling <object> payload resolved from the zip');
    assert.equal(pageOne.attachments[0].name, 'attachment.txt');
    assert.match(pageOne.blocks[0].content, /class="inline-attachment"/);
    assert.match(pageOne.blocks[0].content, /data-tag="to-do"/);
    // The single positioned outline carries its geometry.
    assert.deepEqual(
        { x: pageOne.blocks[0].x, y: pageOne.blocks[0].y, width: pageOne.blocks[0].width },
        { x: 48, y: 90, width: 600 },
    );
});

test('ZIP import: an <img> is inlined by basename fallback; a missing one is left alone (REVIEW §2/§9)', async (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);
    seedNotebook(window);

    await importZipFixture(window, 'My Section.zip', 'zip-src');
    const html = getState(window).notebooks[0].sections.at(-1).pages[0].blocks[0].content;

    // src="assets/diagram.png" is not a real path in the zip; it resolves via
    // the unique "diagram.png" basename in "Page One_files/".
    assert.match(html, /<img[^>]+src="data:image\/png;base64,[A-Za-z0-9+/=]+"/, 'resolved image embedded');
    assert.doesNotMatch(html, /src="assets\/diagram\.png"/, 'original src replaced');
    // A referenced file that is not in the ZIP is left untouched, not dropped.
    assert.match(html, /<img[^>]+src="Page One_files\/missing\.png"/);
});

test('decodeDataUrl: base64 passthrough, plain text re-encoded, non-data-URL null', (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);

    const b64 = window.decodeDataUrl('data:text/plain;base64,aGk=');
    assert.equal(b64.type, 'text/plain');
    assert.equal(b64.data, 'aGk=');

    const plainText = window.decodeDataUrl('data:text/plain,hi%20there');
    assert.equal(plainText.type, 'text/plain');
    assert.equal(plainText.data, 'aGkgdGhlcmU='); // btoa('hi there')

    assert.equal(window.decodeDataUrl('https://example.com/x.png'), null);
});

test('decodeDataUrl: a malformed escape no longer aborts the import (REVIEW §9)', (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);
    const out = window.decodeDataUrl('data:text/plain,100%');
    assert.equal(out.type, 'text/plain');
    assert.equal(window.atob(out.data), '100%', 'raw body kept when decodeURIComponent fails');
});

// --- PLAN.md Phase 5: `.one` / `.onepkg` inventory (no native parsing) --------

// Build a minimal but structurally valid MS-CAB cabinet: CFHEADER + one
// CFFOLDER + a CFFILE record per entry. Only the fields `listCabinetFiles`
// reads (`coffFiles`, `cFiles`, and each CFFILE's `cbFile` + NUL-terminated
// name) need to be meaningful; folder data and compression are omitted.
function makeCab(entries) {
    const enc = new TextEncoder();
    const names = entries.map((e) => enc.encode(e.name));
    const HEADER = 36;
    const FOLDER = 8;
    const coffFiles = HEADER + FOLDER;
    const filesLen = names.reduce((sum, n) => sum + 16 + n.length + 1, 0);
    const buf = Buffer.alloc(coffFiles + filesLen);

    buf.write('MSCF', 0, 'ascii');
    buf.writeUInt32LE(buf.length, 8);        // cbCabinet
    buf.writeUInt32LE(coffFiles, 16);        // coffFiles
    buf.writeUInt8(3, 24);                   // versionMinor
    buf.writeUInt8(1, 25);                   // versionMajor
    buf.writeUInt16LE(1, 26);               // cFolders
    buf.writeUInt16LE(entries.length, 28);  // cFiles

    let off = coffFiles;
    entries.forEach((e, i) => {
        buf.writeUInt32LE(e.size >>> 0, off); // CFFILE.cbFile
        Buffer.from(names[i]).copy(buf, off + 16);
        buf.writeUInt8(0, off + 16 + names[i].length);
        off += 16 + names[i].length + 1;
    });
    return buf;
}

const cabFile = (window, name, entries) =>
    new window.File([new window.Uint8Array(makeCab(entries))], name, { type: '' });

test('.onepkg import: inventories the contained .one files, reports no native parsing (PLAN Phase 5)', async (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);
    seedNotebook(window);

    const file = cabFile(window, 'Team Notebook.onepkg', [
        { name: '{4B4C7E60-1111-4A2B-9C3D-000000000001}.one', size: 4096 },
        { name: 'OneToc2', size: 240 },
        { name: '{4B4C7E60-2222-4A2B-9C3D-000000000002}.one', size: 2048 },
    ]);
    const report = await window.parseOnePkgImport(file);

    assert.equal(report.ok, false);
    assert.deepEqual(plain(report.oneFiles), [
        '{4B4C7E60-1111-4A2B-9C3D-000000000001}.one (4.0 KB)',
        '{4B4C7E60-2222-4A2B-9C3D-000000000002}.one (2.0 KB)',
    ], 'only the .one members are listed, with sizes');
    assert.match(report.message, /2 \.one section files/);
    assert.match(report.message, /not supported yet/i);
    assert.match(report.message, /Microsoft Graph/);
    // Nothing was imported into the seeded notebook.
    assert.equal(getState(window).notebooks[0].sections[0].pages.length, 1);
});

test('.onepkg import: a non-cabinet file is reported, not silently dropped', async (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);
    seedNotebook(window);

    const file = new window.File([new window.Uint8Array([0x50, 0x4b, 0x03, 0x04, 1, 2, 3, 4])], 'bad.onepkg', { type: '' });
    const report = await window.parseOnePkgImport(file);

    assert.equal(report.ok, false);
    assert.match(report.message, /not a readable OneNote package/);
    assert.match(report.message, /MSCF/);
});

test('.one import: honest report about the binary format, no import', async (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);
    seedNotebook(window);

    const file = new window.File([new window.Uint8Array(64)], 'Chapter 1.one', { type: '' });
    const report = await window.parseOnePkgImport(file);

    assert.equal(report.ok, false);
    assert.deepEqual(plain(report.oneFiles), []);
    assert.match(report.message, /native OneNote section file/);
    assert.match(report.message, /MS-ONESTORE/);
    assert.equal(getState(window).notebooks[0].sections[0].pages.length, 1);
});

test('showImportNotice: renders persistent, escaped advisories into the Import modal', (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);

    const notice = window.document.getElementById('importNotice');
    assert.ok(notice, 'the modal has an #importNotice slot');
    assert.ok(notice.classList.contains('hidden'), 'hidden until there is something to say');

    window.showImportNotice(['a.onepkg: 1 .one section file — <x>.one (1 B).']);
    assert.equal(notice.classList.contains('hidden'), false);
    assert.match(notice.textContent, /a\.onepkg: 1 \.one section file/);
    assert.doesNotMatch(notice.innerHTML, /<x>\.one/, 'message text is HTML-escaped');
    assert.match(notice.innerHTML, /&lt;x&gt;\.one/);

    // closeImportModal clears it so a stale notice never greets the next open.
    window.closeImportModal();
    assert.ok(notice.classList.contains('hidden'));
    assert.equal(notice.innerHTML, '');
});
