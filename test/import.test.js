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

test('ZIP import: section from filename, pages sorted, folder depth -> level', async (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);
    seedNotebook(window);

    await importZipFixture(window, 'My Section.zip', 'zip-src');

    const nb = getState(window).notebooks[0];
    assert.equal(nb.sections.length, 2, 'seed section + imported section');
    const section = nb.sections.at(-1);
    assert.equal(section.name, 'My Section', 'name from the zip filename');
    assert.deepEqual(plain(section.pages.map((p) => p.title)), ['Page One', 'Page Two', 'Deep Dive']);

    // CURRENT: nesting is derived from folder depth below the common prefix.
    assert.deepEqual(plain(section.pages.map((p) => p.level)), [0, 0, 1]);

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

test('ZIP import: <img> whose bytes live in a *_files/ folder is inlined as a data URL (REVIEW §2)', async (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);
    seedNotebook(window);

    await importZipFixture(window, 'My Section.zip', 'zip-src');
    const html = getState(window).notebooks[0].sections.at(-1).pages[0].blocks[0].content;

    assert.match(html, /<img[^>]+src="data:image\/png;base64,[A-Za-z0-9+/=]+"/, 'resolved image embedded');
    assert.doesNotMatch(html, /src="Page One_files\/diagram\.png"/, 'original relative src replaced');
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
