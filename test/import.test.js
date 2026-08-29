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
    assert.deepEqual(plain(page.tags), ['Imported OneNote']);
    assert.equal(page.blocks.length, 1, 'whole body collapses into one block (REVIEW §2)');

    const html = page.blocks[0].content;
    assert.doesNotMatch(html, /<script/i);
    assert.doesNotMatch(html, /onclick/i);
    assert.doesNotMatch(html, /onerror/i);
    assert.match(html, /data-tag="to-do"/, 'data-tag preserved');
    assert.match(html, /data-tag="to-do:completed"/);
    assert.match(html, /position:\s*absolute/, 'positioned outline preserved (but not split out)');
    assert.match(html, /<table/i);
    assert.match(html, /<img[^>]+src="data:image\/png/i, 'inline data-URL image kept');
    assert.doesNotMatch(html, /<object/i, '<object> replaced by a span');
    assert.match(html, /class="inline-attachment"[^>]+data-attachment-id=/);
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

test('HTML import: creation date is fabricated, not read from <meta name="created">', async (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);
    seedNotebook(window);

    await window.parseHtmlImport(fixtureFile(window, 'onenote-page.html', 'text/html', 'onenote-page.html'));
    const page = getState(window).notebooks[0].sections[0].pages[0];

    // CURRENT: set to import time.
    assert.equal(new Date(page.createdAt).getUTCFullYear(), new Date().getUTCFullYear());
});

test('HTML import: reads <meta name="created"> (REVIEW §2)', { todo: true }, async (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);
    seedNotebook(window);
    await window.parseHtmlImport(fixtureFile(window, 'onenote-page.html', 'text/html', 'onenote-page.html'));
    const page = getState(window).notebooks[0].sections[0].pages[0];
    assert.equal(new Date(page.createdAt).toISOString(), '2019-06-14T17:30:00.000Z');
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
    assert.equal(pageOne.attachments.length, 1, 'sibling <object> payload resolved from the zip');
    assert.equal(pageOne.attachments[0].name, 'attachment.txt');
    assert.match(pageOne.blocks[0].content, /class="inline-attachment"/);
    assert.match(pageOne.blocks[0].content, /data-tag="to-do"/);
});

test('ZIP import: <img> whose bytes live in a *_files/ folder is left unresolved (REVIEW §2)', async (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);
    seedNotebook(window);

    await importZipFixture(window, 'My Section.zip', 'zip-src');
    const pageOne = getState(window).notebooks[0].sections.at(-1).pages[0];

    // CURRENT: the src is kept verbatim; the image is not inlined or attached.
    assert.match(pageOne.blocks[0].content, /<img[^>]+src="Page One_files\/diagram\.png"/);
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

test('decodeDataUrl: a bare % in a non-base64 data URL currently throws (REVIEW §9)', (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);
    assert.throws(() => window.decodeDataUrl('data:text/plain,100%'), { name: 'URIError' });
});
