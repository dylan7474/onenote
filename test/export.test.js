'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { createApp, seedNotebook, getState } = require('./harness');

// Capture the data: URL that exportCurrentNotebook drives an <a download> with.
function captureDownload(window, run) {
    const captured = {};
    window.HTMLAnchorElement.prototype.click = function click() {
        captured.href = this.getAttribute('href');
        captured.name = this.getAttribute('download');
    };
    run();
    return captured;
}

test('exportCurrentNotebook emits JSON that round-trips the active notebook', (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);
    const { notebook } = seedNotebook(window);
    notebook.name = 'Field Notes';
    notebook.sections[0].pages[0].blocks[0].content = '<p>keep me</p>';

    const { href, name } = captureDownload(window, () => window.exportCurrentNotebook());

    assert.match(href, /^data:text\/json;charset=utf-8,/);
    assert.equal(name, 'Field_Notes_onenote_export.json', 'spaces -> underscores, fixed suffix');

    const parsed = JSON.parse(decodeURIComponent(href.replace(/^data:text\/json;charset=utf-8,/, '')));
    assert.equal(parsed.name, 'Field Notes');
    assert.equal(parsed.sections.length, 1);
    assert.equal(parsed.sections[0].pages[0].title, 'Seed Page');
    assert.equal(parsed.sections[0].pages[0].blocks[0].content, '<p>keep me</p>');
});

test('exportCurrentNotebook is a no-op when there is no notebook', (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);
    window.__setState({
        notebooks: [], activeNotebookId: null, activeSectionId: null, activePageId: null,
        isDrawMode: false, darkMode: false, sectionLayout: 'vertical',
    });

    const captured = captureDownload(window, () => window.exportCurrentNotebook());
    assert.equal(captured.href, undefined, 'no download triggered');
});

const RT_PAGE = {
    id: 'p_rt',
    title: 'Round Trip',
    createdAt: '2021-01-02T03:04:05.000Z',
    updatedAt: '2021-02-03T04:05:06.000Z',
    tags: ['Imported OneNote', 'Important'],
    attachments: [
        { id: 'a1', name: 'note.txt', type: 'text/plain', size: 5, createdAt: '2021-01-01T00:00:00.000Z', data: 'aGVsbG8=' },
    ],
    blocks: [
        {
            id: 'b1', x: 0, y: 0,
            content: '<h2>Heading</h2><p>Body text.</p>'
                + '<p data-tag="to-do"><input type="checkbox"> Task one</p>'
                + '<p data-tag="to-do:completed"><input type="checkbox" checked> Task two</p>'
                + '<p><span class="inline-attachment" data-attachment-id="a1" data-attachment="note.txt">note.txt</span></p>',
        },
        { id: 'b2', x: 0, y: 0, content: '<ul><li>bullet</li></ul>' },
    ],
};

test('pageToOneNoteHtml: OneNote "supported input" shape (REVIEW §3)', (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);
    seedNotebook(window, { page: RT_PAGE });

    const html = window.pageToOneNoteHtml(window.getActivePage());

    assert.match(html, /^<!DOCTYPE html>/);
    assert.match(html, /<meta name="created" content="2021-01-02T03:04:05.000Z">/);
    assert.match(html, /<meta name="lastModified" content="2021-02-03T04:05:06.000Z">/);
    assert.match(html, /<title>Round Trip<\/title>/);
    // one positioned <div> per block
    assert.equal((html.match(/position:absolute/g) || []).length, 2);
    // checkboxes -> data-tag, no <input> survives
    assert.doesNotMatch(html, /<input/i);
    assert.match(html, /<p data-tag="to-do">Task one<\/p>/);
    assert.match(html, /<p data-tag="to-do:completed">Task two<\/p>/);
    // inline chip -> <object data-attachment> with an embedded data URL
    assert.match(html, /<object data-attachment="note\.txt" data="data:text\/plain;base64,aGVsbG8=" type="text\/plain"><\/object>/);
});

test('export -> re-import round-trips a page', async (t) => {
    const src = createApp();
    t.after(src.dispose);
    seedNotebook(src.window, { page: RT_PAGE });
    const html = src.window.pageToOneNoteHtml(src.window.getActivePage());

    const dst = createApp();
    t.after(dst.dispose);
    seedNotebook(dst.window);
    await dst.window.parseHtmlImport(new dst.window.File([html], 'rt.html', { type: 'text/html' }));

    const page = getState(dst.window).notebooks[0].sections[0].pages[0];
    assert.equal(page.title, 'Round Trip');
    assert.equal(page.createdAt, '2021-01-02T03:04:05.000Z');
    assert.equal(page.updatedAt, '2021-02-03T04:05:06.000Z');
    assert.equal(page.blocks.length, 2, 'block count preserved');

    const body = page.blocks.map((b) => b.content).join('\n');
    assert.match(body, /data-tag="to-do"><input type="checkbox">\s*Task one/);
    assert.match(body, /data-tag="to-do:completed"><input type="checkbox" checked[^>]*>\s*Task two/);
    assert.match(body, /<ul><li>bullet<\/li><\/ul>/);

    assert.equal(page.attachments.length, 1);
    assert.equal(page.attachments[0].name, 'note.txt');
    assert.equal(src.window.atob(page.attachments[0].data), 'hello');
});

test('pageToOneNoteHtml: page-level attachments become a trailing <object> block', (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);
    seedNotebook(window, {
        page: {
            id: 'p', title: 'Files', createdAt: '2021-01-01T00:00:00.000Z', updatedAt: '2021-01-01T00:00:00.000Z',
            tags: [],
            attachments: [{ id: 'x', name: 'budget.csv', type: 'text/csv', size: 3, createdAt: '2021-01-01T00:00:00.000Z', data: 'YSxi' }],
            blocks: [{ id: 'b', x: 0, y: 0, content: '<p>No inline reference to the file.</p>' }],
        },
    });

    const html = window.pageToOneNoteHtml(window.getActivePage());
    assert.match(html, /<object data-attachment="budget\.csv" data="data:text\/csv;base64,YSxi" type="text\/csv"><\/object>/);
});

const TRIP_SECTION = {
    id: 'sec_trip', name: 'My Trip', color: '#8b5cf6',
    pages: [
        {
            id: 'sp0', title: 'Day 1',
            createdAt: '2021-05-01T00:00:00.000Z', updatedAt: '2021-05-02T00:00:00.000Z',
            tags: ['Imported OneNote'],
            attachments: [{ id: 'f1', name: 'ticket.txt', type: 'text/plain', size: 5, createdAt: '2021-05-01T00:00:00.000Z', data: 'aGVsbG8=' }],
            blocks: [{
                id: 'b0', x: 0, y: 0,
                content: '<p>Arrived.</p>'
                    + '<p data-tag="to-do"><input type="checkbox"> Check in</p>'
                    + '<p><span class="inline-attachment" data-attachment-id="f1" data-attachment="ticket.txt">ticket.txt</span></p>',
            }],
        },
        {
            id: 'sp1', title: 'Day 1 — Notes', level: 1,
            createdAt: '2021-05-01T00:00:00.000Z', updatedAt: '2021-05-01T00:00:00.000Z',
            tags: [],
            blocks: [{ id: 'b1', x: 0, y: 0, content: '<ul><li>subpage note</li></ul>' }],
        },
    ],
};

test('copyActivePageHtml writes text/html to the clipboard', async (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);
    seedNotebook(window, { page: RT_PAGE });

    await window.copyActivePageHtml();
    const items = window.__clipboard.items;
    assert.ok(Array.isArray(items) && items[0], 'navigator.clipboard.write called with a ClipboardItem');
    const htmlBlob = items[0].items['text/html'];
    assert.equal(htmlBlob.type, 'text/html');
    assert.ok(htmlBlob.size > 0);
    assert.ok(items[0].items['text/plain'], 'plain-text alternative included');

    // Fallback path (no ClipboardItem support) writes the string directly.
    window.ClipboardItem = undefined;
    await window.copyActivePageHtml();
    assert.match(window.__clipboard.text, /<title>Round Trip<\/title>/);
    assert.match(window.__clipboard.text, /position:absolute/);
});

test('sectionToOneNoteHtml: one <h1> per page, data-level on subpages, no positioning', (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);
    seedNotebook(window, { section: TRIP_SECTION });

    const html = window.sectionToOneNoteHtml(window.getActiveSection());

    assert.match(html, /<title>My Trip<\/title>/);
    assert.match(html, /<meta name="created" content="2021-05-01T00:00:00.000Z">/);
    assert.match(html, /<h1>Day 1<\/h1>/);
    assert.match(html, /<h1 data-level="1">Day 1 — Notes<\/h1>/);
    assert.doesNotMatch(html, /position:absolute/);
    assert.doesNotMatch(html, /<input/i);
    assert.match(html, /<object data-attachment="ticket\.txt" data="data:text\/plain;base64,aGVsbG8=" type="text\/plain">/);
});

test('section export -> re-import round-trips pages and subpage levels', async (t) => {
    const src = createApp();
    t.after(src.dispose);
    seedNotebook(src.window, { section: TRIP_SECTION });
    const html = src.window.sectionToOneNoteHtml(src.window.getActiveSection());

    const dst = createApp();
    t.after(dst.dispose);
    seedNotebook(dst.window);
    await dst.window.parseHtmlImport(new dst.window.File([html], 's.html', { type: 'text/html' }));

    const pages = getState(dst.window).notebooks[0].sections[0].pages;
    assert.deepEqual([pages[0].title, pages[1].title], ['Day 1', 'Day 1 — Notes']);
    assert.equal(pages[0].level || 0, 0);
    assert.equal(pages[1].level, 1, 'subpage level survives via data-level');
    assert.match(pages[0].blocks.map((b) => b.content).join(''), /data-tag="to-do"><input type="checkbox">\s*Check in/);
    assert.equal(pages[0].attachments.length, 1);
    assert.equal(src.window.atob(pages[0].attachments[0].data), 'hello');
    assert.match(pages[1].blocks[0].content, /subpage note/);
});
