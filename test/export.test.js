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
