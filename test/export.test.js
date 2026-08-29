'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { createApp, seedNotebook } = require('./harness');

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

// REVIEW §3 / PLAN Phase 2: there is no HTML or clipboard export yet.
test('exports page/section as OneNote-compatible HTML (REVIEW §3)', { todo: true }, (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);
    seedNotebook(window);
    assert.equal(typeof window.exportPageHtml, 'function');
});
