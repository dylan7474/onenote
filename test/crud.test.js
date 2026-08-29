'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { createApp, seedNotebook, getState } = require('./harness');

const plain = (arr) => Array.from(arr);

test('getActive* fall back to the first entry when the active id is stale', (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);
    seedNotebook(window);
    const st = getState(window);
    st.activeNotebookId = 'gone';
    st.activeSectionId = 'gone';
    st.activePageId = 'gone';

    assert.equal(window.getActiveNotebook().id, 'nb_seed');
    assert.equal(window.getActiveSection().id, 'sec_seed');
    assert.equal(window.getActivePage().id, 'page_seed');
});

test('promptCreateNotebook adds a notebook and makes it active', (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);
    seedNotebook(window);
    window.prompt = () => 'Travel';

    window.promptCreateNotebook();

    const st = getState(window);
    assert.equal(st.notebooks.length, 2);
    const created = st.notebooks.at(-1);
    assert.equal(created.name, 'Travel');
    assert.equal(created.sections[0].name, 'Quick Notes');
    assert.equal(st.activeNotebookId, created.id);
    assert.equal(st.activeSectionId, created.sections[0].id);
    assert.equal(st.activePageId, created.sections[0].pages[0].id);
});

test('promptCreateNotebook is a no-op on an empty name', (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);
    seedNotebook(window);
    window.prompt = () => '   ';

    window.promptCreateNotebook();
    assert.equal(getState(window).notebooks.length, 1);
});

test('promptCreateSection appends a section to the active notebook', (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);
    seedNotebook(window);
    window.prompt = () => 'Meetings';

    window.promptCreateSection();

    const nb = getState(window).notebooks[0];
    assert.equal(nb.sections.length, 2);
    assert.equal(nb.sections.at(-1).name, 'Meetings');
    assert.equal(getState(window).activeSectionId, nb.sections.at(-1).id);
});

test('addNewPage prepends an Untitled Page and selects it', (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);
    seedNotebook(window);

    window.addNewPage();

    const pages = getState(window).notebooks[0].sections[0].pages;
    assert.equal(pages.length, 2);
    assert.equal(pages[0].title, 'Untitled Page');
    assert.equal(getState(window).activePageId, pages[0].id);
});

test('deletePage keeps at least one page', (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);
    const { page } = seedNotebook(window);

    window.deletePage(page.id);
    assert.equal(getState(window).notebooks[0].sections[0].pages.length, 1, 'last page is protected');

    window.addNewPage();
    const pages = getState(window).notebooks[0].sections[0].pages;
    window.deletePage(pages[0].id);
    assert.equal(getState(window).notebooks[0].sections[0].pages.length, 1);
});

test('makeSubpage / promotePage move a page between nesting levels', (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);
    const { section } = seedNotebook(window, {
        section: {
            id: 'sec_seed', name: 'S', color: '#8b5cf6',
            pages: [
                { id: 'p1', title: 'One', createdAt: '2020-01-01T00:00:00.000Z', updatedAt: '2020-01-01T00:00:00.000Z', tags: [], blocks: [{ id: 'b1', content: '<p>1</p>', x: 0, y: 0 }] },
                { id: 'p2', title: 'Two', createdAt: '2020-01-01T00:00:00.000Z', updatedAt: '2020-01-01T00:00:00.000Z', tags: [], blocks: [{ id: 'b2', content: '<p>2</p>', x: 0, y: 0 }] },
            ],
        },
    });

    window.makeSubpage('p1'); // index 0 -> rejected
    assert.equal(section.pages[0].level, undefined, 'first page cannot become a subpage');

    window.makeSubpage('p2');
    assert.equal(getState(window).notebooks[0].sections[0].pages[1].level, 1);

    window.promotePage('p2');
    assert.equal(getState(window).notebooks[0].sections[0].pages[1].level, 0);
});

test('deleteBlock removes only the target block', (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);
    const { page } = seedNotebook(window);
    page.blocks = [
        { id: 'blk_a', content: '<p>a</p>', x: 0, y: 0 },
        { id: 'blk_b', content: '<p>b</p>', x: 0, y: 0 },
    ];

    window.deleteBlock('blk_a');

    const ids = plain(getState(window).notebooks[0].sections[0].pages[0].blocks.map((b) => b.id));
    assert.deepEqual(ids, ['blk_b']);
});
