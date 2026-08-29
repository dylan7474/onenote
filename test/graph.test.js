'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { createApp, seedNotebook } = require('./harness');

const plain = (v) => JSON.parse(JSON.stringify(v));

const A_PAGE = {
    id: 'p_src', title: 'Trip Plan',
    createdAt: '2021-06-07T08:09:10.000Z', updatedAt: '2021-06-08T11:12:13.000Z',
    tags: ['Imported OneNote'],
    attachments: [{ id: 'a1', name: 'map.txt', type: 'text/plain', size: 5, createdAt: '2021-06-07T00:00:00.000Z', data: 'aGVsbG8=' }],
    blocks: [{
        id: 'b1', x: 0, y: 0,
        content: '<h2>Plan</h2><p data-tag="to-do"><input type="checkbox"> Pack</p>'
            + '<p><span class="inline-attachment" data-attachment-id="a1" data-attachment="map.txt">map.txt</span></p>',
    }],
};

test('*ToGraph: internal names map to Graph displayName / content', (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);

    assert.deepEqual(plain(window.notebookToGraph({ name: 'Personal' })), { displayName: 'Personal' });
    assert.deepEqual(plain(window.notebookToGraph({})), { displayName: 'Untitled notebook' });
    assert.deepEqual(plain(window.sectionToGraph({ name: 'Work' })), { displayName: 'Work' });

    const content = window.pageToGraphContent(A_PAGE);
    assert.equal(content, window.pageToOneNoteHtml(A_PAGE), 'page content is the OneNote-compatible HTML');
    assert.match(content, /<title>Trip Plan<\/title>/);
});

test('pageFromGraph: Graph page resource -> internal page', async (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);

    const graphPage = {
        id: '1-abc',
        title: 'Trip Plan',
        level: 1,
        createdDateTime: '2021-06-07T08:09:10Z',
        lastModifiedDateTime: '2021-06-08T11:12:13Z',
        content: window.pageToGraphContent(A_PAGE),
    };

    const page = await window.pageFromGraph(graphPage);
    assert.equal(page.id, '1-abc');
    assert.equal(page.title, 'Trip Plan');
    assert.equal(page.level, 1);
    assert.equal(page.createdAt, '2021-06-07T08:09:10.000Z');
    assert.equal(page.updatedAt, '2021-06-08T11:12:13.000Z');
    assert.ok(plain(page.tags).includes('Graph'));
    assert.match(page.blocks.map((b) => b.content).join(''), /data-tag="to-do"><input type="checkbox">\s*Pack/);
    assert.equal(page.attachments.length, 1);
    assert.equal(window.atob(page.attachments[0].data), 'hello');
});

test('pageFromGraph: missing fields fall back safely', async (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);
    const page = await window.pageFromGraph({ content: '<p>hi</p>' });
    assert.equal(page.title, 'Untitled Page');
    assert.equal(page.level, 0);
    assert.match(page.id, /^page_/);
    assert.equal(new Date(page.createdAt).getUTCFullYear(), new Date().getUTCFullYear());
    assert.equal(page.blocks.length, 1);
});

test('notebookFromGraph: assembles a full internal notebook', async (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);

    const nb = await window.notebookFromGraph({
        id: 'nb-1', displayName: 'From Graph',
        sections: [
            { id: 's-1', displayName: 'One', pages: [{ id: 'p-1', title: 'A', content: '<p>a</p>' }] },
            { id: 's-2', displayName: 'Two', pages: [
                { id: 'p-2', title: 'B', content: '<p>b</p>' },
                { id: 'p-3', title: 'B sub', level: 1, content: '<p>b2</p>' },
            ] },
        ],
    });

    assert.equal(nb.name, 'From Graph');
    assert.deepEqual(plain(nb.sections.map((s) => s.name)), ['One', 'Two']);
    for (const s of nb.sections) assert.match(s.color, /^#[0-9a-f]{6}$/);
    assert.notEqual(nb.sections[0].color, nb.sections[1].color, 'sections get distinct palette colours');
    assert.deepEqual(plain(nb.sections[1].pages.map((p) => [p.title, p.level])), [['B', 0], ['B sub', 1]]);
});

test('Graph round-trip: pageFromGraph -> pageToGraphContent keeps the essentials', async (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);

    const original = window.pageToGraphContent(A_PAGE);
    const back = window.pageToGraphContent(await window.pageFromGraph({
        id: 'x', title: 'Trip Plan',
        createdDateTime: A_PAGE.createdAt, lastModifiedDateTime: A_PAGE.updatedAt,
        content: original,
    }));

    assert.match(back, /<title>Trip Plan<\/title>/);
    assert.match(back, /<meta name="created" content="2021-06-07T08:09:10.000Z">/);
    assert.match(back, /<p data-tag="to-do">Pack<\/p>/);
    assert.match(back, /<object data-attachment="map\.txt" data="data:text\/plain;base64,aGVsbG8=" type="text\/plain">/);
});
