'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { createApp, seedNotebook } = require('./harness');

// Records { url, init } for every call; routes by URL substring.
function makeFetch(routes, calls) {
    return async (url, init) => {
        calls.push({ url, init: init || {} });
        for (const [needle, handler] of routes) {
            if (url.includes(needle)) {
                const r = typeof handler === 'function' ? handler(url, init) : handler;
                const headers = new Map(Object.entries(r.headers || {}));
                return {
                    ok: (r.status || 200) < 400,
                    status: r.status || 200,
                    headers: { get: (k) => headers.get(k) ?? headers.get(k.toLowerCase()) ?? null },
                    json: async () => r.json,
                    text: async () => (r.text != null ? r.text : JSON.stringify(r.json)),
                };
            }
        }
        throw new Error('unexpected fetch: ' + url);
    };
}

const PAGE = {
    id: 'p_local', title: 'Weekly Notes',
    createdAt: '2023-01-02T03:04:05.000Z', updatedAt: '2023-01-02T03:04:05.000Z',
    tags: [], attachments: [],
    blocks: [{ id: 'b', x: 0, y: 0, content: '<h2>Notes</h2><p data-tag="to-do"><input type="checkbox"> Ship it</p>' }],
};

test('graphSend POSTs text/html through the proxy and returns the created page', async (t) => {
    const calls = [];
    const { window, dispose } = createApp({
        onFetch: makeFetch([
            ['/sections/s1/pages', { status: 201, json: { id: 'p9', links: { oneNoteWebUrl: { href: 'https://onenote.com/p9' } } } }],
        ], calls),
    });
    t.after(dispose);

    const created = await window.graphSend('/v1.0/me/onenote/sections/s1/pages', { body: '<html><title>x</title></html>', contentType: 'text/html' });
    assert.equal(created.id, 'p9');

    const post = calls.find((c) => c.url.includes('/pages'));
    assert.equal(post.url, '/api/graph/v1.0/me/onenote/sections/s1/pages');
    assert.equal(post.init.method, 'POST');
    assert.equal(post.init.headers['Content-Type'], 'text/html');
    assert.equal(post.init.body, '<html><title>x</title></html>');
});

test('pushActivePageToGraph sends the OneNote-compatible HTML of the open page', async (t) => {
    const calls = [];
    const { window, dispose } = createApp({
        onFetch: makeFetch([
            ['/api/graph/config', { json: { enabled: true, connected: true } }],
            ['/onenote/notebooks?', { json: { value: [{ id: 'nb1', displayName: 'Work' }] } }],
            ['/notebooks/nb1/sections', { json: { value: [{ id: 's1', displayName: 'Inbox' }, { id: 's2', displayName: 'Archive' }] } }],
            ['/sections/s1/pages', { status: 201, json: { id: 'created1', links: { oneNoteWebUrl: { href: 'https://onenote.com/created1' } } } }],
        ], calls),
    });
    t.after(dispose);
    seedNotebook(window, { page: PAGE });

    const toasts = [];
    window.showToast = (m) => toasts.push(m);
    let opened = null;
    window.open = (url) => { opened = url; return null; };

    await window.refreshGraphStatus();   // connected -> loads notebooks + sections, fills the selects
    assert.equal(window.document.getElementById('graphSectionSelect').value, 's1');

    await window.pushActivePageToGraph();

    const post = calls.find((c) => c.init.method === 'POST' && c.url.includes('/pages'));
    assert.ok(post, 'a page POST was made');
    assert.equal(post.url, '/api/graph/v1.0/me/onenote/sections/s1/pages');
    assert.equal(post.init.headers['Content-Type'], 'text/html');
    assert.match(post.init.body, /<title>Weekly Notes<\/title>/);
    assert.match(post.init.body, /<p data-tag="to-do">Ship it<\/p>/, 'checkbox serialised as data-tag, no <input>');
    assert.doesNotMatch(post.init.body, /<input/i);

    assert.ok(toasts.some((m) => /created in OneNote/i.test(m)));
    assert.equal(opened, 'https://onenote.com/created1');
});

test('pushActivePageToGraph surfaces a Graph error without throwing', async (t) => {
    const calls = [];
    const { window, dispose } = createApp({
        onFetch: makeFetch([
            ['/api/graph/config', { json: { enabled: true, connected: true } }],
            ['/onenote/notebooks?', { json: { value: [{ id: 'nb1', displayName: 'Work' }] } }],
            ['/notebooks/nb1/sections', { json: { value: [{ id: 's1', displayName: 'Inbox' }] } }],
            ['/sections/s1/pages', { status: 400, json: { error: { message: 'Invalid HTML input' } } }],
        ], calls),
    });
    t.after(dispose);
    seedNotebook(window, { page: PAGE });

    const toasts = [];
    window.showToast = (m) => toasts.push(m);
    window.open = () => null;

    await window.refreshGraphStatus();
    await window.pushActivePageToGraph();   // must not reject

    assert.ok(toasts.some((m) => /Invalid HTML input/.test(m)), 'error message shown to the user');
});
