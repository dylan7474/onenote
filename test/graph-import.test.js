'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { createApp, seedNotebook, getState } = require('./harness');

// jsdom-realm arrays trip assert/strict deep equality; compare by value.
const plain = (v) => JSON.parse(JSON.stringify(v));

// A tiny Graph-proxy mock. `routes` maps a URL-substring to a handler that
// returns { status?, json?, text?, headers? }.
function makeFetch(routes, log) {
    return async (url) => {
        if (log) log.push(url);
        for (const [needle, handler] of routes) {
            if (url.includes(needle)) {
                const r = typeof handler === 'function' ? handler(url) : handler;
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

const PAGE_HTML = (title, tag) => `<!DOCTYPE html><html><head>`
    + `<meta name="created" content="2022-03-04T05:06:07Z"><title>${title}</title></head>`
    + `<body><div style="position:absolute;left:48px;top:90px">`
    + `<h1>${title}</h1><p data-tag="${tag}"><input type="checkbox"> ${title} task</p>`
    + `</div></body></html>`;

test('refreshGraphStatus reflects /api/graph/config', async (t) => {
    const { window, dispose } = createApp({
        onFetch: makeFetch([['/api/graph/config', { json: { enabled: true, connected: false } }]]),
    });
    t.after(dispose);

    await window.refreshGraphStatus();
    assert.equal(window.graphState.enabled, true);
    assert.equal(window.graphState.connected, false);
    assert.ok(!window.document.getElementById('graphImportSection').classList.contains('hidden'));
});

test('importSelectedGraphNotebook walks the tree via the proxy and adopts the adapter output', async (t) => {
    const log = [];
    const routes = [
        ['/api/graph/config', { json: { enabled: true, connected: true } }],
        ['/onenote/notebooks?', { json: { value: [{ id: 'nb1', displayName: 'Work' }] } }],
        // sections paginated
        ['/notebooks/nb1/sections', (url) => url.includes('skiptoken')
            ? { json: { value: [{ id: 's2', displayName: 'Sec B' }] } }
            : { json: { value: [{ id: 's1', displayName: 'Sec A' }], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/onenote/notebooks/nb1/sections?$skiptoken=abc' } }],
        ['/sections/s1/pages', { json: { value: [
            { id: 'p1', title: 'Alpha', level: 0, order: 0, createdDateTime: '2022-03-04T05:06:07Z', lastModifiedDateTime: '2022-03-05T00:00:00Z' },
            { id: 'p2', title: 'Beta', level: 1, order: 1 },
        ] } }],
        ['/sections/s2/pages', { json: { value: [{ id: 'p3', title: 'Gamma', level: 0, order: 0 }] } }],
        // p2 content is throttled once, then served
        ['/pages/p2/content', (() => { let n = 0; return () => (n++ === 0
            ? { status: 429, headers: { 'Retry-After': '0' }, text: '' }
            : { text: PAGE_HTML('Beta', 'to-do') }); })()],
        ['/pages/p1/content', { text: PAGE_HTML('Alpha', 'important') }],
        ['/pages/p3/content', { text: PAGE_HTML('Gamma', 'to-do') }],
    ];

    const { window, dispose } = createApp({ onFetch: makeFetch(routes, log) });
    t.after(dispose);
    seedNotebook(window);

    await window.refreshGraphStatus();            // -> connected, auto-loads notebooks
    assert.deepEqual(plain(window.graphState.notebooks).map((n) => n.id), ['nb1']);

    await window.importSelectedGraphNotebook();

    const notebooks = getState(window).notebooks;
    const imported = notebooks.at(-1);
    assert.equal(imported.name, 'Work');
    assert.deepEqual(plain(imported.sections).map((s) => s.name), ['Sec A', 'Sec B'], 'paginated sections both imported');
    for (const s of imported.sections) assert.match(s.color, /^#[0-9a-f]{6}$/);

    const secA = imported.sections[0];
    assert.deepEqual(plain(secA.pages).map((p) => [p.title, p.level || 0]), [['Alpha', 0], ['Beta', 1]]);
    assert.equal(secA.pages[0].createdAt, '2022-03-04T05:06:07.000Z', 'Graph date normalised');
    assert.match(secA.pages[1].blocks.map((b) => b.content).join(''), /data-tag="to-do"><input type="checkbox">\s*Beta task/);
    assert.ok(plain(secA.pages[0].tags).includes('Graph'));

    // 429 was retried, not fatal
    assert.equal(log.filter((u) => u.includes('/pages/p2/content')).length, 2);
    // active selection points at the freshly imported notebook
    assert.equal(getState(window).activeNotebookId, imported.id);
});

test('graphGet surfaces a 401 as "not connected"', async (t) => {
    const { window, dispose } = createApp({
        onFetch: makeFetch([['/api/graph/', { status: 401, json: { error: 'Not connected to Microsoft Graph' } }]]),
    });
    t.after(dispose);

    await assert.rejects(() => window.graphGet('/v1.0/me/onenote/notebooks'), /not connected/i);
    assert.equal(window.graphState.connected, false);
});

// --- Graph-hosted images / file attachments pulled in via the resource proxy --

const PNG_B64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
const RES = (id) => `https://graph.microsoft.com/v1.0/users/u/onenote/resources/${id}/$value`;
const PAGE_WITH_RESOURCES = `<!DOCTYPE html><html><head><title>Assets</title></head><body>`
    + `<div style="position:absolute;left:0px;top:0px">`
    + `<p>See the chart</p>`
    + `<img src="${RES('1-img')}" data-fullres-src="${RES('1-full')}" alt="chart">`
    + `<object data-attachment="report.pdf" data="${RES('2-pdf')}" type="application/pdf"></object>`
    + `</div></body></html>`;

function graphTreeRoutes(contentRoute, resourceRoute) {
    return [
        ['/api/graph/config', { json: { enabled: true, connected: true } }],
        ['/onenote/notebooks?', { json: { value: [{ id: 'nb1', displayName: 'Work' }] } }],
        ['/notebooks/nb1/sections', { json: { value: [{ id: 's1', displayName: 'Sec A' }] } }],
        ['/sections/s1/pages', { json: { value: [{ id: 'p1', title: 'Assets', level: 0, order: 0 }] } }],
        ['/pages/p1/content', contentRoute],
        ['/api/graph/resource?url=', resourceRoute],
    ];
}

test('Graph import inlines Graph-hosted images and stores linked files via the resource proxy', async (t) => {
    const PDF_B64 = Buffer.from('%PDF-1.4 tiny', 'utf8').toString('base64');
    const hits = [];
    const routes = graphTreeRoutes({ text: PAGE_WITH_RESOURCES }, (url) => {
        hits.push(url);
        return url.includes('2-pdf')
            ? { json: { type: 'application/pdf', data: PDF_B64 } }
            : { json: { type: 'image/png', data: PNG_B64 } };
    });

    const { window, dispose } = createApp({ onFetch: makeFetch(routes) });
    t.after(dispose);
    seedNotebook(window);

    await window.refreshGraphStatus();
    await window.importSelectedGraphNotebook();

    const page = getState(window).notebooks.at(-1).sections[0].pages[0];
    const html = page.blocks.map((b) => b.content).join('');

    assert.ok(html.includes(`src="data:image/png;base64,${PNG_B64}"`), 'image inlined as a data URL');
    assert.doesNotMatch(html, /graph\.microsoft\.com/, 'no live Graph URLs left in the content');
    assert.doesNotMatch(html, /data-fullres-src/, 'fullres hint dropped once inlined');

    assert.equal(plain(page.attachments).length, 1);
    assert.equal(page.attachments[0].name, 'report.pdf');
    assert.equal(page.attachments[0].type, 'application/pdf');
    assert.equal(window.atob(page.attachments[0].data), '%PDF-1.4 tiny');
    assert.doesNotMatch(html, /\(unavailable\)/);

    assert.equal(hits.filter((u) => u.startsWith('/api/graph/resource?url=')).length, 2, 'both resources fetched via the proxy');
});

test('Graph import survives a resource the proxy cannot fetch (falls back to the placeholder)', async (t) => {
    const routes = graphTreeRoutes({ text: PAGE_WITH_RESOURCES }, { status: 502, json: { error: 'Graph resource returned 500' } });

    const { window, dispose } = createApp({ onFetch: makeFetch(routes) });
    t.after(dispose);
    seedNotebook(window);

    await window.refreshGraphStatus();
    await window.importSelectedGraphNotebook(); // must not throw

    const page = getState(window).notebooks.at(-1).sections[0].pages[0];
    const html = page.blocks.map((b) => b.content).join('');

    assert.equal(plain(page.attachments).length, 0, 'unreachable file is not stored');
    assert.match(html, /report\.pdf \(unavailable\)/, 'the <object> becomes the unavailable placeholder');
});
