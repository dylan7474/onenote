'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const http = require('http');
const fs = require('fs');
const os = require('os');
const path = require('path');

function request(base, method, pathname, body) {
    return new Promise((resolve, reject) => {
        const req = http.request(base + pathname, { method, headers: { 'content-type': 'application/json' } }, (res) => {
            let data = '';
            res.on('data', (c) => { data += c; });
            res.on('end', () => resolve({ status: res.statusCode, body: data }));
        });
        req.on('error', reject);
        if (body !== undefined) req.end(typeof body === 'string' ? body : JSON.stringify(body));
        else req.end();
    });
}

test('server: /api/state validation and schema-version gate', async (t) => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'onenote-srv-'));
    process.env.DATA_FILE = path.join(dir, 'state.json');
    delete require.cache[require.resolve('../server.js')];
    const { server, SCHEMA_VERSION } = require('../server.js');
    await new Promise((r) => server.listen(0, r));
    const base = `http://127.0.0.1:${server.address().port}`;
    t.after(() => { server.close(); fs.rmSync(dir, { recursive: true, force: true }); });

    const put = (b) => request(base, 'PUT', '/api/state', b);

    assert.equal((await request(base, 'GET', '/api/state')).status, 404, 'no state yet');
    assert.equal((await put({ notebooks: [], schemaVersion: SCHEMA_VERSION })).status, 200);
    assert.equal((await put({ notebooks: [] })).status, 200, 'missing schemaVersion is allowed');
    assert.equal((await put({ notebooks: 'nope' })).status, 400);
    assert.equal((await put({ notebooks: [], schemaVersion: -1 })).status, 400);
    assert.equal((await put({ notebooks: [], schemaVersion: 'x' })).status, 400);
    assert.equal((await put({ notebooks: [], schemaVersion: SCHEMA_VERSION + 1 })).status, 409,
        'a state newer than the server is refused');
    assert.equal((await put('{ not json')).status, 400);

    // Last accepted write persists and reads back.
    await put({ notebooks: [{ id: 'nb', name: 'N', sections: [] }], schemaVersion: SCHEMA_VERSION });
    const got = await request(base, 'GET', '/api/state');
    assert.equal(got.status, 200);
    const state = JSON.parse(got.body);
    assert.equal(state.schemaVersion, SCHEMA_VERSION);
    assert.equal(state.notebooks[0].id, 'nb');
});
