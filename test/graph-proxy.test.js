'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const http = require('http');
const crypto = require('crypto');

const GRAPH_ENV = ['GRAPH_CLIENT_ID', 'GRAPH_CLIENT_SECRET', 'GRAPH_TENANT', 'GRAPH_REDIRECT_URI', 'GRAPH_SCOPES'];
function clearGraphEnv() {
    for (const key of GRAPH_ENV) delete process.env[key];
}

const graph = require('../graph');

test('graphTarget maps only /api/graph/{v1.0,beta}/* paths', () => {
    assert.equal(
        graph.graphTarget('/api/graph/v1.0/me/onenote/notebooks', '?$top=5'),
        'https://graph.microsoft.com/v1.0/me/onenote/notebooks?$top=5',
    );
    assert.equal(graph.graphTarget('/api/graph/beta/me', ''), 'https://graph.microsoft.com/beta/me');
    assert.equal(graph.graphTarget('/api/graph/login', ''), null);
    assert.equal(graph.graphTarget('/api/graph/v2.0/me', ''), null);
});

test('generatePkce: challenge is the base64url sha256 of the verifier', () => {
    const { codeVerifier, codeChallenge } = graph.generatePkce();
    assert.match(codeVerifier, /^[A-Za-z0-9_-]+$/);
    assert.equal(codeChallenge, crypto.createHash('sha256').update(codeVerifier).digest('base64url'));
});

test('authorizeUrl includes the app config and PKCE params', () => {
    clearGraphEnv();
    process.env.GRAPH_CLIENT_ID = 'test-client';
    process.env.GRAPH_TENANT = 'consumers';
    graph.reloadConfig();

    const url = new URL(graph.authorizeUrl({ state: 'STATE', codeChallenge: 'CHALL' }));
    assert.equal(url.origin + url.pathname, 'https://login.microsoftonline.com/consumers/oauth2/v2.0/authorize');
    assert.equal(url.searchParams.get('client_id'), 'test-client');
    assert.equal(url.searchParams.get('response_type'), 'code');
    assert.equal(url.searchParams.get('state'), 'STATE');
    assert.equal(url.searchParams.get('code_challenge'), 'CHALL');
    assert.equal(url.searchParams.get('code_challenge_method'), 'S256');
    assert.match(url.searchParams.get('scope'), /Notes\.ReadWrite/);
});

// --- HTTP-level, via the real server (no network to Microsoft) ---------------

function req(base, method, pathname, { headers = {}, body } = {}) {
    return new Promise((resolve, reject) => {
        const r = http.request(base + pathname, { method, headers }, (res) => {
            let data = '';
            res.on('data', (c) => { data += c; });
            res.on('end', () => resolve({ status: res.statusCode, headers: res.headers, body: data }));
        });
        r.on('error', reject);
        r.end(body);
    });
}

test('/api/graph over the server: disabled by default, enabled with GRAPH_CLIENT_ID', async (t) => {
    clearGraphEnv();
    delete require.cache[require.resolve('../server.js')];
    const { server } = require('../server.js');
    await new Promise((r) => server.listen(0, r));
    const base = `http://127.0.0.1:${server.address().port}`;
    t.after(() => { server.close(); clearGraphEnv(); graph.reloadConfig(); });

    // Disabled
    graph.reloadConfig();
    let res = await req(base, 'GET', '/api/graph/config');
    assert.equal(res.status, 200);
    assert.deepEqual(JSON.parse(res.body), { enabled: false, connected: false });
    assert.equal((await req(base, 'GET', '/api/graph/v1.0/me')).status, 501, 'proxy is closed when unconfigured');
    assert.equal((await req(base, 'GET', '/api/graph/login')).status, 501);

    // Enabled
    process.env.GRAPH_CLIENT_ID = 'test-client';
    graph.reloadConfig();

    res = await req(base, 'GET', '/api/graph/config');
    assert.deepEqual(JSON.parse(res.body), { enabled: true, connected: false });

    res = await req(base, 'GET', '/api/graph/login');
    assert.equal(res.status, 302);
    assert.match(res.headers.location, /^https:\/\/login\.microsoftonline\.com\/common\/oauth2\/v2\.0\/authorize\?/);
    assert.match(res.headers.location, /client_id=test-client/);
    assert.match(res.headers.location, /code_challenge=/, 'PKCE used when no client secret');
    assert.match(res.headers['set-cookie'][0], /^onenote_gsid=[A-Za-z0-9_-]+; Path=\/; HttpOnly; SameSite=Lax$/);

    // callback with no session / bad state
    assert.equal((await req(base, 'GET', '/api/graph/callback')).status, 400);
    assert.equal((await req(base, 'GET', '/api/graph/callback?code=x&state=y', {
        headers: { cookie: 'onenote_gsid=doesnotexist000000' },
    })).status, 400);

    // proxy without a connected session
    assert.equal((await req(base, 'GET', '/api/graph/v1.0/me')).status, 401);
});

test('/api/graph/resource: gated by config, session, and a graph.microsoft.com allowlist', async (t) => {
    clearGraphEnv();
    delete require.cache[require.resolve('../server.js')];
    const { server } = require('../server.js');
    await new Promise((r) => server.listen(0, r));
    const base = `http://127.0.0.1:${server.address().port}`;
    t.after(() => {
        server.close();
        clearGraphEnv();
        graph.reloadConfig();
        for (const k of [...graph._sessions.keys()]) graph._sessions.delete(k);
    });

    const RES = 'https://graph.microsoft.com/v1.0/users/u/onenote/resources/1-x/$value';

    // disabled entirely
    graph.reloadConfig();
    assert.equal((await req(base, 'GET', '/api/graph/resource?url=' + encodeURIComponent(RES))).status, 501);

    // enabled, but no connected session
    process.env.GRAPH_CLIENT_ID = 'test-client';
    graph.reloadConfig();
    assert.equal((await req(base, 'GET', '/api/graph/resource?url=' + encodeURIComponent(RES))).status, 401);

    // enabled + connected session: the URL must be a graph.microsoft.com
    // resource, and validation happens before any outbound request
    const sid = 'sess_' + 'a'.repeat(16);
    graph._sessions.set(sid, { accessToken: 'tok', expiresAt: Date.now() + 3600e3 });
    const withSession = (p) => req(base, 'GET', p, { headers: { cookie: `onenote_gsid=${sid}` } });

    assert.equal((await withSession('/api/graph/resource')).status, 400, 'missing url rejected');
    assert.equal((await withSession('/api/graph/resource?url=' + encodeURIComponent('https://evil.example/x'))).status, 400, 'off-host url rejected');
    assert.equal((await withSession('/api/graph/resource?url=' + encodeURIComponent('http://graph.microsoft.com/v1.0/x'))).status, 400, 'non-https rejected');
});
