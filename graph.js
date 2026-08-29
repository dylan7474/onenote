'use strict';

// Server-side Microsoft Graph OAuth 2.0 (authorization code) + a thin proxy for
// the OneNote API. Node built-ins only. The whole feature is disabled unless
// GRAPH_CLIENT_ID is set, so a default deploy is unaffected.
//
// Env:
//   GRAPH_CLIENT_ID       Azure app (client) id                      [required to enable]
//   GRAPH_CLIENT_SECRET   client secret -> confidential client       [optional; PKCE if unset]
//   GRAPH_TENANT          common | consumers | organizations | <id>  [default: common]
//   GRAPH_REDIRECT_URI    must match the app registration            [default: http://localhost:<PORT>/api/graph/callback]
//   GRAPH_SCOPES          space-separated                            [default: openid offline_access User.Read Notes.ReadWrite Notes.Create]
//
// Routes (all under /api/graph):
//   GET  /config      -> { enabled, connected }
//   GET  /login       -> 302 to the Microsoft authorize endpoint
//   GET  /callback    -> exchanges ?code, stores the token in an in-memory
//                        session keyed by the `onenote_gsid` cookie, 302 to /
//   POST /logout      -> drops the session token
//   *    /v1.0/...     -> proxied to https://graph.microsoft.com/v1.0/... with
//                        the session's bearer token (refreshed on expiry)

const https = require('https');
const crypto = require('crypto');

const DEFAULT_SCOPES = 'openid offline_access User.Read Notes.ReadWrite Notes.Create';
const GRAPH_HOST = 'graph.microsoft.com';

let config = null;

function reloadConfig() {
    const clientId = process.env.GRAPH_CLIENT_ID || '';
    const tenant = process.env.GRAPH_TENANT || 'common';
    const port = Number(process.env.PORT || 3020);
    config = {
        clientId,
        clientSecret: process.env.GRAPH_CLIENT_SECRET || '',
        tenant,
        redirectUri: process.env.GRAPH_REDIRECT_URI || `http://localhost:${port}/api/graph/callback`,
        scopes: process.env.GRAPH_SCOPES || DEFAULT_SCOPES,
        authorizeEndpoint: `https://login.microsoftonline.com/${tenant}/oauth2/v2.0/authorize`,
        tokenEndpoint: `https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`,
        enabled: Boolean(clientId),
    };
    return config;
}

function getConfig() {
    return config || reloadConfig();
}

function isEnabled() {
    return getConfig().enabled;
}

// sessionId -> { accessToken, refreshToken, expiresAt, codeVerifier?, state? }
const sessions = new Map();

function newSessionId() {
    return crypto.randomBytes(18).toString('base64url');
}

function generatePkce() {
    const codeVerifier = crypto.randomBytes(32).toString('base64url');
    const codeChallenge = crypto.createHash('sha256').update(codeVerifier).digest('base64url');
    return { codeVerifier, codeChallenge };
}

function authorizeUrl({ state, codeChallenge }) {
    const cfg = getConfig();
    const params = new URLSearchParams({
        client_id: cfg.clientId,
        response_type: 'code',
        redirect_uri: cfg.redirectUri,
        response_mode: 'query',
        scope: cfg.scopes,
        state,
    });
    if (codeChallenge) {
        params.set('code_challenge', codeChallenge);
        params.set('code_challenge_method', 'S256');
    }
    return `${cfg.authorizeEndpoint}?${params.toString()}`;
}

// Map /api/graph/v1.0/<rest> (or /api/graph/beta/<rest>) to the Graph URL.
function graphTarget(pathname, search) {
    const m = pathname.match(/^\/api\/graph\/(v1\.0|beta)\/(.*)$/);
    if (!m) return null;
    return `https://${GRAPH_HOST}/${m[1]}/${m[2]}${search || ''}`;
}

function parseCookies(header) {
    const out = {};
    (header || '').split(';').forEach(part => {
        const i = part.indexOf('=');
        if (i > 0) out[part.slice(0, i).trim()] = decodeURIComponent(part.slice(i + 1).trim());
    });
    return out;
}

function readBody(req, limit = 1 << 20) {
    return new Promise((resolve, reject) => {
        const chunks = [];
        let size = 0;
        req.on('data', c => { size += c.length; if (size > limit) { reject(new Error('body too large')); req.destroy(); } else chunks.push(c); });
        req.on('end', () => resolve(Buffer.concat(chunks)));
        req.on('error', reject);
    });
}

function httpsJson(method, urlString, { headers = {}, body } = {}) {
    return new Promise((resolve, reject) => {
        const u = new URL(urlString);
        const req = https.request({
            method,
            hostname: u.hostname,
            path: u.pathname + u.search,
            headers,
        }, (res) => {
            const chunks = [];
            res.on('data', c => chunks.push(c));
            res.on('end', () => {
                const text = Buffer.concat(chunks).toString('utf8');
                let json = null;
                try { json = text ? JSON.parse(text) : null; } catch { /* leave null */ }
                resolve({ status: res.statusCode, headers: res.headers, text, json });
            });
        });
        req.on('error', reject);
        if (body) req.write(body);
        req.end();
    });
}

async function exchangeToken(params) {
    const cfg = getConfig();
    const form = new URLSearchParams({
        client_id: cfg.clientId,
        redirect_uri: cfg.redirectUri,
        scope: cfg.scopes,
        ...params,
    });
    if (cfg.clientSecret) form.set('client_secret', cfg.clientSecret);
    const res = await httpsJson('POST', cfg.tokenEndpoint, {
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: form.toString(),
    });
    if (res.status !== 200 || !res.json || !res.json.access_token) {
        const detail = (res.json && (res.json.error_description || res.json.error)) || `token endpoint returned ${res.status}`;
        const err = new Error(detail);
        err.status = res.status;
        throw err;
    }
    return res.json;
}

function storeToken(sessionId, tok) {
    const record = sessions.get(sessionId) || {};
    record.accessToken = tok.access_token;
    if (tok.refresh_token) record.refreshToken = tok.refresh_token;
    record.expiresAt = Date.now() + (Number(tok.expires_in || 3600) - 60) * 1000;
    delete record.codeVerifier;
    delete record.state;
    sessions.set(sessionId, record);
}

async function validAccessToken(sessionId) {
    const record = sessions.get(sessionId);
    if (!record || !record.accessToken) return null;
    if (Date.now() < record.expiresAt) return record.accessToken;
    if (!record.refreshToken) return null;
    const tok = await exchangeToken({ grant_type: 'refresh_token', refresh_token: record.refreshToken });
    storeToken(sessionId, tok);
    return sessions.get(sessionId).accessToken;
}

function send(res, status, value, headers = {}) {
    const body = value === undefined ? '' : JSON.stringify(value);
    res.writeHead(status, { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store', ...headers });
    res.end(body);
}

function ensureSessionCookie(req, res) {
    const cookies = parseCookies(req.headers.cookie);
    let sessionId = cookies.onenote_gsid;
    if (!sessionId || !/^[A-Za-z0-9_-]{16,64}$/.test(sessionId)) {
        sessionId = newSessionId();
        res.setHeader('Set-Cookie', `onenote_gsid=${sessionId}; Path=/; HttpOnly; SameSite=Lax`);
    }
    return sessionId;
}

// Returns true if the request was a /api/graph/* route and has been handled.
async function handle(req, res, url) {
    if (!url.pathname.startsWith('/api/graph')) return false;
    const cfg = getConfig();

    if (url.pathname === '/api/graph/config' && req.method === 'GET') {
        const sessionId = parseCookies(req.headers.cookie).onenote_gsid;
        const record = sessionId && sessions.get(sessionId);
        send(res, 200, { enabled: cfg.enabled, connected: Boolean(record && record.accessToken) });
        return true;
    }

    if (!cfg.enabled) {
        send(res, 501, { error: 'Microsoft Graph is not configured on this server' });
        return true;
    }

    if (url.pathname === '/api/graph/login' && req.method === 'GET') {
        const sessionId = ensureSessionCookie(req, res);
        const state = crypto.randomBytes(16).toString('base64url');
        const record = sessions.get(sessionId) || {};
        record.state = state;
        if (!cfg.clientSecret) {
            const pkce = generatePkce();
            record.codeVerifier = pkce.codeVerifier;
            sessions.set(sessionId, record);
            res.writeHead(302, { Location: authorizeUrl({ state, codeChallenge: pkce.codeChallenge }) });
        } else {
            sessions.set(sessionId, record);
            res.writeHead(302, { Location: authorizeUrl({ state }) });
        }
        res.end();
        return true;
    }

    if (url.pathname === '/api/graph/callback' && req.method === 'GET') {
        const sessionId = parseCookies(req.headers.cookie).onenote_gsid;
        const record = sessionId && sessions.get(sessionId);
        const code = url.searchParams.get('code');
        const state = url.searchParams.get('state');
        if (!record || !code || !state || state !== record.state) {
            send(res, 400, { error: 'Invalid or expired sign-in attempt' });
            return true;
        }
        try {
            const tok = await exchangeToken({
                grant_type: 'authorization_code',
                code,
                ...(record.codeVerifier ? { code_verifier: record.codeVerifier } : {}),
            });
            storeToken(sessionId, tok);
            res.writeHead(302, { Location: '/' });
            res.end();
        } catch (error) {
            console.error('Graph token exchange failed:', error.message);
            send(res, 502, { error: 'Sign-in failed', detail: error.message });
        }
        return true;
    }

    if (url.pathname === '/api/graph/logout' && req.method === 'POST') {
        const sessionId = parseCookies(req.headers.cookie).onenote_gsid;
        if (sessionId) sessions.delete(sessionId);
        send(res, 200, { connected: false });
        return true;
    }

    const target = graphTarget(url.pathname, url.search);
    if (target) {
        const sessionId = parseCookies(req.headers.cookie).onenote_gsid;
        let accessToken;
        try {
            accessToken = sessionId && await validAccessToken(sessionId);
        } catch (error) {
            send(res, 502, { error: 'Token refresh failed', detail: error.message });
            return true;
        }
        if (!accessToken) {
            send(res, 401, { error: 'Not connected to Microsoft Graph' });
            return true;
        }
        let body;
        try {
            // Page HTML can carry inlined image/attachment data URLs.
            body = ['POST', 'PATCH', 'PUT'].includes(req.method) ? await readBody(req, 25 * 1024 * 1024) : undefined;
        } catch {
            send(res, 413, { error: 'Request body too large' });
            return true;
        }
        const headers = { Authorization: `Bearer ${accessToken}` };
        if (req.headers['content-type']) headers['Content-Type'] = req.headers['content-type'];
        const upstream = await httpsJson(req.method, target, { headers, body });
        res.writeHead(upstream.status, {
            'Content-Type': upstream.headers['content-type'] || 'application/json; charset=utf-8',
            'Cache-Control': 'no-store',
        });
        res.end(upstream.text);
        return true;
    }

    send(res, 404, { error: 'Not found' });
    return true;
}

module.exports = {
    handle,
    isEnabled,
    reloadConfig,
    getConfig,
    // exported for tests
    authorizeUrl,
    graphTarget,
    generatePkce,
    _sessions: sessions,
};
