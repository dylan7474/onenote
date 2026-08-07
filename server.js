const http = require('http');
const fs = require('fs');
const path = require('path');

const port = Number(process.env.PORT || 3020);
const dataFile = process.env.DATA_FILE || path.join(__dirname, 'data', 'state.json');
const indexFile = path.join(__dirname, 'index.html');
const maxBodyBytes = 50 * 1024 * 1024;
const microsoftClientId = process.env.MS_CLIENT_ID;
const microsoftClientSecret = process.env.MS_CLIENT_SECRET;
const microsoftTenantId = process.env.MS_TENANT_ID || 'common';
const microsoftRedirectUri = process.env.MS_REDIRECT_URI || `http://localhost:${port}/api/microsoft/callback`;
const microsoftAuthority = `https://login.microsoftonline.com/${encodeURIComponent(microsoftTenantId)}/oauth2/v2.0`;
const microsoftScopes = 'openid profile offline_access User.Read Notes.Read';
const microsoftMissingSettings = [
    !microsoftClientId && 'MS_CLIENT_ID',
    !microsoftClientSecret && 'MS_CLIENT_SECRET'
].filter(Boolean);
const microsoftSessions = new Map();
const microsoftAuthRequests = new Map();

function randomId() {
    return require('crypto').randomBytes(24).toString('base64url');
}

function cookies(req) {
    return Object.fromEntries((req.headers.cookie || '').split(';').map(value => value.trim().split('=').map(decodeURIComponent)).filter(pair => pair.length === 2));
}

function getMicrosoftSession(req) {
    return microsoftSessions.get(cookies(req).onenote_ms_session);
}

function redirect(res, location, cookie) {
    const headers = { Location: location, 'Cache-Control': 'no-store' };
    if (cookie) headers['Set-Cookie'] = cookie;
    res.writeHead(302, headers);
    res.end();
}

async function tokenRequest(parameters) {
    const response = await fetch(`${microsoftAuthority}/token`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams(parameters)
    });
    const value = await response.json();
    if (!response.ok) throw new Error(value.error_description || value.error || 'Microsoft sign-in failed');
    return value;
}

async function accessToken(req) {
    const session = getMicrosoftSession(req);
    if (!session) throw Object.assign(new Error('Connect Microsoft OneNote first'), { status: 401 });
    if (session.expiresAt > Date.now() + 60_000) return session.accessToken;
    const token = await tokenRequest({
        client_id: microsoftClientId,
        client_secret: microsoftClientSecret,
        grant_type: 'refresh_token',
        refresh_token: session.refreshToken,
        redirect_uri: microsoftRedirectUri,
        scope: microsoftScopes
    });
    session.accessToken = token.access_token;
    session.refreshToken = token.refresh_token || session.refreshToken;
    session.expiresAt = Date.now() + token.expires_in * 1000;
    return session.accessToken;
}

async function graphGet(token, graphPath, accept = 'application/json') {
    const response = await fetch(graphPath.startsWith('https://') ? graphPath : `https://graph.microsoft.com/v1.0${graphPath}`, {
        headers: { Authorization: `Bearer ${token}`, Accept: accept }
    });
    if (!response.ok) {
        const detail = await response.text();
        throw Object.assign(new Error(`Microsoft Graph returned ${response.status}: ${detail.slice(0, 300)}`), { status: response.status });
    }
    return accept === 'application/json' ? response.json() : response.text();
}

async function graphCollection(token, graphPath) {
    const values = [];
    let next = graphPath;
    while (next) {
        const result = await graphGet(token, next);
        values.push(...(result.value || []));
        next = result['@odata.nextLink'];
    }
    return values;
}

function safeId(prefix) {
    return `${prefix}_${Date.now()}_${randomId().slice(0, 8)}`;
}

async function importMicrosoftNotebook(token, notebook) {
    const sections = await graphCollection(token, `/me/onenote/notebooks/${encodeURIComponent(notebook.id)}/sections?$select=id,displayName,createdDateTime`);
    const collectGroups = async (ownerPath, prefix = '') => {
        const groups = await graphCollection(token, `${ownerPath}/sectionGroups?$select=id,displayName`);
        const groupedSections = [];
        for (const group of groups) {
            const groupPrefix = `${prefix}${group.displayName || 'Section group'} / `;
            const groupPath = `/me/onenote/sectionGroups/${encodeURIComponent(group.id)}`;
            const children = await graphCollection(token, `${groupPath}/sections?$select=id,displayName,createdDateTime`);
            groupedSections.push(...children.map(section => ({ ...section, displayName: groupPrefix + section.displayName })));
            groupedSections.push(...await collectGroups(groupPath, groupPrefix));
        }
        return groupedSections;
    };
    sections.push(...await collectGroups(`/me/onenote/notebooks/${encodeURIComponent(notebook.id)}`));
    const importedSections = [];
    for (const [sectionIndex, section] of sections.entries()) {
        const pages = await graphCollection(token, `/me/onenote/sections/${encodeURIComponent(section.id)}/pages?$select=id,title,createdDateTime,lastModifiedDateTime,level,order&$orderby=order`);
        const importedPages = [];
        for (const page of pages) {
            const content = await graphGet(token, `/me/onenote/pages/${encodeURIComponent(page.id)}/content?includeIDs=true`, 'text/html');
            importedPages.push({
                id: safeId('page'),
                title: page.title || 'Untitled page',
                level: Math.max(0, Math.min(2, Number(page.level) || 0)),
                createdAt: page.createdDateTime || new Date().toISOString(),
                updatedAt: page.lastModifiedDateTime || page.createdDateTime || new Date().toISOString(),
                tags: ['Microsoft OneNote'],
                source: { provider: 'microsoft-graph', id: page.id },
                blocks: [{ id: safeId('blk'), content, x: 0, y: 0 }],
                attachments: [],
                inkData: null
            });
        }
        importedSections.push({ id: safeId('sec'), name: section.displayName || 'Untitled section', color: ['#8b5cf6', '#3b82f6', '#06b6d4', '#10b981', '#f59e0b', '#ec4899'][sectionIndex % 6], pages: importedPages });
    }
    return { id: safeId('nb'), name: notebook.displayName || 'Microsoft OneNote', sections: importedSections, source: { provider: 'microsoft-graph', id: notebook.id } };
}

function sendJson(res, status, value) {
    const body = value === undefined ? '' : JSON.stringify(value);
    res.writeHead(status, {
        'Content-Type': 'application/json; charset=utf-8',
        'Content-Length': Buffer.byteLength(body),
        'Cache-Control': 'no-store'
    });
    res.end(body);
}

async function readState(res) {
    try {
        const contents = await fs.promises.readFile(dataFile, 'utf8');
        sendJson(res, 200, JSON.parse(contents));
    } catch (error) {
        if (error.code === 'ENOENT') return sendJson(res, 404, { error: 'No saved state' });
        console.error('Unable to read saved state:', error);
        sendJson(res, 500, { error: 'Unable to read saved state' });
    }
}

function writeState(req, res) {
    let size = 0;
    const chunks = [];

    req.on('data', chunk => {
        size += chunk.length;
        if (size > maxBodyBytes) {
            sendJson(res, 413, { error: 'State is too large' });
            req.destroy();
            return;
        }
        chunks.push(chunk);
    });

    req.on('end', async () => {
        if (size > maxBodyBytes) return;
        try {
            const value = JSON.parse(Buffer.concat(chunks).toString('utf8'));
            if (!value || !Array.isArray(value.notebooks)) {
                return sendJson(res, 400, { error: 'Invalid application state' });
            }

            await fs.promises.mkdir(path.dirname(dataFile), { recursive: true });
            const temporaryFile = `${dataFile}.${process.pid}.${Date.now()}.${Math.random().toString(16).slice(2)}.tmp`;
            await fs.promises.writeFile(temporaryFile, JSON.stringify(value), 'utf8');
            await fs.promises.rename(temporaryFile, dataFile);
            sendJson(res, 200, { saved: true });
        } catch (error) {
            if (error instanceof SyntaxError) return sendJson(res, 400, { error: 'Invalid JSON' });
            console.error('Unable to save state:', error);
            sendJson(res, 500, { error: 'Unable to save state' });
        }
    });
}

function readJsonBody(req, limit = 1024 * 1024) {
    return new Promise((resolve, reject) => {
        let size = 0;
        const chunks = [];
        req.on('data', chunk => {
            size += chunk.length;
            if (size > limit) {
                reject(Object.assign(new Error('Request is too large'), { status: 413 }));
                req.destroy();
            } else chunks.push(chunk);
        });
        req.on('end', () => {
            try { resolve(chunks.length ? JSON.parse(Buffer.concat(chunks).toString('utf8')) : {}); }
            catch { reject(Object.assign(new Error('Invalid JSON'), { status: 400 })); }
        });
        req.on('error', reject);
    });
}

async function handleMicrosoftApi(req, res, url) {
    if (url.pathname === '/api/microsoft/status' && req.method === 'GET') {
        return sendJson(res, 200, {
            configured: microsoftMissingSettings.length === 0,
            connected: Boolean(getMicrosoftSession(req)),
            redirectUri: microsoftRedirectUri,
            missingSettings: microsoftMissingSettings
        });
    }
    if (url.pathname === '/api/microsoft/connect' && req.method === 'GET') {
        if (!microsoftClientId || !microsoftClientSecret) return sendJson(res, 503, { error: 'Microsoft import is not configured on this server' });
        const state = randomId();
        microsoftAuthRequests.set(state, Date.now() + 10 * 60 * 1000);
        const parameters = new URLSearchParams({ client_id: microsoftClientId, response_type: 'code', redirect_uri: microsoftRedirectUri, response_mode: 'query', scope: microsoftScopes, state });
        return redirect(res, `${microsoftAuthority}/authorize?${parameters}`);
    }
    if (url.pathname === '/api/microsoft/callback' && req.method === 'GET') {
        const expiresAt = microsoftAuthRequests.get(url.searchParams.get('state'));
        microsoftAuthRequests.delete(url.searchParams.get('state'));
        if (!expiresAt || expiresAt < Date.now() || url.searchParams.get('error')) return redirect(res, '/?microsoft=error');
        try {
            const token = await tokenRequest({ client_id: microsoftClientId, client_secret: microsoftClientSecret, grant_type: 'authorization_code', code: url.searchParams.get('code'), redirect_uri: microsoftRedirectUri, scope: microsoftScopes });
            const sessionId = randomId();
            microsoftSessions.set(sessionId, { accessToken: token.access_token, refreshToken: token.refresh_token, expiresAt: Date.now() + token.expires_in * 1000 });
            const secure = microsoftRedirectUri.startsWith('https:') ? '; Secure' : '';
            return redirect(res, '/?microsoft=connected', `onenote_ms_session=${encodeURIComponent(sessionId)}; Path=/; HttpOnly; SameSite=Lax${secure}`);
        } catch (error) {
            console.error('Microsoft callback failed:', error);
            return redirect(res, '/?microsoft=error');
        }
    }
    if (url.pathname === '/api/microsoft/disconnect' && req.method === 'POST') {
        microsoftSessions.delete(cookies(req).onenote_ms_session);
        return sendJson(res, 200, { disconnected: true });
    }
    if (url.pathname === '/api/microsoft/notebooks' && req.method === 'GET') {
        const token = await accessToken(req);
        const notebooks = await graphCollection(token, '/me/onenote/notebooks?$select=id,displayName,createdDateTime,lastModifiedDateTime&$orderby=displayName');
        return sendJson(res, 200, { notebooks: notebooks.map(({ id, displayName, createdDateTime, lastModifiedDateTime }) => ({ id, displayName, createdDateTime, lastModifiedDateTime })) });
    }
    if (url.pathname === '/api/microsoft/import' && req.method === 'POST') {
        const token = await accessToken(req);
        const body = await readJsonBody(req);
        const selectedIds = Array.isArray(body.notebookIds) ? new Set(body.notebookIds.map(String)) : new Set();
        if (!selectedIds.size) throw Object.assign(new Error('Select at least one notebook'), { status: 400 });
        const available = await graphCollection(token, '/me/onenote/notebooks?$select=id,displayName');
        const selected = available.filter(notebook => selectedIds.has(String(notebook.id)));
        if (selected.length !== selectedIds.size) throw Object.assign(new Error('One or more notebooks were not found'), { status: 404 });
        const notebooks = [];
        for (const notebook of selected) notebooks.push(await importMicrosoftNotebook(token, notebook));
        return sendJson(res, 200, { notebooks });
    }
    return false;
}

const server = http.createServer(async (req, res) => {
    const url = new URL(req.url, `http://${req.headers.host || 'localhost'}`);
    if (url.pathname === '/api/state' && req.method === 'GET') return readState(res);
    if (url.pathname === '/api/state' && req.method === 'PUT') return writeState(req, res);

    if (url.pathname.startsWith('/api/microsoft/')) {
        try {
            const handled = await handleMicrosoftApi(req, res, url);
            if (handled !== false) return;
        } catch (error) {
            console.error('Microsoft import failed:', error);
            return sendJson(res, error.status || 500, { error: error.message || 'Microsoft import failed' });
        }
    }

    if ((url.pathname === '/' || url.pathname === '/index.html') && req.method === 'GET') {
        res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
        return fs.createReadStream(indexFile).pipe(res);
    }

    sendJson(res, 404, { error: 'Not found' });
});

server.listen(port, () => {
    console.log(`OneNote Web listening on http://localhost:${port}`);
    console.log(`Saving application state to ${dataFile}`);
    if (microsoftMissingSettings.length) {
        console.warn(`Microsoft OneNote import disabled: set ${microsoftMissingSettings.join(' and ')} and restart the server.`);
    } else {
        console.log(`Microsoft OneNote callback URI: ${microsoftRedirectUri}`);
    }
});
