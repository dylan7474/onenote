const http = require('http');
const fs = require('fs');
const path = require('path');
const graph = require('./graph');

const port = Number(process.env.PORT || 3020);
const dataFile = process.env.DATA_FILE || path.join(__dirname, 'data', 'state.json');
const indexFile = path.join(__dirname, 'index.html');
const vendorDir = path.join(__dirname, 'vendor');
const maxBodyBytes = 50 * 1024 * 1024;

// Highest client state-schema version this build understands. Keep in sync with
// SCHEMA_VERSION in index.html. The client owns migrations; the server only
// refuses to persist a state stamped newer than it knows, so a rolled-back
// deploy can't overwrite the data file with a format it might not preserve.
const SCHEMA_VERSION = 1;

const vendorContentTypes = {
    '.js': 'application/javascript; charset=utf-8',
    '.css': 'text/css; charset=utf-8',
    '.map': 'application/json; charset=utf-8'
};

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
            if (value.schemaVersion != null) {
                const version = Number(value.schemaVersion);
                if (!Number.isInteger(version) || version < 0) {
                    return sendJson(res, 400, { error: 'Invalid schemaVersion' });
                }
                if (version > SCHEMA_VERSION) {
                    return sendJson(res, 409, { error: 'State schema is newer than this server' });
                }
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

async function serveVendorAsset(pathname, res) {
    const relative = path.normalize(decodeURIComponent(pathname.slice('/vendor/'.length)));
    const resolved = path.join(vendorDir, relative);
    if (resolved !== vendorDir && !resolved.startsWith(vendorDir + path.sep)) {
        return sendJson(res, 403, { error: 'Forbidden' });
    }

    const contentType = vendorContentTypes[path.extname(resolved).toLowerCase()];
    if (!contentType) return sendJson(res, 404, { error: 'Not found' });

    try {
        const data = await fs.promises.readFile(resolved);
        res.writeHead(200, {
            'Content-Type': contentType,
            'Content-Length': data.length,
            'Cache-Control': 'public, max-age=86400'
        });
        res.end(data);
    } catch (error) {
        if (error.code === 'ENOENT' || error.code === 'EISDIR') return sendJson(res, 404, { error: 'Not found' });
        console.error('Unable to read vendor asset:', error);
        sendJson(res, 500, { error: 'Unable to read asset' });
    }
}

const server = http.createServer((req, res) => {
    const url = new URL(req.url, `http://${req.headers.host || 'localhost'}`);
    if (url.pathname === '/api/state' && req.method === 'GET') return readState(res);
    if (url.pathname === '/api/state' && req.method === 'PUT') return writeState(req, res);
    if (url.pathname.startsWith('/api/graph')) {
        graph.handle(req, res, url).catch(error => {
            console.error('Graph handler error:', error);
            if (!res.headersSent) sendJson(res, 500, { error: 'Graph handler error' });
        });
        return;
    }
    if (url.pathname.startsWith('/vendor/') && req.method === 'GET') return serveVendorAsset(url.pathname, res);

    if ((url.pathname === '/' || url.pathname === '/index.html') && req.method === 'GET') {
        res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
        return fs.createReadStream(indexFile).pipe(res);
    }

    sendJson(res, 404, { error: 'Not found' });
});

if (require.main === module) {
    server.listen(port, () => {
        console.log(`OneNote Web listening on http://localhost:${port}`);
        console.log(`Saving application state to ${dataFile}`);
    });
}

module.exports = { server, SCHEMA_VERSION };
