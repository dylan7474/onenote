const http = require('http');
const fs = require('fs');
const path = require('path');

const port = Number(process.env.PORT || 3020);
const dataFile = process.env.DATA_FILE || path.join(__dirname, 'data', 'state.json');
const indexFile = path.join(__dirname, 'index.html');
const maxBodyBytes = 50 * 1024 * 1024;

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

const server = http.createServer((req, res) => {
    const url = new URL(req.url, `http://${req.headers.host || 'localhost'}`);
    if (url.pathname === '/api/state' && req.method === 'GET') return readState(res);
    if (url.pathname === '/api/state' && req.method === 'PUT') return writeState(req, res);

    if ((url.pathname === '/' || url.pathname === '/index.html') && req.method === 'GET') {
        res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
        return fs.createReadStream(indexFile).pipe(res);
    }

    sendJson(res, 404, { error: 'Not found' });
});

server.listen(port, () => {
    console.log(`OneNote Web listening on http://localhost:${port}`);
    console.log(`Saving application state to ${dataFile}`);
});
