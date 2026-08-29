'use strict';

// Characterization-test harness.
//
// index.html ships its entire application as one inline <script>. Until that is
// split into modules (PLAN.md Phase 1), these tests load the page into jsdom,
// then evaluate the vendored libraries and that inline script inside the jsdom
// window so the real functions can be called directly.
//
// Tests here pin CURRENT behaviour, including the gaps REVIEW.md documents, so
// later refactors are caught. Where current behaviour is a known gap it is
// marked `{ todo: ... }` or called out in a comment with the REVIEW.md section.

const fs = require('fs');
const path = require('path');
const { JSDOM, VirtualConsole } = require('jsdom');

const REPO_ROOT = path.join(__dirname, '..');
const INDEX_HTML = fs.readFileSync(path.join(REPO_ROOT, 'index.html'), 'utf8');
const DOMPURIFY_SRC = fs.readFileSync(path.join(REPO_ROOT, 'vendor', 'dompurify.min.js'), 'utf8');
const JSZIP_SRC = fs.readFileSync(path.join(REPO_ROOT, 'vendor', 'jszip.min.js'), 'utf8');

// Node-side copy of the same JSZip, used only to pack fixture directories into a
// zip the app's importer can read.
const JSZipNode = require(path.join(REPO_ROOT, 'vendor', 'jszip.min.js'));

function extractMainScript(html) {
    const scripts = [...html.matchAll(/<script>([\s\S]*?)<\/script>/g)].map((m) => m[1]);
    const main = scripts.find((s) => s.includes('function sanitizeHtml'));
    if (!main) throw new Error('could not locate the main inline <script> in index.html');
    return main;
}

const MAIN_SCRIPT = extractMainScript(INDEX_HTML);

// Minimal 2D context stub. jsdom has no canvas backend; the ink layer only needs
// these calls to not throw during render.
function makeContext2dStub() {
    const noop = () => {};
    return {
        canvas: null,
        lineWidth: 1, strokeStyle: '#000', fillStyle: '#000', lineCap: 'round', lineJoin: 'round',
        globalAlpha: 1, globalCompositeOperation: 'source-over',
        clearRect: noop, fillRect: noop, strokeRect: noop,
        beginPath: noop, closePath: noop, moveTo: noop, lineTo: noop, arc: noop,
        stroke: noop, fill: noop, save: noop, restore: noop,
        translate: noop, scale: noop, rotate: noop, setTransform: noop,
        drawImage: noop, putImageData: noop,
        getImageData: () => ({ data: new Uint8ClampedArray(4), width: 1, height: 1 }),
        createImageData: () => ({ data: new Uint8ClampedArray(4), width: 1, height: 1 }),
        measureText: () => ({ width: 0 }),
    };
}

/**
 * Build a fresh, isolated app environment.
 *
 * @param {object} [opts]
 * @param {(url:string, init?:object) => Promise<any>} [opts.onFetch]
 *        Handler for the app's `fetch` calls. Defaults to a 404 (no saved
 *        state), so the app keeps whatever state a test seeds.
 * @returns {{ window: Window, dom: JSDOM, dispose: () => void }}
 */
function createApp({ onFetch } = {}) {
    const virtualConsole = new VirtualConsole();
    virtualConsole.on('jsdomError', (err) => {
        if (/Not implemented/.test(err.message)) return;
        throw err;
    });

    const dom = new JSDOM(INDEX_HTML, {
        url: 'http://localhost/',
        runScripts: 'outside-only',
        virtualConsole,
    });
    const { window } = dom;

    // Neutralise the app's own `window.load` bootstrap (it fetches and renders on
    // a timer); tests drive the functions they need explicitly. Registered in the
    // capture phase before the inline script runs, so it wins.
    window.addEventListener('load', (event) => event.stopImmediatePropagation(), true);

    // Deterministic browser-API stubs.
    // JSZip's async pipeline schedules work via setImmediate, which jsdom's
    // window realm lacks; without it zip-entry decoding never resolves.
    window.setImmediate = (fn, ...args) => setTimeout(fn, 0, ...args);
    window.clearImmediate = (id) => clearTimeout(id);

    window.lucide = { createIcons() {} };
    window.prompt = () => null;
    window.confirm = () => true;
    window.alert = () => {};
    window.scrollTo = () => {};

    // Clipboard stub. Records the last write on window.__clipboard.
    window.__clipboard = { items: null, text: null };
    window.ClipboardItem = function ClipboardItem(items) { this.items = items; };
    Object.defineProperty(window.navigator, 'clipboard', {
        configurable: true,
        value: {
            write: async (data) => { window.__clipboard.items = data; },
            writeText: async (text) => { window.__clipboard.text = text; },
        },
    });
    window.HTMLCanvasElement.prototype.getContext = function getContext() {
        const ctx = makeContext2dStub();
        ctx.canvas = this;
        return ctx;
    };
    window.fetch = async (url, init) => {
        if (typeof onFetch === 'function') return onFetch(url, init);
        return {
            ok: false,
            status: 404,
            json: async () => ({ error: 'No saved state' }),
            text: async () => '',
        };
    };

    window.eval(DOMPURIFY_SRC);
    window.eval(JSZIP_SRC);
    // The inline script declares its state with `let state = {...}`, which an
    // indirect eval does not expose as a global. Append accessors *in the same
    // eval* so tests can read and replace that binding.
    window.eval(
        MAIN_SCRIPT +
        '\n;window.__getState = () => state;' +
        '\n;window.__setState = (next) => { state = next; };'
    );

    const dispose = () => {
        try { window.clearTimeout(window.saveTimer); } catch { /* ignore */ }
        window.close();
    };

    return { dom, window, dispose };
}

/** Current application state object (the inline script's `let state`). */
function getState(window) {
    return window.__getState();
}

/**
 * Seed a minimal notebook/section/page tree and point the active-* ids at it.
 */
function seedNotebook(window, overrides = {}) {
    const page = {
        id: 'page_seed',
        title: 'Seed Page',
        createdAt: '2020-01-01T00:00:00.000Z',
        updatedAt: '2020-01-01T00:00:00.000Z',
        tags: [],
        blocks: [{ id: 'blk_seed', content: '<p>seed</p>', x: 0, y: 0 }],
        ...(overrides.page || {}),
    };
    const section = {
        id: 'sec_seed', name: 'Seed Section', color: '#8b5cf6', pages: [page],
        ...(overrides.section || {}),
    };
    const notebook = {
        id: 'nb_seed', name: 'Seed Notebook', sections: [section],
        ...(overrides.notebook || {}),
    };
    window.__setState({
        notebooks: [notebook],
        activeNotebookId: notebook.id,
        activeSectionId: section.id,
        activePageId: page.id,
        isDrawMode: false,
        darkMode: false,
        sectionLayout: 'vertical',
    });
    return { notebook, section, page };
}

function fixturePath(...parts) {
    return path.join(__dirname, 'fixtures', ...parts);
}

function readFixture(...parts) {
    return fs.readFileSync(fixturePath(...parts));
}

function walkDir(root, dir = root, out = []) {
    for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
        const full = path.join(dir, entry.name);
        if (entry.isDirectory()) walkDir(root, full, out);
        else out.push([path.relative(root, full).split(path.sep).join('/'), fs.readFileSync(full)]);
    }
    return out;
}

/** Pack a fixture directory into a zip Buffer. */
async function zipFixtureDir(...parts) {
    const zip = new JSZipNode();
    for (const [rel, buf] of walkDir(fixturePath(...parts))) zip.file(rel, buf);
    return zip.generateAsync({ type: 'nodebuffer' });
}

/**
 * Zip a fixture directory and run it through the app's `parseZipImport`. The
 * bytes are re-wrapped as a window-realm Uint8Array so the in-page JSZip
 * recognises them.
 */
async function importZipFixture(window, name, ...parts) {
    const buf = await zipFixtureDir(...parts);
    const file = new window.File([new window.Uint8Array(buf)], name, { type: 'application/zip' });
    return window.parseZipImport(file);
}

/** Wrap a fixture file as a window-realm File for the HTML/JSON importers. */
function fixtureFile(window, name, type, ...parts) {
    const text = readFixture(...parts).toString('utf8');
    return new window.File([text], name, { type });
}

module.exports = {
    createApp, seedNotebook, getState,
    fixturePath, readFixture, fixtureFile, zipFixtureDir, importZipFixture,
    REPO_ROOT,
};
