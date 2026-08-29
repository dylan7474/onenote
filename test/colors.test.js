'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { createApp, seedNotebook, getState } = require('./harness');

test('nearestSectionColor snaps arbitrary hex to a named palette entry', (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);

    assert.equal(window.sectionColorName('#8b5cf6'), 'Purple', 'exact palette hex');
    assert.equal(window.sectionColorName('#3c81f6'), 'Blue', 'near a palette hex');
    assert.equal(window.sectionColorName('#000000'), 'Black');
    assert.equal(window.sectionColorName('#22c55e'), 'Green');
    assert.equal(window.sectionColorName('fff'), 'Gray', '3-digit hex, white -> nearest is Gray');

    // Unparseable input falls back to the first palette entry, never throws.
    assert.equal(window.sectionColorName('not-a-color'), 'Purple');
    assert.equal(window.snapSectionColor(undefined), '#8b5cf6');

    // snapSectionColor always returns a canonical palette hex.
    assert.match(window.snapSectionColor('#123456'), /^#[0-9a-f]{6}$/);
});

test('new sections get deterministic palette colours (no Math.random)', (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);
    seedNotebook(window);
    window.prompt = () => 'A';
    window.promptCreateSection();
    window.prompt = () => 'B';
    window.promptCreateSection();

    const sections = getState(window).notebooks[0].sections;
    const colors = sections.map((s) => s.color);
    assert.equal(new Set(colors).size, colors.length, 'each section a distinct colour');
    for (const c of colors) assert.match(c, /^#[0-9a-f]{6}$/);
});

test('JSON import snaps a foreign section colour to the palette', async (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);
    seedNotebook(window);

    const foreign = JSON.stringify({
        id: 'nb_f', name: 'Foreign',
        sections: [{ id: 's', name: 'S', color: '#112299', pages: [
            { id: 'p', title: 'P', createdAt: '2021-01-01T00:00:00.000Z', updatedAt: '2021-01-01T00:00:00.000Z', tags: [], blocks: [{ id: 'b', content: '<p>x</p>', x: 0, y: 0 }] },
        ] }],
    });
    await window.parseJsonImport(new window.File([foreign], 'f.json', { type: 'application/json' }));

    const imported = getState(window).notebooks.at(-1);
    assert.equal(imported.sections[0].color, window.snapSectionColor('#112299'));
    assert.notEqual(imported.sections[0].color, '#112299', 'raw hex replaced');
});
