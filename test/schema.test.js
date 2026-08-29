'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { createApp, getState } = require('./harness');

// Objects returned from the jsdom realm have a different prototype; compare by value.
const plain = (v) => JSON.parse(JSON.stringify(v));

test('migrateState: a v0 blob is brought to the current schema', (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);

    const migrated = window.migrateState({
        notebooks: [{
            id: 'nb', name: 'N',
            sections: [{
                id: 's', name: 'S', color: '#112299',
                pages: [{ id: 'p', title: 'P', level: '5', blocks: null, tags: undefined }],
            }],
        }],
    });

    assert.equal(migrated.schemaVersion, 1);
    assert.equal(migrated.notebooks[0].sections[0].color, window.snapSectionColor('#112299'), 'colour snapped');
    const page = migrated.notebooks[0].sections[0].pages[0];
    assert.deepEqual([...page.blocks], [], 'missing blocks -> []');
    assert.deepEqual([...page.tags], [], 'missing tags -> []');
    assert.equal(page.level, 2, 'out-of-range level clamped');
});

test('migrateState: tolerates junk and is idempotent', (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);

    assert.deepEqual(plain(window.migrateState(undefined)), { schemaVersion: 1, notebooks: [] });
    assert.deepEqual(plain(window.migrateState({})), { schemaVersion: 1, notebooks: [] });

    const once = plain(window.migrateState({ notebooks: [{ sections: [{ color: '#010203', pages: [] }] }] }));
    const twice = plain(window.migrateState(once));
    assert.deepEqual(once, twice, 'running migrateState again changes nothing');
});

test('loadStateFromServer runs the loaded state through migrateState', async (t) => {
    const v0 = {
        notebooks: [{ id: 'nb', name: 'N', sections: [{ id: 's', name: 'S', color: '#101090', pages: [] }] }],
        activeNotebookId: 'nb',
    };
    const { window, dispose } = createApp({
        onFetch: async () => ({ ok: true, status: 200, json: async () => v0 }),
    });
    t.after(dispose);

    await window.loadStateFromServer();

    const state = getState(window);
    assert.equal(state.schemaVersion, 1);
    assert.equal(state.notebooks[0].sections[0].color, window.snapSectionColor('#101090'));
});
