'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { createApp, seedNotebook } = require('./harness');

test('sanitizeHtml strips script tags and inline event handlers', (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);
    const out = window.sanitizeHtml(
        '<b>keep</b><script>window.x=1<\/script><p onclick="window.y=1">click</p><img src=z onerror="window.w=1">'
    );
    assert.match(out, /<b>keep<\/b>/);
    assert.doesNotMatch(out, /<script/i);
    assert.doesNotMatch(out, /onclick/i);
    assert.doesNotMatch(out, /onerror/i);
});

test('sanitizeHtml preserves the OneNote / Graph structure the app depends on', (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);

    assert.match(window.sanitizeHtml('<p data-tag="to-do">t</p>'), /data-tag="to-do"/);
    assert.match(window.sanitizeHtml('<p data-tag="to-do:completed">t</p>'), /data-tag="to-do:completed"/);
    assert.match(
        window.sanitizeHtml('<img data-render-src="r.png" data-fullres-src="f.png" src="data:image/png;base64,AAAA" width="12" height="8">'),
        /data-fullres-src="f\.png"/
    );
    assert.match(
        window.sanitizeHtml('<table class="onenote-table"><tr><td colspan="2" rowspan="3" style="background:#eee">c</td></tr></table>'),
        /colspan="2"[\s\S]*rowspan="3"/
    );
    assert.match(
        window.sanitizeHtml('<div style="position:absolute;left:48px;top:96px;width:300px">outline</div>'),
        /position:\s*absolute/
    );

    const attachment = window.sanitizeHtml(
        '<span class="inline-attachment" contenteditable="false" data-attachment-id="att_1" data-attachment="a.pdf">a.pdf</span>'
    );
    assert.match(attachment, /data-attachment-id="att_1"/);
    assert.match(attachment, /contenteditable="false"/);
});

test('sanitizeHtml keeps checkboxes but drops their inline handler', (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);
    const out = window.sanitizeHtml('<input type="checkbox" checked onclick="this.setAttribute(\'checked\', this.checked)"> item');
    assert.match(out, /type="checkbox"/);
    assert.doesNotMatch(out, /onclick/i);
});

test('sanitizeHtml is null/empty safe', (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);
    assert.equal(window.sanitizeHtml(''), '');
    assert.equal(window.sanitizeHtml(null), '');
    assert.equal(window.sanitizeHtml(undefined), '');
});

test('sanitizeImportedTree cleans block bodies for every notebook shape', (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);

    const fullBackup = { notebooks: [{ sections: [{ pages: [{ blocks: [{ content: '<i>ok</i><script>bad<\/script>' }] }] }] }] };
    window.sanitizeImportedTree(fullBackup);
    assert.doesNotMatch(fullBackup.notebooks[0].sections[0].pages[0].blocks[0].content, /script/i);

    const singleNotebook = { sections: [{ pages: [{ blocks: [{ content: '<img src=x onerror=go()>' }] }] }] };
    window.sanitizeImportedTree(singleNotebook);
    assert.doesNotMatch(singleNotebook.sections[0].pages[0].blocks[0].content, /onerror/i);

    const section = { pages: [{ blocks: [{ content: '<p onclick="go()">p</p>' }] }] };
    window.sanitizeImportedTree(section);
    assert.doesNotMatch(section.pages[0].blocks[0].content, /onclick/i);
});

test('render path sanitizes stored block content that predates sanitization', (t) => {
    const { window, dispose } = createApp();
    t.after(dispose);
    const { page } = seedNotebook(window);
    page.blocks = [{ id: 'blk_x', content: '<p>hi</p><script>window.__rendered_pwned=1<\/script>', x: 0, y: 0 }];

    window.renderActivePage();

    const editor = window.document.querySelector('#pageContentArea .editable-content');
    assert.ok(editor, 'editor rendered');
    assert.doesNotMatch(editor.innerHTML, /<script/i);
    assert.equal(window.__rendered_pwned, undefined);
});
