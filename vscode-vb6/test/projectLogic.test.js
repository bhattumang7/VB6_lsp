'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('path');

const logic = require('../src/projectLogic');

const folder = path.join(path.parse(process.cwd()).root, 'work', 'app');
const p = (...parts) => path.join(folder, ...parts);

// ── chooseProject ───────────────────────────────────────────────────────────────

const exists = () => true;

test('a configured project is opened, resolved against the folder', () => {
    const choice = logic.chooseProject({
        folderPath: folder, configured: 'sub/App.vbp', candidates: [], autoLoad: false, exists,
    });
    assert.deepEqual(choice, { action: 'open', path: p('sub', 'App.vbp'), source: 'setting' });
});

test('an absolute configured project is kept', () => {
    const abs = path.join(path.parse(folder).root, 'elsewhere', 'X.vbg');
    const choice = logic.chooseProject({ folderPath: folder, configured: abs, candidates: [], autoLoad: true, exists });
    assert.deepEqual(choice, { action: 'open', path: abs, source: 'setting' });
});

test('a configured project that is not on disk is reported, not replaced by another', () => {
    const choice = logic.chooseProject({
        folderPath: folder, configured: 'Gone.vbp', candidates: [p('Other.vbp')], autoLoad: true, exists: () => false,
    });
    assert.deepEqual(choice, { action: 'none', reason: 'configuredMissing', path: p('Gone.vbp') });
});

test('whitespace in the setting counts as no setting', () => {
    const choice = logic.chooseProject({
        folderPath: folder, configured: '   ', candidates: [p('A.vbp')], autoLoad: true, exists,
    });
    assert.deepEqual(choice, { action: 'open', path: p('A.vbp'), source: 'auto' });
});

test('with auto-load off nothing is chosen', () => {
    const choice = logic.chooseProject({
        folderPath: folder, configured: '', candidates: [p('A.vbp')], autoLoad: false, exists,
    });
    assert.deepEqual(choice, { action: 'none', reason: 'disabled' });
});

test('the only project is opened', () => {
    const choice = logic.chooseProject({
        folderPath: folder, configured: '', candidates: [p('A.vbp')], autoLoad: true, exists,
    });
    assert.deepEqual(choice, { action: 'open', path: p('A.vbp'), source: 'auto' });
});

test('the only group wins over the projects it lists', () => {
    const choice = logic.chooseProject({
        folderPath: folder, configured: '', candidates: [p('A.vbp'), p('B.vbp'), p('All.vbg')], autoLoad: true, exists,
    });
    assert.deepEqual(choice, { action: 'open', path: p('All.vbg'), source: 'auto' });
});

test('several groups need a choice', () => {
    const choice = logic.chooseProject({
        folderPath: folder, configured: '', candidates: [p('A.vbg'), p('B.vbg')], autoLoad: true, exists,
    });
    assert.deepEqual(choice, { action: 'none', reason: 'several' });
});

test('several projects and no group need a choice', () => {
    const choice = logic.chooseProject({
        folderPath: folder, configured: '', candidates: [p('A.vbp'), p('B.vbp')], autoLoad: true, exists,
    });
    assert.deepEqual(choice, { action: 'none', reason: 'several' });
});

test('a folder without projects has nothing to open', () => {
    const choice = logic.chooseProject({ folderPath: folder, configured: '', candidates: [], autoLoad: true, exists });
    assert.deepEqual(choice, { action: 'none', reason: 'noProjects' });
});

// ── settings ────────────────────────────────────────────────────────────────────

test('a project inside the folder is stored relative with forward slashes', () => {
    assert.equal(logic.settingValue(folder, p('sub', 'dir', 'App.vbp')), 'sub/dir/App.vbp');
});

test('a project outside the folder is stored absolute', () => {
    const outside = path.join(path.parse(folder).root, 'other', 'App.vbp');
    assert.equal(logic.settingValue(folder, outside), outside);
});

test('a sibling whose name starts with two dots is still inside', () => {
    assert.equal(logic.settingValue(folder, p('..hidden', 'A.vbp')), '..hidden/A.vbp');
});

test('file kind helpers ignore case', () => {
    assert.equal(logic.isProjectFile('X.VBP'), true);
    assert.equal(logic.isProjectFile('X.vbg'), true);
    assert.equal(logic.isProjectFile('X.bas'), false);
    assert.equal(logic.isGroupFile('X.VBG'), true);
    assert.equal(logic.isGroupFile('X.vbp'), false);
});

// ── status bar ──────────────────────────────────────────────────────────────────

const loaded = {
    state: 'loaded', path: p('App.vbp'), name: 'App', isGroup: false, projectCount: 1, memberCount: 40,
    missingCount: 0, errorCount: 0, warningCount: 0, informationCount: 0, message: null,
};

test('status bar: no project', () => {
    assert.deepEqual(logic.statusBarView({ state: 'none' }), {
        text: '$(folder-opened) VB6: No project',
        tooltip: 'Open a VB6 project (.vbp) or group (.vbg)',
        warning: false,
        command: 'vb6.project.select',
    });
});

test('status bar: loading', () => {
    assert.deepEqual(logic.statusBarView({ state: 'loading', path: p('App.vbp') }), {
        text: '$(sync~spin) VB6: Loading App.vbp',
        tooltip: `Loading ${p('App.vbp')}`,
        warning: false,
        command: 'vb6.project.select',
    });
});

test('status bar: a clean project', () => {
    assert.deepEqual(logic.statusBarView(loaded), {
        text: '$(project) App',
        tooltip: `Project: ${p('App.vbp')}\n40 files`,
        warning: false,
        command: 'vb6.project.select',
    });
});

test('status bar: a group with problems', () => {
    const status = {
        ...loaded, name: 'All', isGroup: true, projectCount: 3, memberCount: 1, missingCount: 1,
        errorCount: 2, warningCount: 1, informationCount: 4, message: 'note',
    };
    assert.deepEqual(logic.statusBarView(status), {
        text: '$(project) All · 3 projects $(warning) 3',
        tooltip: `Group: ${p('App.vbp')}\n1 file\n1 missing\n2 errors\n1 warning\n4 notes\nnote`,
        warning: true,
        command: 'vb6.project.select',
    });
});

test('status bar: a group of one project is singular', () => {
    const view = logic.statusBarView({ ...loaded, isGroup: true, projectCount: 1 });
    assert.equal(view.text, '$(project) App · 1 project');
});

test('status bar: a failed open', () => {
    assert.deepEqual(logic.statusBarView({ state: 'error', path: p('Bad.vbg'), message: 'no header' }), {
        text: '$(error) VB6: cannot open Bad.vbg',
        tooltip: 'no header',
        warning: true,
        command: 'vb6.project.select',
    });
});

test('status bar: a failed open without a reason', () => {
    assert.equal(logic.statusBarView({ state: 'error', path: 'x.vbp' }).tooltip, 'The project could not be opened');
});

// ── load summary ────────────────────────────────────────────────────────────────

test('no summary for a clean load', () => {
    assert.equal(logic.loadSummary(loaded), null);
});

test('no summary while not loaded', () => {
    assert.equal(logic.loadSummary({ state: 'loading' }), null);
});

test('the summary counts missing files apart from other errors', () => {
    const status = { ...loaded, missingCount: 2, errorCount: 3, warningCount: 1, informationCount: 2 };
    assert.equal(
        logic.loadSummary(status),
        'Loaded "App": 2 of 40 files missing, 1 other error, 1 warning, 2 notes.'
    );
});

test('the summary of notes alone', () => {
    assert.equal(logic.loadSummary({ ...loaded, informationCount: 1 }), 'Loaded "App": 1 note.');
});

// ── tree ────────────────────────────────────────────────────────────────────────

function unit(name, overrides) {
    const dir = p(name);
    return {
        path: path.join(dir, `${name}.vbp`),
        uri: `file:///work/app/${name}/${name}.vbp`,
        name, kind: 'exe', listed: true, startup: true, members: [], references: [],
        ...overrides,
    };
}

function member(unitName, name, extension, overrides) {
    return {
        name, extension, path: p(unitName, `${name}.${extension}`),
        uri: `file:///work/app/${unitName}/${name}.${extension}`, line: 3, missing: false,
        ...overrides,
    };
}

test('no tree without a project', () => {
    assert.deepEqual(logic.buildTree(null), []);
});

test('a project is grouped like the VB6 project explorer', () => {
    const u = unit('App', {
        members: [
            member('App', 'modB', 'bas'),
            member('App', 'frmMain', 'frm'),
            member('App', 'modA', 'bas'),
            member('App', 'clsThing', 'cls', { missing: true }),
            member('App', 'readme', 'txt'),
        ],
        references: [
            { label: 'OLE Automation', kind: 'typeLibrary', status: 'loaded', line: 2, detail: '{x}' },
            { label: 'ADO', kind: 'typeLibrary', status: 'notLoaded', line: 3, detail: '{y}' },
        ],
    });
    const info = { path: u.path, uri: u.uri, name: 'App', isGroup: false, projects: [u], problems: [] };

    const tree = logic.buildTree(info);

    assert.deepEqual(tree, [
        {
            id: u.uri, kind: 'project', label: 'App', description: 'startup · exe', tooltip: u.path,
            icon: 'project', warning: false, collapsed: false, open: { uri: u.uri },
            children: [
                {
                    id: `${u.uri}/folder/frm`, kind: 'folder', label: 'Forms', description: '1', icon: 'folder',
                    warning: false, collapsed: false,
                    children: [{
                        id: `${u.uri}/file:///work/app/App/frmMain.frm`, kind: 'member', label: 'frmMain',
                        description: 'frmMain.frm', tooltip: p('App', 'frmMain.frm'), icon: 'window', warning: false,
                        open: { uri: 'file:///work/app/App/frmMain.frm' },
                    }],
                },
                {
                    id: `${u.uri}/folder/bas`, kind: 'folder', label: 'Modules', description: '2', icon: 'folder',
                    warning: false, collapsed: false,
                    children: [
                        {
                            id: `${u.uri}/file:///work/app/App/modA.bas`, kind: 'member', label: 'modA',
                            description: 'modA.bas', tooltip: p('App', 'modA.bas'), icon: 'symbol-namespace',
                            warning: false, open: { uri: 'file:///work/app/App/modA.bas' },
                        },
                        {
                            id: `${u.uri}/file:///work/app/App/modB.bas`, kind: 'member', label: 'modB',
                            description: 'modB.bas', tooltip: p('App', 'modB.bas'), icon: 'symbol-namespace',
                            warning: false, open: { uri: 'file:///work/app/App/modB.bas' },
                        },
                    ],
                },
                {
                    id: `${u.uri}/folder/cls`, kind: 'folder', label: 'Class Modules', description: '1', icon: 'folder',
                    warning: false, collapsed: false,
                    children: [{
                        id: `${u.uri}/file:///work/app/App/clsThing.cls`, kind: 'member', label: 'clsThing',
                        description: 'missing', tooltip: `${p('App', 'clsThing.cls')} (file not found)`,
                        icon: 'warning', warning: true, open: null,
                    }],
                },
                {
                    id: `${u.uri}/folder/other`, kind: 'folder', label: 'Other files', description: '1', icon: 'folder',
                    warning: false, collapsed: true,
                    children: [{
                        id: `${u.uri}/file:///work/app/App/readme.txt`, kind: 'member', label: 'readme',
                        description: 'readme.txt', tooltip: p('App', 'readme.txt'), icon: 'file', warning: false,
                        open: { uri: 'file:///work/app/App/readme.txt' },
                    }],
                },
                {
                    id: `${u.uri}/folder/references`, kind: 'folder', label: 'References', description: '2',
                    icon: 'references', warning: false, collapsed: true,
                    children: [
                        {
                            id: `${u.uri}/ref/2`, kind: 'reference', label: 'OLE Automation', description: 'loaded',
                            tooltip: 'OLE Automation\n{x}', icon: 'library', warning: false,
                            open: { uri: u.uri, line: 2 },
                        },
                        {
                            id: `${u.uri}/ref/3`, kind: 'reference', label: 'ADO', description: 'not loaded',
                            tooltip: 'ADO\n{y}', icon: 'warning', warning: true, open: { uri: u.uri, line: 3 },
                        },
                    ],
                },
            ],
        },
    ]);
});

test('a member without a name in the project is labelled by its file', () => {
    const u = unit('App', { members: [member('App', 'frmMain', 'frm', { name: '' })] });
    const tree = logic.buildTree({ path: u.path, uri: u.uri, name: 'App', isGroup: false, projects: [u], problems: [] });
    assert.equal(tree[0].children[0].children[0].label, 'frmMain.frm');
});

test('a group is the root, with referenced-only projects apart', () => {
    const app = unit('App');
    const lib = unit('Lib', { startup: false, kind: 'oledll' });
    const hidden = unit('Hidden', { startup: false, listed: false, kind: 'oledll' });
    const info = {
        path: p('All.vbg'), uri: 'file:///work/app/All.vbg', name: 'All', isGroup: true,
        projects: [app, lib, hidden], problems: [],
    };

    const tree = logic.buildTree(info);

    assert.equal(tree.length, 1);
    const root = tree[0];
    assert.deepEqual(
        { kind: root.kind, label: root.label, description: root.description, icon: root.icon, open: root.open },
        { kind: 'group', label: 'All', description: 'group', icon: 'root-folder', open: { uri: 'file:///work/app/All.vbg' } }
    );
    assert.deepEqual(root.children.map((c) => [c.kind, c.label, c.description]), [
        ['project', 'App', 'startup · exe'],
        ['project', 'Lib', 'oledll'],
        ['folder', 'Referenced projects', '1'],
    ]);
    assert.deepEqual(root.children[2].children.map((c) => c.label), ['Hidden']);
});

test('a single project shows its referenced-only projects in a folder next to it', () => {
    const app = unit('App');
    const hidden = unit('Lib', { startup: false, listed: false });
    const tree = logic.buildTree({ path: app.path, uri: app.uri, name: 'App', isGroup: false, projects: [app, hidden], problems: [] });
    assert.deepEqual(tree.map((n) => n.label), ['App', 'Referenced projects']);
});

test('reference statuses are marked', () => {
    const u = unit('App', {
        references: [
            { label: 'P1', kind: 'project', status: 'projectLoaded', line: 1, detail: 'd' },
            { label: 'P2', kind: 'project', status: 'projectMissing', line: 2, detail: 'd' },
            { label: 'P3', kind: 'project', status: 'projectCircular', line: 3, detail: 'd' },
            { label: 'P4', kind: 'project', status: 'projectNotExposed', line: 4, detail: 'd' },
            { label: 'P5', kind: 'control', status: 'somethingNew', line: 5, detail: 'd' },
        ],
    });
    const tree = logic.buildTree({ path: u.path, uri: u.uri, name: 'App', isGroup: false, projects: [u], problems: [] });
    const refs = tree[0].children[0].children.map((r) => [r.description, r.icon, r.warning]);
    assert.deepEqual(refs, [
        ['project', 'project', false],
        ['project not found', 'error', true],
        ['circular reference', 'error', true],
        ['exposes no types', 'warning', true],
        ['somethingNew', 'question', false],
    ]);
});

// ── picker ──────────────────────────────────────────────────────────────────────

test('the picker lists groups first, then projects, and marks the open one', () => {
    const items = logic.pickerItems([p('b', 'B.vbp'), p('All.vbg'), p('a', 'A.vbp')], p('b', 'B.vbp'));
    assert.deepEqual(items, [
        { path: p('All.vbg'), label: '$(root-folder) All.vbg', description: 'group', detail: folder },
        { path: p('a', 'A.vbp'), label: '$(project) A.vbp', description: 'project', detail: p('a') },
        { path: p('b', 'B.vbp'), label: '$(project) B.vbp', description: 'open', detail: p('b') },
    ]);
});

test('the picker marks nothing open when no project is', () => {
    const items = logic.pickerItems([p('A.vbp')], null);
    assert.equal(items[0].description, 'project');
});

// ── project page ────────────────────────────────────────────────────────────────

const page = (extra) =>
    logic.renderProjectPage({ nonce: 'N0NCE', cspSource: 'vscode-resource:', loaded: false, ...extra });

test('escaping covers every html special character', () => {
    assert.equal(logic.esc(`<a href="x">'&'</a>`), '&lt;a href=&quot;x&quot;&gt;&#39;&amp;&#39;&lt;/a&gt;');
});

test('the page of an unloaded project offers to load it', () => {
    const u = unit('App', {
        members: [member('App', 'modMain', 'bas'), member('App', 'modGone', 'bas', { missing: true })],
        references: [{ label: 'ADO', kind: 'typeLibrary', status: 'notLoaded', line: 5, detail: '{y}' }],
    });
    const info = {
        path: u.path, uri: u.uri, name: 'App', isGroup: false, projects: [u],
        problems: [{ file: u.path, line: 4, severity: 'error', message: 'File not found: x' }],
    };

    const html = page({ info });

    assert.match(html, /<h1>App <span class="badge">Project<\/span> <span class="badge">not loaded<\/span><\/h1>/);
    assert.match(html, /<button class="primary" data-action="load">Load Project<\/button>/);
    assert.match(html, /<button data-action="openText">Edit as Text<\/button>/);
    assert.match(html, /<h2>Problems \(1\)<\/h2>/);
    assert.match(html, /<li class="error"><span class="sev">Error<\/span> <a href="#" data-action="openLine" data-path="[^"]*" data-line="4">App\.vbp:4<\/a> File not found: x<\/li>/);
    assert.match(html, /<tr class=""><td><a href="#" data-action="openFile" data-uri="file:\/\/\/work\/app\/App\/modMain.bas">modMain<\/a><\/td><td>\.bas<\/td>/);
    assert.match(html, /<tr class="missing"><td>modGone<\/td><td>\.bas<\/td><td>[^<]*<\/td><td>missing<\/td><\/tr>/);
    assert.match(html, /<tr class="missing"><td><a href="#" data-action="openLine" data-uri="[^"]*" data-line="5">ADO<\/a><\/td><td>typeLibrary<\/td><td>\{y\}<\/td><td>not loaded<\/td><\/tr>/);
    assert.match(html, /<span class="badge">startup<\/span>/);
});

test('the page of the loaded project offers reload and close', () => {
    const u = unit('App');
    const html = page({ info: { path: u.path, uri: u.uri, name: 'App', isGroup: false, projects: [u], problems: [] }, loaded: true });
    assert.match(html, /<span class="badge ok">loaded<\/span>/);
    assert.match(html, /<button data-action="reload">Reload<\/button><button data-action="close">Close Project<\/button>/);
    assert.doesNotMatch(html, /data-action="load"/);
    assert.doesNotMatch(html, /Problems \(/);
});

test('the page of a group loads a group and marks referenced-only projects', () => {
    const app = unit('App');
    const lib = unit('Lib', { startup: false, listed: false });
    const html = page({
        info: { path: p('All.vbg'), uri: 'file:///work/app/All.vbg', name: 'All', isGroup: true, projects: [app, lib], problems: [] },
    });
    assert.match(html, /<span class="badge">Project group<\/span>/);
    assert.match(html, /Load Group<\/button>/);
    assert.match(html, /<h2>Projects<\/h2>/);
    assert.match(html, /<span class="badge">referenced only<\/span>/);
});

test('the page of an unreadable file shows why', () => {
    const html = page({ info: null, error: 'bad <header>' });
    assert.match(html, /<p class="error">bad &lt;header&gt;<\/p>/);
    assert.match(html, /data-action="openText"/);
    assert.doesNotMatch(html, /data-action="load"/);
});

test('the page of an unreadable file without a reason says so', () => {
    assert.match(page({ info: null }), /The file could not be read\./);
});

test('the page escapes names that come from the project file', () => {
    const u = unit('App', { name: '<img src=x onerror=alert(1)>' });
    const html = page({ info: { path: u.path, uri: u.uri, name: u.name, isGroup: false, projects: [u], problems: [] } });
    assert.doesNotMatch(html, /<img/);
    assert.match(html, /&lt;img src=x onerror=alert\(1\)&gt;/);
});

test('the page only runs its own script', () => {
    const html = page({ info: null });
    assert.match(html, /content="default-src 'none'; style-src vscode-resource: 'unsafe-inline'; script-src 'nonce-N0NCE';"/);
    assert.match(html, /<script nonce="N0NCE">/);
});

test('isInside tells folders apart from lookalikes', () => {
    assert.equal(logic.isInside(folder, p('a.vbp')), true);
    assert.equal(logic.isInside(folder, p('..hidden', 'a.vbp')), true);
    assert.equal(logic.isInside(folder, folder), false);
    assert.equal(logic.isInside(folder, path.dirname(folder)), false);
    assert.equal(logic.isInside(folder, path.join(path.dirname(folder), 'app2', 'a.vbp')), false);
});
