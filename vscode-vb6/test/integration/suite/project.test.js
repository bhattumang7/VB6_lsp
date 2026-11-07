'use strict';
const assert = require('assert');
const path = require('path');
const vscode = require('vscode');

const workspace = process.env.VB6_WORKSPACE;
const uriOf = (name) => vscode.Uri.file(path.join(workspace, name));

async function waitFor(what, condition, timeout = 30000) {
    const started = Date.now();
    for (;;) {
        const value = await condition();
        if (value) {
            return value;
        }
        if (Date.now() - started > timeout) {
            throw new Error(`timed out waiting for ${what}`);
        }
        await new Promise((resolve) => setTimeout(resolve, 100));
    }
}

describe('VB6 project in VS Code', function () {
    let api;

    before(async function () {
        const extension = vscode.extensions.getExtension('vb6-lsp.vb6-lsp');
        assert.ok(extension, 'the extension is installed in the development host');
        api = await extension.activate();
        await api.ready();
    });

    it('opens the project named by the vb6.project setting', async function () {
        const state = await waitFor('the project to load', () => {
            const s = api.state();
            return s.status.state === 'loaded' && s.info ? s : null;
        });

        assert.strictEqual(state.status.name, 'App');
        assert.strictEqual(state.status.isGroup, false);
        assert.strictEqual(state.status.memberCount, 3);
        assert.strictEqual(state.status.missingCount, 1);
        assert.ok(state.statusBar.text.includes('App'), state.statusBar.text);
        assert.ok(JSON.stringify(state.tree).includes('modMain'), JSON.stringify(state.tree));
    });

    it('analyses files that are not open and reports their problems', async function () {
        const diagnostics = await waitFor('diagnostics for the unopened module', () => {
            const found = vscode.languages.getDiagnostics(uriOf('modMain.bas'));
            return found.length > 0 ? found : null;
        });

        assert.strictEqual(diagnostics.length, 1);
        assert.strictEqual(diagnostics[0].range.start.line, 3);
        assert.strictEqual(diagnostics[0].message, 'Sub or Function not defined');
    });

    it('reports a listed file that is missing on the project file line', async function () {
        const found = await waitFor('the project file diagnostics', () => {
            const d = vscode.languages.getDiagnostics(uriOf('App.vbp'));
            return d.length > 0 ? d : null;
        });

        assert.strictEqual(found.length, 1);
        assert.ok(found[0].message.endsWith('modGone.bas'), found[0].message);
        assert.strictEqual(found[0].range.start.line, 4);
        assert.strictEqual(found[0].severity, vscode.DiagnosticSeverity.Error);
    });

    it('resolves a name defined in a module that is not open', async function () {
        const document = await vscode.workspace.openTextDocument(uriOf('modMain.bas'));
        await vscode.window.showTextDocument(document);

        const locations = await vscode.commands.executeCommand(
            'vscode.executeDefinitionProvider',
            document.uri,
            new vscode.Position(2, 6)
        );

        assert.strictEqual(locations.length, 1);
        assert.strictEqual(path.basename(locations[0].uri.fsPath), 'modUtil.bas');
        assert.strictEqual(locations[0].range.start.line, 1);
    });

    it('opens another project from the explorer command and remembers it', async function () {
        await vscode.commands.executeCommand('vb6.project.openFile', uriOf('Other.vbp'));

        const state = await waitFor('the other project to load', () => {
            const s = api.state();
            return s.status.state === 'loaded' && s.status.name === 'Other' ? s : null;
        });

        assert.strictEqual(state.status.memberCount, 1);
        assert.strictEqual(state.status.missingCount, 0);
        const remembered = await waitFor('the setting to be stored', () => {
            const value = vscode.workspace.getConfiguration('vb6').get('project');
            return value === 'Other.vbp' ? value : null;
        });
        assert.strictEqual(remembered, 'Other.vbp');
    });

    it('shows a .vbp file as a project page', async function () {
        await vscode.commands.executeCommand('vscode.openWith', uriOf('App.vbp'), 'vb6.projectEditor');

        const tab = await waitFor('the project page to open', () => {
            const active = vscode.window.tabGroups.activeTabGroup.activeTab;
            return active && active.input && active.input.viewType === 'vb6.projectEditor' ? active : null;
        });

        assert.strictEqual(path.basename(tab.input.uri.fsPath), 'App.vbp');
    });

    it('closes the project', async function () {
        await vscode.commands.executeCommand('vb6.project.close');

        const state = await waitFor('the project to close', () => {
            const s = api.state();
            return s.status.state === 'none' ? s : null;
        });

        assert.deepStrictEqual(state.tree, []);
    });
});
