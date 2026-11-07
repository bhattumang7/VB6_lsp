'use strict';
// Runs the integration suite inside a real VS Code with the extension and the debug build of the
// language server (build it first: `cargo build -p vb6-lsp --bin vb6-lsp`).
const fs = require('fs');
const os = require('os');
const path = require('path');
const { runTests, downloadAndUnzipVSCode } = require('@vscode/test-electron');

/** A folder with two projects; `App` is the one the `vb6.project` setting names. */
function createWorkspace(scratch, server) {
    const workspace = path.join(scratch, 'workspace');
    fs.mkdirSync(path.join(workspace, '.vscode'), { recursive: true });
    const write = (name, text) => fs.writeFileSync(path.join(workspace, name), text.replace(/\n/g, '\r\n'));
    write('.vscode/settings.json', JSON.stringify({ 'vb6.lsp.serverPath': server, 'vb6.project': 'App.vbp' }));
    write(
        'App.vbp',
        'Type=Exe\nName="App"\nModule=modMain; modMain.bas\nModule=modUtil; modUtil.bas\nModule=modGone; modGone.bas\n'
    );
    write('modMain.bas', 'Option Explicit\nSub Run()\n    Greet\n    NoSuchProc\nEnd Sub\n');
    write('modUtil.bas', 'Option Explicit\nPublic Sub Greet()\nEnd Sub\n');
    write('Other.vbp', 'Type=Exe\nName="Other"\nModule=modOther; modOther.bas\n');
    write('modOther.bas', 'Option Explicit\nSub Other()\nEnd Sub\n');
    return workspace;
}

async function main() {
    const root = path.resolve(__dirname, '..', '..');
    const serverName = process.platform === 'win32' ? 'vb6-lsp.exe' : 'vb6-lsp';
    const server = path.resolve(root, '..', 'target', 'debug', serverName);
    if (!fs.existsSync(server)) {
        throw new Error(`Build the language server first: ${server} does not exist`);
    }
    const scratch = fs.mkdtempSync(path.join(os.tmpdir(), 'vb6-vscode-'));
    const workspace = createWorkspace(scratch, server);
    const vscodeExecutablePath = process.env.VSCODE_EXECUTABLE || (await downloadAndUnzipVSCode());
    const exitCode = await runTests({
        vscodeExecutablePath,
        extensionDevelopmentPath: root,
        extensionTestsPath: path.join(__dirname, 'suite', 'index.js'),
        launchArgs: [
            workspace,
            '--disable-extensions',
            '--user-data-dir', path.join(scratch, 'user'),
            '--extensions-dir', path.join(scratch, 'extensions'),
        ],
        extensionTestsEnv: { VB6_WORKSPACE: workspace },
    });
    process.exit(exitCode);
}

main().catch((e) => {
    console.error(e);
    process.exit(1);
});
