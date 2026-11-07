'use strict';
const fs = require('fs');
const path = require('path');
const vscode = require('vscode');
const { LanguageClient, TransportKind } = require('vscode-languageclient/node');
const logic = require('./projectLogic');

const CHOSEN_PROJECT_KEY = 'vb6.pendingProject';

let client;
let clientReady;
let statusItem;
let currentStatus = { state: 'none' };
let currentInfo = null;
let treeProvider;
const pages = new Set();

// ── Server ──────────────────────────────────────────────────────────────────────

function resolveServerPath(context) {
    const configured = vscode.workspace.getConfiguration('vb6.lsp').get('serverPath');
    if (configured) {
        return configured;
    }
    const exe = process.platform === 'win32' ? 'vb6-lsp.exe' : 'vb6-lsp';
    const bundled = path.join(context.extensionPath, 'server', exe);
    return fs.existsSync(bundled) ? bundled : exe;
}

async function execute(command, ...args) {
    await clientReady;
    return client.sendRequest('workspace/executeCommand', { command, arguments: args });
}

// ── Project discovery ───────────────────────────────────────────────────────────

async function findProjectFiles() {
    const uris = await vscode.workspace.findFiles('**/*.{vbp,vbg}', '**/node_modules/**', 200);
    return uris.map((u) => u.fsPath);
}

function primaryFolder() {
    const folders = vscode.workspace.workspaceFolders;
    return folders && folders.length > 0 ? folders[0] : undefined;
}

async function initialProject(context) {
    const folder = primaryFolder();
    if (!folder) {
        return null;
    }
    const pending = context.globalState.get(CHOSEN_PROJECT_KEY);
    if (pending && logic.pathKey(pending.folder) === logic.pathKey(folder.uri.fsPath)) {
        await context.globalState.update(CHOSEN_PROJECT_KEY, undefined);
        await rememberProject(pending.project);
        return pending.project;
    }
    const config = vscode.workspace.getConfiguration('vb6', folder.uri);
    const choice = logic.chooseProject({
        folderPath: folder.uri.fsPath,
        configured: config.get('project'),
        candidates: await findProjectFiles(),
        autoLoad: config.get('autoLoadProject', true),
        exists: (p) => fs.existsSync(p),
    });
    if (choice.action === 'open') {
        return choice.path;
    }
    if (choice.reason === 'configuredMissing') {
        vscode.window.showWarningMessage(`The VB6 project set in "vb6.project" was not found: ${choice.path}`);
    } else if (choice.reason === 'several') {
        vscode.window
            .showInformationMessage('This folder has several VB6 projects. Choose the one to work on.', 'Select Project…')
            .then((pick) => pick && vscode.commands.executeCommand('vb6.project.select'));
    }
    return null;
}

/** Store the opened project in the `vb6.project` setting so the folder reopens with it. */
async function rememberProject(projectPath) {
    const folder = primaryFolder();
    if (!folder) {
        return;
    }
    try {
        const target = (vscode.workspace.workspaceFolders || []).length > 1
            ? vscode.ConfigurationTarget.WorkspaceFolder
            : vscode.ConfigurationTarget.Workspace;
        await vscode.workspace
            .getConfiguration('vb6', folder.uri)
            .update('project', logic.settingValue(folder.uri.fsPath, projectPath), target);
    } catch (e) {
        vscode.window.showWarningMessage(`Could not store the project in the workspace settings: ${e.message}`);
    }
}

// ── Opening and closing ─────────────────────────────────────────────────────────

/** Open `projectPath` (a .vbp or .vbg) as the project the language server analyses. */
async function openProject(context, projectPath) {
    if (!logic.isProjectFile(projectPath)) {
        vscode.window.showErrorMessage(`${projectPath} is not a VB6 project (.vbp) or group (.vbg).`);
        return;
    }
    const folders = vscode.workspace.workspaceFolders || [];
    const inside = folders.some((f) => logic.isInside(f.uri.fsPath, projectPath));
    if (folders.length === 0) {
        // Nothing is open: open the project's folder, and load the project once it is there.
        await context.globalState.update(CHOSEN_PROJECT_KEY, { folder: path.dirname(projectPath), project: projectPath });
        await vscode.commands.executeCommand('vscode.openFolder', vscode.Uri.file(path.dirname(projectPath)));
        return;
    }
    if (!inside) {
        const pick = await vscode.window.showInformationMessage(
            `${path.basename(projectPath)} is outside the open folder.`,
            'Open Containing Folder',
            'Use Anyway'
        );
        if (!pick) {
            return;
        }
        if (pick === 'Open Containing Folder') {
            await context.globalState.update(CHOSEN_PROJECT_KEY, { folder: path.dirname(projectPath), project: projectPath });
            await vscode.commands.executeCommand('vscode.openFolder', vscode.Uri.file(path.dirname(projectPath)));
            return;
        }
    }
    try {
        await execute('vb6.openProject', projectPath);
        await rememberProject(projectPath);
    } catch (e) {
        // The server has already told the user why.
        console.log('VB6 LSP: open project failed:', e && e.message);
    }
}

async function closeProject() {
    await execute('vb6.closeProject');
}

async function selectProject(context) {
    const found = await findProjectFiles();
    const items = logic.pickerItems(found, currentStatus.state === 'loaded' ? currentStatus.path : null);
    const extra = [{ label: '$(folder-opened) Browse…', browse: true }];
    if (currentStatus.state === 'loaded') {
        extra.push({ label: '$(close) Close Project', close: true });
    }
    const pick = await vscode.window.showQuickPick([...items, ...extra], {
        placeHolder: 'Select the VB6 project or group to open',
        matchOnDetail: true,
    });
    if (!pick) {
        return;
    }
    if (pick.close) {
        return closeProject();
    }
    if (pick.browse) {
        return browseForProject(context);
    }
    return openProject(context, pick.path);
}

async function browseForProject(context) {
    const picked = await vscode.window.showOpenDialog({
        canSelectMany: false,
        openLabel: 'Open Project',
        filters: { 'VB6 Project or Group': ['vbp', 'vbg'] },
        defaultUri: primaryFolder() && primaryFolder().uri,
    });
    if (picked && picked.length > 0) {
        return openProject(context, picked[0].fsPath);
    }
}

// ── Status, tree and pages ──────────────────────────────────────────────────────

function showStatus(status) {
    currentStatus = status;
    const view = logic.statusBarView(status);
    statusItem.text = view.text;
    statusItem.tooltip = view.tooltip;
    statusItem.command = view.command;
    statusItem.backgroundColor = view.warning ? new vscode.ThemeColor('statusBarItem.warningBackground') : undefined;
    statusItem.show();
    vscode.commands.executeCommand('setContext', 'vb6.projectLoaded', status.state === 'loaded');
}

async function onStatus(status) {
    showStatus(status);
    if (status.state === 'loaded') {
        try {
            currentInfo = await execute('vb6.projectInfo');
        } catch (e) {
            currentInfo = null;
        }
        const summary = logic.loadSummary(status);
        if (summary) {
            vscode.window.showWarningMessage(summary, 'Show Problems').then((pick) => {
                if (pick) {
                    vscode.commands.executeCommand('workbench.actions.view.problems');
                }
            });
        }
    } else if (status.state === 'none') {
        currentInfo = null;
    }
    treeProvider.refresh();
    for (const page of pages) {
        page.render();
    }
}

class ProjectTree {
    constructor() {
        this._changed = new vscode.EventEmitter();
        this.onDidChangeTreeData = this._changed.event;
    }

    refresh() {
        this._changed.fire();
    }

    getTreeItem(node) {
        const state = node.children
            ? node.collapsed
                ? vscode.TreeItemCollapsibleState.Collapsed
                : vscode.TreeItemCollapsibleState.Expanded
            : vscode.TreeItemCollapsibleState.None;
        const item = new vscode.TreeItem(node.label, state);
        item.id = node.id;
        item.description = node.description;
        item.tooltip = node.tooltip;
        item.iconPath = new vscode.ThemeIcon(
            node.icon,
            node.warning ? new vscode.ThemeColor('list.warningForeground') : undefined
        );
        item.contextValue = node.kind;
        if (node.open) {
            const options = node.open.line
                ? { selection: new vscode.Range(node.open.line - 1, 0, node.open.line - 1, 0) }
                : undefined;
            item.command = {
                command: 'vscode.open',
                title: 'Open',
                arguments: [vscode.Uri.parse(node.open.uri), options],
            };
        }
        return item;
    }

    getChildren(node) {
        return node ? node.children || [] : logic.buildTree(currentInfo);
    }
}

class ProjectPage {
    constructor(document, panel, context) {
        this.document = document;
        this.panel = panel;
        this.context = context;
        panel.webview.options = { enableScripts: true };
        panel.webview.onDidReceiveMessage((m) => this.onMessage(m));
        panel.onDidDispose(() => pages.delete(this));
        pages.add(this);
        this.render();
    }

    isLoaded() {
        return (
            currentStatus.state === 'loaded' &&
            logic.pathKey(currentStatus.path) === logic.pathKey(this.document.uri.fsPath)
        );
    }

    async render() {
        let info = null;
        let error = null;
        if (this.isLoaded() && currentInfo) {
            info = currentInfo;
        } else {
            try {
                info = await execute('vb6.projectPreview', this.document.uri.fsPath);
            } catch (e) {
                error = e && e.message;
            }
        }
        const nonce = String(Math.random()).slice(2) + String(Date.now());
        this.panel.webview.html = logic.renderProjectPage({
            info,
            error,
            loaded: this.isLoaded(),
            nonce,
            cspSource: this.panel.webview.cspSource,
        });
    }

    async onMessage(message) {
        const here = this.document.uri;
        switch (message.action) {
            case 'load':
                return openProject(this.context, here.fsPath);
            case 'reload':
                return openProject(this.context, here.fsPath);
            case 'close':
                return closeProject();
            case 'openText':
                return vscode.commands.executeCommand('vscode.openWith', here, 'default');
            case 'openFile':
                return vscode.commands.executeCommand('vscode.open', vscode.Uri.parse(message.uri));
            case 'openLine': {
                const target = message.uri ? vscode.Uri.parse(message.uri) : vscode.Uri.file(message.path);
                const line = Math.max((message.line || 1) - 1, 0);
                return vscode.commands.executeCommand('vscode.openWith', target, 'default', {
                    selection: new vscode.Range(line, 0, line, 0),
                });
            }
            default:
                return undefined;
        }
    }
}

class ProjectEditorProvider {
    constructor(context) {
        this.context = context;
    }

    openCustomDocument(uri) {
        return { uri, dispose() {} };
    }

    resolveCustomEditor(document, panel) {
        new ProjectPage(document, panel, this.context);
    }
}

// ── Activation ──────────────────────────────────────────────────────────────────

async function activate(context) {
    statusItem = vscode.window.createStatusBarItem(vscode.StatusBarAlignment.Left, 0);
    context.subscriptions.push(statusItem);
    treeProvider = new ProjectTree();
    context.subscriptions.push(
        vscode.window.createTreeView('vb6.projectExplorer', { treeDataProvider: treeProvider, showCollapseAll: true }),
        vscode.window.registerCustomEditorProvider('vb6.projectEditor', new ProjectEditorProvider(context), {
            supportsMultipleEditorsPerDocument: true,
        }),
        vscode.commands.registerCommand('vb6.project.select', () => selectProject(context)),
        vscode.commands.registerCommand('vb6.project.open', () => browseForProject(context)),
        vscode.commands.registerCommand('vb6.project.openFile', (uri) => uri && openProject(context, uri.fsPath)),
        vscode.commands.registerCommand('vb6.project.close', () => closeProject()),
        vscode.commands.registerCommand('vb6.project.reload', () => {
            if (currentStatus.path) {
                return openProject(context, currentStatus.path);
            }
        }),
        vscode.commands.registerCommand('vb6.project.showProblems', () =>
            vscode.commands.executeCommand('workbench.actions.view.problems')
        ),
        vscode.commands.registerCommand('vb6.project.openAsText', (uri) =>
            vscode.commands.executeCommand('vscode.openWith', uri, 'default')
        )
    );
    showStatus({ state: 'none' });

    const serverPath = resolveServerPath(context);
    const project = await initialProject(context);
    const serverOptions = {
        run: { command: serverPath, transport: TransportKind.stdio },
        debug: { command: serverPath, transport: TransportKind.stdio },
    };
    const clientOptions = {
        documentSelector: [{ scheme: 'file', language: 'vb6' }],
        initializationOptions: project ? { project } : {},
        synchronize: {
            fileEvents: vscode.workspace.createFileSystemWatcher('**/*.{bas,cls,frm,ctl,pag,dob,dsr,vbp,vbg}'),
        },
    };
    client = new LanguageClient('vb6Lsp', 'VB6 Language Server', serverOptions, clientOptions);
    client.onNotification('vb6/projectStatus', (status) => onStatus(status));
    clientReady = client.start().catch((e) => {
        vscode.window
            .showErrorMessage(`The VB6 language server could not start (${serverPath}): ${e.message}`, 'Open Settings')
            .then((pick) => pick && vscode.commands.executeCommand('workbench.action.openSettings', 'vb6.lsp.serverPath'));
        throw e;
    });
    context.subscriptions.push({ dispose: () => client && client.stop() });

    return {
        /** Resolves when the language server has started. */
        ready: () => clientReady,
        /** What the extension shows for the open project: for tests and for other extensions. */
        state: () => ({
            status: currentStatus,
            info: currentInfo,
            statusBar: { text: statusItem.text, tooltip: String(statusItem.tooltip || ''), command: statusItem.command },
            tree: logic.buildTree(currentInfo),
        }),
    };
}

function deactivate() {
    return client ? client.stop() : undefined;
}

module.exports = { activate, deactivate };
