const path = require('path');
const { workspace, ExtensionContext } = require('vscode');
const {
    LanguageClient,
    TransportKind
} = require('vscode-languageclient/node');

let client;

function activate(context) {
    // Get the server path from settings or use default
    const config = workspace.getConfiguration('vb6.lsp');
    let serverPath = config.get('serverPath');

    if (!serverPath) {
        // Default path - adjust this to your build location
        serverPath = 'c:\\projects\\VB6_lsp\\target\\release\\vb6-lsp.exe';
    }

    console.log('VB6 LSP: Starting server from', serverPath);

    // Server options - run the executable
    const serverOptions = {
        run: {
            command: serverPath,
            transport: TransportKind.stdio
        },
        debug: {
            command: serverPath,
            transport: TransportKind.stdio
        }
    };

    // Client options
    const clientOptions = {
        // Register the server for VB6 documents
        documentSelector: [
            { scheme: 'file', language: 'vb6' }
        ],
        synchronize: {
            // Notify the server about file changes to VB6 files
            fileEvents: workspace.createFileSystemWatcher('**/*.{bas,cls,frm,vb,ctl}')
        }
    };

    // Create the language client and start it
    client = new LanguageClient(
        'vb6Lsp',
        'VB6 Language Server',
        serverOptions,
        clientOptions
    );

    // Start the client (also starts the server)
    client.start();
    console.log('VB6 LSP: Client started');
}

function deactivate() {
    if (!client) {
        return undefined;
    }
    return client.stop();
}

module.exports = {
    activate,
    deactivate
};
