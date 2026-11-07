'use strict';
// Decisions and views of the VB6 project feature that do not need the VS Code API: which
// project to open, what the status bar and the project tree show, and the project page.

const path = require('path');

const isWindows = process.platform === 'win32';

/** A key under which two spellings of one path compare equal. */
function pathKey(p) {
    const normalized = path.normalize(p);
    return isWindows ? normalized.toLowerCase() : normalized;
}

function isProjectFile(p) {
    return /\.(vbp|vbg)$/i.test(p);
}

function isGroupFile(p) {
    return /\.vbg$/i.test(p);
}

/**
 * The value stored in the `vb6.project` setting for `projectPath`: relative to the workspace
 * folder (with forward slashes, so it can be committed) when inside it, else absolute.
 */
function settingValue(folderPath, projectPath) {
    return isInside(folderPath, projectPath)
        ? path.relative(folderPath, projectPath).split(path.sep).join('/')
        : projectPath;
}

/** Whether `target` is a file or folder below `folderPath`. */
function isInside(folderPath, target) {
    const relative = path.relative(folderPath, target);
    return (
        relative !== '' &&
        relative !== '..' &&
        !relative.startsWith('..' + path.sep) &&
        !path.isAbsolute(relative)
    );
}

/** The absolute path a `vb6.project` setting names, or null when it is empty. */
function resolveSetting(folderPath, value) {
    if (typeof value !== 'string' || value.trim() === '') {
        return null;
    }
    const v = value.trim();
    return path.isAbsolute(v) ? path.normalize(v) : path.normalize(path.join(folderPath, v));
}

/**
 * Which project to open when the workspace folder is opened.
 *
 * `configured` is the `vb6.project` setting, `candidates` the project and group files found in
 * the folder, `exists` tells whether a path is a file.
 */
function chooseProject({ folderPath, configured, candidates, autoLoad, exists }) {
    const target = resolveSetting(folderPath, configured);
    if (target !== null) {
        return exists(target)
            ? { action: 'open', path: target, source: 'setting' }
            : { action: 'none', reason: 'configuredMissing', path: target };
    }
    if (!autoLoad) {
        return { action: 'none', reason: 'disabled' };
    }
    const groups = candidates.filter(isGroupFile);
    const projects = candidates.filter((c) => !isGroupFile(c));
    if (groups.length === 1) {
        return { action: 'open', path: groups[0], source: 'auto' };
    }
    if (groups.length === 0 && projects.length === 1) {
        return { action: 'open', path: projects[0], source: 'auto' };
    }
    if (groups.length + projects.length === 0) {
        return { action: 'none', reason: 'noProjects' };
    }
    return { action: 'none', reason: 'several' };
}

/** The status bar item for a `vb6/projectStatus` notification. */
function statusBarView(status) {
    const command = 'vb6.project.select';
    switch (status.state) {
        case 'loading':
            return {
                text: `$(sync~spin) VB6: Loading ${path.basename(status.path || '')}`,
                tooltip: `Loading ${status.path}`,
                warning: false,
                command,
            };
        case 'loaded': {
            const problems = status.errorCount + status.warningCount;
            let text = `$(project) ${status.name}`;
            if (status.isGroup) {
                text += ` · ${status.projectCount} ${status.projectCount === 1 ? 'project' : 'projects'}`;
            }
            if (problems > 0) {
                text += ` $(warning) ${problems}`;
            }
            const lines = [`${status.isGroup ? 'Group' : 'Project'}: ${status.path}`];
            lines.push(`${status.memberCount} ${status.memberCount === 1 ? 'file' : 'files'}`);
            if (status.missingCount > 0) {
                lines.push(`${status.missingCount} missing`);
            }
            if (status.errorCount > 0) {
                lines.push(`${status.errorCount} ${status.errorCount === 1 ? 'error' : 'errors'}`);
            }
            if (status.warningCount > 0) {
                lines.push(`${status.warningCount} ${status.warningCount === 1 ? 'warning' : 'warnings'}`);
            }
            if (status.informationCount > 0) {
                lines.push(`${status.informationCount} ${status.informationCount === 1 ? 'note' : 'notes'}`);
            }
            if (status.message) {
                lines.push(status.message);
            }
            return { text, tooltip: lines.join('\n'), warning: problems > 0, command };
        }
        case 'error':
            return {
                text: `$(error) VB6: cannot open ${path.basename(status.path || '')}`,
                tooltip: status.message || 'The project could not be opened',
                warning: true,
                command,
            };
        default:
            return {
                text: '$(folder-opened) VB6: No project',
                tooltip: 'Open a VB6 project (.vbp) or group (.vbg)',
                warning: false,
                command,
            };
    }
}

function plural(n, one, many) {
    return `${n} ${n === 1 ? one : many}`;
}

/** The message shown after a project loads with problems; null when it loaded cleanly. */
function loadSummary(status) {
    if (status.state !== 'loaded') {
        return null;
    }
    const parts = [];
    if (status.missingCount > 0) {
        parts.push(`${status.missingCount} of ${status.memberCount} files missing`);
    }
    const others = status.errorCount - status.missingCount;
    if (others > 0) {
        parts.push(plural(others, 'other error', 'other errors'));
    }
    if (status.warningCount > 0) {
        parts.push(plural(status.warningCount, 'warning', 'warnings'));
    }
    if (status.informationCount > 0) {
        parts.push(plural(status.informationCount, 'note', 'notes'));
    }
    if (parts.length === 0) {
        return null;
    }
    return `Loaded "${status.name}": ${parts.join(', ')}.`;
}

// ── Project tree ────────────────────────────────────────────────────────────────

const FOLDERS = [
    { ext: 'frm', label: 'Forms', icon: 'window' },
    { ext: 'bas', label: 'Modules', icon: 'symbol-namespace' },
    { ext: 'cls', label: 'Class Modules', icon: 'symbol-class' },
    { ext: 'ctl', label: 'User Controls', icon: 'symbol-interface' },
    { ext: 'pag', label: 'Property Pages', icon: 'symbol-property' },
    { ext: 'dob', label: 'User Documents', icon: 'file' },
    { ext: 'dsr', label: 'Designers', icon: 'symbol-event' },
];

const MEMBER_ICONS = Object.fromEntries(FOLDERS.map((f) => [f.ext, f.icon]));

const REFERENCE_STATUS = {
    loaded: { text: 'loaded', icon: 'library', warning: false },
    notLoaded: { text: 'not loaded', icon: 'warning', warning: true },
    projectLoaded: { text: 'project', icon: 'project', warning: false },
    projectMissing: { text: 'project not found', icon: 'error', warning: true },
    projectCircular: { text: 'circular reference', icon: 'error', warning: true },
    projectNotExposed: { text: 'exposes no types', icon: 'warning', warning: true },
};

function memberLabel(member) {
    return member.name || path.basename(member.path);
}

function memberNode(unitId, member) {
    const label = memberLabel(member);
    return {
        id: `${unitId}/${member.uri}`,
        kind: 'member',
        label,
        description: member.missing ? 'missing' : path.basename(member.path),
        tooltip: member.missing ? `${member.path} (file not found)` : member.path,
        icon: member.missing ? 'warning' : MEMBER_ICONS[member.extension] || 'file',
        warning: member.missing,
        open: member.missing ? null : { uri: member.uri },
    };
}

function referenceNode(unit, reference) {
    const status = REFERENCE_STATUS[reference.status] || { text: reference.status, icon: 'question', warning: false };
    return {
        id: `${unit.uri}/ref/${reference.line}`,
        kind: 'reference',
        label: reference.label,
        description: status.text,
        tooltip: `${reference.label}\n${reference.detail}`,
        icon: status.icon,
        warning: status.warning,
        open: { uri: unit.uri, line: reference.line },
    };
}

function unitNode(unit) {
    const children = [];
    for (const folder of FOLDERS) {
        const members = unit.members
            .filter((m) => m.extension === folder.ext)
            .sort((a, b) => memberLabel(a).toLowerCase().localeCompare(memberLabel(b).toLowerCase()));
        if (members.length > 0) {
            children.push({
                id: `${unit.uri}/folder/${folder.ext}`,
                kind: 'folder',
                label: folder.label,
                description: String(members.length),
                icon: 'folder',
                warning: false,
                collapsed: false,
                children: members.map((m) => memberNode(unit.uri, m)),
            });
        }
    }
    const others = unit.members.filter((m) => !MEMBER_ICONS[m.extension]);
    if (others.length > 0) {
        children.push({
            id: `${unit.uri}/folder/other`,
            kind: 'folder',
            label: 'Other files',
            description: String(others.length),
            icon: 'folder',
            warning: false,
            collapsed: true,
            children: others.map((m) => memberNode(unit.uri, m)),
        });
    }
    if (unit.references.length > 0) {
        children.push({
            id: `${unit.uri}/folder/references`,
            kind: 'folder',
            label: 'References',
            description: String(unit.references.length),
            icon: 'references',
            warning: false,
            collapsed: true,
            children: unit.references.map((r) => referenceNode(unit, r)),
        });
    }
    return {
        id: unit.uri,
        kind: 'project',
        label: unit.name,
        description: `${unit.startup ? 'startup · ' : ''}${unit.kind}`,
        tooltip: unit.path,
        icon: 'project',
        warning: false,
        collapsed: false,
        open: { uri: unit.uri },
        children,
    };
}

/** The Project Explorer tree for a `vb6.projectInfo` result. */
function buildTree(info) {
    if (!info) {
        return [];
    }
    const listed = info.projects.filter((p) => p.listed).map(unitNode);
    const referenced = info.projects.filter((p) => !p.listed).map(unitNode);
    const children = [...listed];
    if (referenced.length > 0) {
        children.push({
            id: `${info.uri}/referenced`,
            kind: 'folder',
            label: 'Referenced projects',
            description: String(referenced.length),
            icon: 'folder',
            warning: false,
            collapsed: true,
            children: referenced,
        });
    }
    if (!info.isGroup) {
        return children;
    }
    return [
        {
            id: info.uri,
            kind: 'group',
            label: info.name,
            description: 'group',
            tooltip: info.path,
            icon: 'root-folder',
            warning: false,
            collapsed: false,
            open: { uri: info.uri },
            children,
        },
    ];
}

// ── Quick pick ──────────────────────────────────────────────────────────────────

/** Quick-pick entries for the project and group files found in the workspace. */
function pickerItems(files, currentPath) {
    const currentKey = currentPath ? pathKey(currentPath) : null;
    return files
        .slice()
        .sort((a, b) => {
            if (isGroupFile(a) !== isGroupFile(b)) {
                return isGroupFile(a) ? -1 : 1;
            }
            return a.toLowerCase().localeCompare(b.toLowerCase());
        })
        .map((file) => ({
            path: file,
            label: `${isGroupFile(file) ? '$(root-folder)' : '$(project)'} ${path.basename(file)}`,
            description: pathKey(file) === currentKey ? 'open' : isGroupFile(file) ? 'group' : 'project',
            detail: path.dirname(file),
        }));
}

// ── Project page ────────────────────────────────────────────────────────────────

function esc(value) {
    return String(value)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

const SEVERITY_LABEL = { error: 'Error', warning: 'Warning', information: 'Note' };

function memberRow(unit, member) {
    const target = member.missing ? '' : ` data-action="openFile" data-uri="${esc(member.uri)}"`;
    return (
        `<tr class="${member.missing ? 'missing' : ''}">` +
        `<td>${target ? `<a href="#"${target}>${esc(memberLabel(member))}</a>` : esc(memberLabel(member))}</td>` +
        `<td>.${esc(member.extension)}</td>` +
        `<td>${esc(member.path)}</td>` +
        `<td>${member.missing ? 'missing' : ''}</td>` +
        `</tr>`
    );
}

function referenceRow(unit, reference) {
    const status = REFERENCE_STATUS[reference.status] || { text: reference.status, warning: false };
    return (
        `<tr class="${status.warning ? 'missing' : ''}">` +
        `<td><a href="#" data-action="openLine" data-uri="${esc(unit.uri)}" data-line="${reference.line}">${esc(reference.label)}</a></td>` +
        `<td>${esc(reference.kind)}</td>` +
        `<td>${esc(reference.detail)}</td>` +
        `<td>${esc(status.text)}</td>` +
        `</tr>`
    );
}

function unitSection(unit) {
    const members = unit.members.length
        ? `<table><thead><tr><th>File</th><th>Type</th><th>Path</th><th></th></tr></thead><tbody>${unit.members
              .map((m) => memberRow(unit, m))
              .join('')}</tbody></table>`
        : '<p class="muted">No files.</p>';
    const references = unit.references.length
        ? `<h4>References</h4><table><thead><tr><th>Name</th><th>Kind</th><th>Location</th><th>Status</th></tr></thead><tbody>${unit.references
              .map((r) => referenceRow(unit, r))
              .join('')}</tbody></table>`
        : '';
    return (
        `<section><h3>${esc(unit.name)} <span class="badge">${esc(unit.kind)}</span>` +
        `${unit.startup ? ' <span class="badge">startup</span>' : ''}` +
        `${unit.listed ? '' : ' <span class="badge">referenced only</span>'}</h3>` +
        `<p class="muted">${esc(unit.path)}</p>${members}${references}</section>`
    );
}

function problemItem(problem) {
    const where = `${path.basename(problem.file)}${problem.line > 0 ? `:${problem.line}` : ''}`;
    return (
        `<li class="${esc(problem.severity)}"><span class="sev">${esc(SEVERITY_LABEL[problem.severity] || problem.severity)}</span> ` +
        `<a href="#" data-action="openLine" data-path="${esc(problem.file)}" data-line="${problem.line}">${esc(where)}</a> ` +
        `${esc(problem.message)}</li>`
    );
}

/**
 * The page shown for a `.vbp`/`.vbg` file.
 *
 * `info` is the description of the file (null when it could not be read, with `error` saying
 * why); `loaded` tells whether this is the project the language server has open.
 */
function renderProjectPage({ info, error, loaded, nonce, cspSource }) {
    const csp =
        `default-src 'none'; style-src ${cspSource} 'unsafe-inline'; script-src 'nonce-${nonce}';`;
    let body;
    if (!info) {
        body =
            `<h1>VB6 project</h1><p class="error">${esc(error || 'The file could not be read.')}</p>` +
            `<div class="actions"><button data-action="openText">Edit as Text</button></div>`;
    } else {
        const what = info.isGroup ? 'Project group' : 'Project';
        const actions = loaded
            ? `<button data-action="reload">Reload</button><button data-action="close">Close Project</button>`
            : `<button class="primary" data-action="load">Load ${info.isGroup ? 'Group' : 'Project'}</button>`;
        const problems = info.problems.length
            ? `<h2>Problems (${info.problems.length})</h2><ul class="problems">${info.problems.map(problemItem).join('')}</ul>`
            : '';
        body =
            `<h1>${esc(info.name)} <span class="badge">${what}</span> ` +
            `${loaded ? '<span class="badge ok">loaded</span>' : '<span class="badge">not loaded</span>'}</h1>` +
            `<p class="muted">${esc(info.path)}</p>` +
            `<div class="actions">${actions}<button data-action="openText">Edit as Text</button></div>` +
            problems +
            `<h2>${info.isGroup ? 'Projects' : 'Contents'}</h2>` +
            info.projects.map(unitSection).join('');
    }
    return (
        `<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">` +
        `<meta http-equiv="Content-Security-Policy" content="${csp}">` +
        `<meta name="viewport" content="width=device-width, initial-scale=1.0">` +
        `<style>${PAGE_STYLE}</style></head><body>${body}` +
        `<script nonce="${nonce}">${PAGE_SCRIPT}</script></body></html>`
    );
}

const PAGE_STYLE = `
body { font-family: var(--vscode-font-family); color: var(--vscode-foreground); padding: 0 20px 40px; }
h1 { font-weight: 600; } h2 { margin-top: 28px; border-bottom: 1px solid var(--vscode-panel-border); padding-bottom: 4px; }
h3 { margin-bottom: 4px; }
.muted { color: var(--vscode-descriptionForeground); }
.badge { font-size: 0.6em; font-weight: normal; padding: 2px 8px; border-radius: 10px; background: var(--vscode-badge-background); color: var(--vscode-badge-foreground); vertical-align: middle; }
.badge.ok { background: var(--vscode-testing-iconPassed); color: var(--vscode-editor-background); }
.actions { margin: 12px 0; display: flex; gap: 8px; }
button { padding: 4px 12px; border: 1px solid var(--vscode-button-border, transparent); background: var(--vscode-button-secondaryBackground); color: var(--vscode-button-secondaryForeground); cursor: pointer; }
button.primary { background: var(--vscode-button-background); color: var(--vscode-button-foreground); }
table { border-collapse: collapse; width: 100%; margin: 8px 0 16px; }
th, td { text-align: left; padding: 4px 10px; border-bottom: 1px solid var(--vscode-panel-border); }
tr.missing td { color: var(--vscode-errorForeground); }
a { color: var(--vscode-textLink-foreground); }
.problems { list-style: none; padding: 0; } .problems li { padding: 3px 0; }
.sev { font-weight: 600; margin-right: 6px; }
li.error .sev { color: var(--vscode-errorForeground); } li.warning .sev { color: var(--vscode-editorWarning-foreground); } li.information .sev { color: var(--vscode-editorInfo-foreground); }
.error { color: var(--vscode-errorForeground); }
`;

const PAGE_SCRIPT = `
const vscode = acquireVsCodeApi();
document.addEventListener('click', (e) => {
  const el = e.target.closest('[data-action]');
  if (!el) { return; }
  e.preventDefault();
  vscode.postMessage({
    action: el.dataset.action,
    uri: el.dataset.uri || null,
    path: el.dataset.path || null,
    line: el.dataset.line ? Number(el.dataset.line) : null,
  });
});
`;

module.exports = {
    pathKey,
    isProjectFile,
    isGroupFile,
    settingValue,
    isInside,
    resolveSetting,
    chooseProject,
    statusBarView,
    loadSummary,
    buildTree,
    pickerItems,
    renderProjectPage,
    esc,
};
