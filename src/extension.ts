import * as vscode from 'vscode';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';
import { execFile } from 'child_process';

export function activate(context: vscode.ExtensionContext) {
  context.subscriptions.push(
    vscode.commands.registerCommand('rtfPreview.open', (uri?: vscode.Uri) => {
      const fileUri = uri || vscode.window.activeTextEditor?.document.uri;
      if (!fileUri) {
        vscode.window.showErrorMessage('No RTF file selected.');
        return;
      }
      openPreview(fileUri);
    })
  );
}

const panels = new Map<string, vscode.WebviewPanel>();

function openPreview(uri: vscode.Uri) {
  const key = uri.toString();
  const existing = panels.get(key);
  if (existing) { existing.reveal(); return; }

  const fileName = path.basename(uri.fsPath);
  const panel = vscode.window.createWebviewPanel(
    'rtfPreview',
    `Preview: ${fileName}`,
    vscode.ViewColumn.Beside,
    { enableScripts: false, localResourceRoots: [] }
  );

  panels.set(key, panel);
  panel.onDidDispose(() => panels.delete(key));

  renderRtf(uri, panel);

  // Watch for changes
  const watcher = vscode.workspace.createFileSystemWatcher(
    new vscode.RelativePattern(vscode.Uri.file(path.dirname(uri.fsPath)), path.basename(uri.fsPath))
  );
  watcher.onDidChange(() => renderRtf(uri, panel));
  panel.onDidDispose(() => watcher.dispose());
}

async function renderRtf(uri: vscode.Uri, panel: vscode.WebviewPanel) {
  panel.webview.html = loadingHtml();

  try {
    const html = await convertWithLibreOffice(uri.fsPath);
    panel.webview.html = html;
  } catch (err: any) {
    // Fallback: try pandoc
    try {
      const html = await convertWithPandoc(uri.fsPath);
      panel.webview.html = html;
    } catch (err2: any) {
      panel.webview.html = errorHtml(err2.message);
    }
  }
}

function convertWithLibreOffice(filePath: string): Promise<string> {
  return new Promise((resolve, reject) => {
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'rtfprev-'));
    const baseName = path.basename(filePath, path.extname(filePath));
    const outHtml = path.join(tmpDir, baseName + '.html');

    // Find libreoffice binary
    const soffice = findBinary(['soffice', 'libreoffice']);
    if (!soffice) {
      cleanup(tmpDir);
      return reject(new Error('LibreOffice not found. Install LibreOffice or ensure soffice is in PATH.'));
    }

    execFile(soffice, [
      '--headless', '--norestore', '--convert-to', 'html',
      '--outdir', tmpDir, filePath
    ], { timeout: 60000 }, (err, _stdout, stderr) => {
      if (err) { cleanup(tmpDir); return reject(new Error(`LibreOffice failed: ${stderr || err.message}`)); }
      if (!fs.existsSync(outHtml)) { cleanup(tmpDir); return reject(new Error('LibreOffice produced no output.')); }

      let html = fs.readFileSync(outHtml, 'utf-8');

      // Inline all referenced images as base64 data URIs
      html = html.replace(/(<img[^>]+src=")([^"]+)(")/g, (_match, pre, src, post) => {
        const imgPath = path.resolve(tmpDir, src);
        if (fs.existsSync(imgPath)) {
          const ext = path.extname(imgPath).slice(1) || 'png';
          const mime = ext === 'jpg' ? 'image/jpeg' : `image/${ext}`;
          const b64 = fs.readFileSync(imgPath).toString('base64');
          return `${pre}data:${mime};base64,${b64}${post}`;
        }
        return _match;
      });

      // Inject better styling
      html = injectStyles(html);

      cleanup(tmpDir);
      resolve(html);
    });
  });
}

function convertWithPandoc(filePath: string): Promise<string> {
  return new Promise((resolve, reject) => {
    const pandoc = findBinary(['pandoc']);
    if (!pandoc) { return reject(new Error('Neither LibreOffice nor pandoc found.')); }

    execFile(pandoc, [filePath, '-f', 'rtf', '-t', 'html', '--self-contained'], { timeout: 60000, maxBuffer: 50 * 1024 * 1024 }, (err, stdout, stderr) => {
      if (err) { return reject(new Error(`pandoc failed: ${stderr || err.message}`)); }
      resolve(injectStyles(stdout));
    });
  });
}

function findBinary(names: string[]): string | null {
  for (const name of names) {
    try {
      const result = require('child_process').execFileSync('which', [name], { encoding: 'utf-8' }).trim();
      if (result) { return result; }
    } catch { /* not found */ }
  }
  return null;
}

function injectStyles(html: string): string {
  const css = `<style>
    body { margin: 20px; background: #fff; color: #000; }
    table { border-collapse: collapse; }
    td, th { padding: 2px 6px; vertical-align: top; }
    img { max-width: 100%; height: auto; }
  </style>`;
  // Insert before </head> or prepend
  if (html.includes('</head>')) {
    return html.replace('</head>', css + '</head>');
  }
  return css + html;
}

function loadingHtml(): string {
  return `<!DOCTYPE html><html><body style="display:flex;justify-content:center;align-items:center;height:80vh;font-family:sans-serif;color:#666"><p>Converting RTF…</p></body></html>`;
}

function errorHtml(msg: string): string {
  return `<!DOCTYPE html><html><body style="margin:20px;font-family:sans-serif"><p style="color:red">Error: ${msg}</p><p>This extension requires LibreOffice or pandoc to be installed.</p></body></html>`;
}

function cleanup(dir: string) {
  try { fs.rmSync(dir, { recursive: true, force: true }); } catch { /* ignore */ }
}

export function deactivate() {}
