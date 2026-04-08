import * as vscode from 'vscode';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';
import { execFile } from 'child_process';
import { rtfToHtml } from './rtfParser';

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

  const watcher = vscode.workspace.createFileSystemWatcher(
    new vscode.RelativePattern(vscode.Uri.file(path.dirname(uri.fsPath)), path.basename(uri.fsPath))
  );
  watcher.onDidChange(() => renderRtf(uri, panel));
  panel.onDidDispose(() => watcher.dispose());
}

async function renderRtf(uri: vscode.Uri, panel: vscode.WebviewPanel) {
  panel.webview.html = loadingHtml();

  // Try LibreOffice first for best fidelity
  const soffice = findBinary(['soffice', 'libreoffice']);
  if (soffice) {
    try {
      const html = await convertWithLibreOffice(uri.fsPath, soffice);
      panel.webview.html = html;
      return;
    } catch { /* fall through */ }
  }

  // Try pandoc
  const pandoc = findBinary(['pandoc']);
  if (pandoc) {
    try {
      const html = await convertWithPandoc(uri.fsPath, pandoc);
      panel.webview.html = html;
      return;
    } catch { /* fall through */ }
  }

  // Fallback: built-in JS parser
  try {
    const rtfContent = fs.readFileSync(uri.fsPath, 'latin1');
    const bodyHtml = rtfToHtml(rtfContent);
    panel.webview.html = wrapHtml(bodyHtml);
  } catch (err: any) {
    panel.webview.html = errorHtml(err.message);
  }
}

function convertWithLibreOffice(filePath: string, soffice: string): Promise<string> {
  return new Promise((resolve, reject) => {
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'rtfprev-'));
    const baseName = path.basename(filePath, path.extname(filePath));
    const outHtml = path.join(tmpDir, baseName + '.html');

    execFile(soffice, ['--headless', '--norestore', '--convert-to', 'html', '--outdir', tmpDir, filePath], { timeout: 60000 }, (err, _stdout, stderr) => {
      if (err || !fs.existsSync(outHtml)) { cleanup(tmpDir); return reject(new Error(stderr || 'conversion failed')); }
      let html = fs.readFileSync(outHtml, 'utf-8');
      // Inline images as base64
      html = html.replace(/(<img[^>]+src=")([^"]+)(")/g, (_m, pre, src, post) => {
        const imgPath = path.resolve(tmpDir, src);
        if (fs.existsSync(imgPath)) {
          const ext = path.extname(imgPath).slice(1) || 'png';
          const mime = ext === 'jpg' ? 'image/jpeg' : `image/${ext}`;
          return `${pre}data:${mime};base64,${fs.readFileSync(imgPath).toString('base64')}${post}`;
        }
        return _m;
      });
      html = injectStyles(html);
      cleanup(tmpDir);
      resolve(html);
    });
  });
}

function convertWithPandoc(filePath: string, pandoc: string): Promise<string> {
  return new Promise((resolve, reject) => {
    execFile(pandoc, [filePath, '-f', 'rtf', '-t', 'html', '--self-contained'], { timeout: 60000, maxBuffer: 50 * 1024 * 1024 }, (err, stdout, stderr) => {
      if (err) return reject(new Error(stderr || err.message));
      resolve(injectStyles(stdout));
    });
  });
}

function findBinary(names: string[]): string | null {
  for (const name of names) {
    try {
      const result = require('child_process').execFileSync('which', [name], { encoding: 'utf-8' }).trim();
      if (result) return result;
    } catch { /* not found */ }
  }
  return null;
}

function wrapHtml(body: string): string {
  return `<!DOCTYPE html>
<html lang="en">
<head><meta charset="UTF-8">
<style>
  body { font-family: 'Times New Roman', serif; font-size: 10pt; margin: 20px; background: #fff; color: #000; }
  table { border-collapse: collapse; margin: 4px auto; }
  td { padding: 2px 6px; border: 1px solid #d0d0d0; }
  hr.page-break { border: none; border-top: 2px dashed #999; margin: 20px 0; }
  img { display: block; margin: 8px auto; }
  sup, sub { font-size: 0.7em; }
</style>
</head>
<body>${body}</body></html>`;
}

function injectStyles(html: string): string {
  const css = `<style>
  body { margin: 20px; background: #fff; color: #000; }
  table { border-collapse: collapse; }
  td, th { padding: 2px 6px; vertical-align: top; border: 1px solid #d0d0d0; }
  img { max-width: 100%; height: auto; }
</style>`;
  return html.includes('</head>') ? html.replace('</head>', css + '</head>') : css + html;
}

function loadingHtml(): string {
  return `<!DOCTYPE html><html><body style="display:flex;justify-content:center;align-items:center;height:80vh;font-family:sans-serif;color:#666"><p>Converting RTF\u2026</p></body></html>`;
}

function errorHtml(msg: string): string {
  return `<!DOCTYPE html><html><body style="margin:20px;font-family:sans-serif"><p style="color:red">Error: ${msg}</p></body></html>`;
}

function cleanup(dir: string) {
  try { fs.rmSync(dir, { recursive: true, force: true }); } catch { /* ignore */ }
}

export function deactivate() {}
