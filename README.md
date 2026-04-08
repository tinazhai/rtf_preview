# RTF Preview

Lightweight VS Code / Positron extension to preview RTF files with full formatting and embedded images.

## Requirements

One of the following must be installed:
- **LibreOffice** (preferred — best fidelity for SAS RTF output)
- **pandoc** (fallback)

## Usage

1. Open any `.rtf` file
2. Click the preview icon in the editor title bar, or right-click → **RTF: Open Preview**
3. The preview opens in a side panel with full formatting

## Features

- High-fidelity rendering via LibreOffice conversion
- Embedded PNG/JPEG images displayed inline
- Table formatting preserved (borders, column widths, cell backgrounds)
- Auto-reloads on file change
- Falls back to pandoc if LibreOffice is unavailable

## Build

```bash
cd rtf-preview
npm install
npm run build
```

The `out/` directory contains pre-compiled JS — no build step needed to run.
