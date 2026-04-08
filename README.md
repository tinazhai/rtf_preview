# RTF Live Preview

Preview RTF files directly in VS Code or Positron with full formatting, tables, and embedded images.

## Features

- **Zero dependencies required** — works out of the box with a built-in RTF parser
- **Enhanced rendering** — automatically uses LibreOffice or pandoc if installed for higher fidelity
- Tables with borders, column widths, cell backgrounds, and alignment
- Horizontal cell merging (colspan) with multi-level header alignment via edge-based column grid
- Merged-cell footnote deduplication (SAS-style repeated footnote text suppressed)
- Default light cell borders on all tables for consistent readability
- Cell padding support
- Text indentation (left indent, first-line/hanging indent)
- Per-border styling (different widths, colors, and styles per side)
- Embedded PNG/JPEG images displayed inline
- Font styling (bold, italic, underline, superscript)
- Page break indicators
- Auto-reloads on file change
- Optimized for SAS-generated RTF output

## Usage

1. Open any `.rtf` file
2. Click the preview icon in the editor title bar, or right-click → **RTF: Open Preview**
3. The preview opens in a side panel

## Rendering Modes

The extension automatically selects the best available renderer:

| Priority | Renderer | Quality | Requirement |
|----------|----------|---------|-------------|
| 1 | Built-in parser | Best for SAS | None (always available) |
| 2 | LibreOffice | Good | `soffice` in PATH |
| 3 | pandoc | Good | `pandoc` in PATH |

## Build from Source

```bash
npm install
npm run build
```

## License

MIT
