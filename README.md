# RTF Live Preview

Preview RTF files directly in VS Code or Positron with full formatting, tables, and embedded images.

## Features

- **Zero dependencies required** — works out of the box with a built-in RTF parser
- **Enhanced rendering** — automatically uses LibreOffice or pandoc if installed for higher fidelity
- Tables with borders, column widths, cell backgrounds, and alignment
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
| 1 | LibreOffice | Best | `soffice` in PATH |
| 2 | pandoc | Good | `pandoc` in PATH |
| 3 | Built-in parser | Good | None (always available) |

## Build from Source

```bash
npm install
npm run build
```

## License

MIT
