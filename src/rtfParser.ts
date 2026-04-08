/** Robust RTF-to-HTML converter for SAS RTF output.
 *  Handles tables with colspan, borders, cell backgrounds, alignment,
 *  embedded PNG/JPEG images, fonts, colors, formatting, page breaks.
 */

interface RtfState {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  fontSize: number;
  fontIndex: number;
  foreColor: number;
  backColor: number;
  superscript: boolean;
  subscript: boolean;
  align: string;
}

interface CellDef {
  borderTop: boolean;
  borderBottom: boolean;
  borderLeft: boolean;
  borderRight: boolean;
  borderWidth: number;
  borderColor: number;
  bgColor: number;
  vertAlign: string;
  rightEdge: number;
}

function defaultState(): RtfState {
  return { bold: false, italic: false, underline: false, strike: false, fontSize: 20, fontIndex: 0, foreColor: 0, backColor: 0, superscript: false, subscript: false, align: 'left' };
}
function cloneState(s: RtfState): RtfState { return { ...s }; }

export function rtfToHtml(rtf: string): string {
  const fonts = parseFontTable(rtf);
  const colors = parseColorTable(rtf);
  const tokens = tokenize(rtf);

  // Two-pass approach: first collect row info for colspan, then render
  const rowInfos = collectRowInfo(tokens);

  // Find max columns across all rows in each table
  const maxCols = computeMaxCols(rowInfos);

  const stateStack: RtfState[] = [];
  let state = defaultState();
  let html = '';
  let inTable = false, inRow = false, inCell = false;
  let cellDefs: CellDef[] = [];
  let cellIndex = 0;
  let skipDepth = 0;
  let groupDepth = 0;
  let inPict = false;
  let pictData = '';
  let pictProps = { width: 0, height: 0, type: 'png' };
  let pendingCellBorders: Partial<CellDef> = {};
  let rowCounter = 0;
  let tableStartRow = 0;

  const skipStarts = new Set(['header', 'footer', 'headerl', 'headerr', 'footerl', 'footerr', 'info', 'stylesheet', 'fonttbl', 'colortbl']);

  for (let i = 0; i < tokens.length; i++) {
    const tok = tokens[i];

    if (tok === '{') {
      groupDepth++;
      stateStack.push(cloneState(state));
      if (skipDepth > 0) { skipDepth++; continue; }
      if (i + 1 < tokens.length) {
        const next = tokens[i + 1];
        if (next === '\\*') {
          if (i + 2 < tokens.length) {
            const dest = tokens[i + 2];
            const m = dest.match(/^\\([a-z]+)/);
            if (m && (skipStarts.has(m[1]) || m[1] === 'bkmkstart' || m[1] === 'bkmkend' || m[1] === 'ud' || m[1] === 'fldinst')) {
              if (m[1] !== 'shppict') { skipDepth = 1; continue; }
            }
          }
          continue;
        }
        // Skip \field groups except shppict
        const cwm = next.match(/^\\([a-z]+)/);
        if (cwm && skipStarts.has(cwm[1])) { skipDepth = 1; continue; }
      }
      continue;
    }

    if (tok === '}') {
      groupDepth--;
      if (skipDepth > 0) { skipDepth--; state = stateStack.pop() || defaultState(); continue; }
      if (inPict) {
        inPict = false;
        if (pictData.length > 0) {
          // Ensure we're in a cell if in a table
          if (inTable && !inCell) {
            if (!inRow) { html += '<tr>'; inRow = true; }
            const colspan = getColspan(cellDefs, cellIndex, maxCols.get(tableStartRow) || 1);
            html += buildTdOpen(cellDefs, cellIndex, state, colors, colspan);
            inCell = true;
          }
          const b64 = hexToBase64(pictData.trim());
          const mime = pictProps.type === 'jpeg' ? 'image/jpeg' : 'image/png';
          const wIn = pictProps.width > 0 ? pictProps.width / 1440 : 8;
          const hIn = pictProps.height > 0 ? pictProps.height / 1440 : 5.5;
          html += `<img src="data:${mime};base64,${b64}" style="width:${wIn}in;height:${hIn}in;max-width:100%">`;
          pictData = '';
        }
      }
      state = stateStack.pop() || defaultState();
      continue;
    }

    if (skipDepth > 0) continue;

    // Inside picture — collect hex data
    if (inPict) {
      if (tok.startsWith('\\')) {
        const cwm = tok.match(/^\\([a-z]+)(-?\d+)?$/);
        if (cwm) {
          const w = cwm[1], p = cwm[2] !== undefined ? parseInt(cwm[2]) : undefined;
          if (w === 'pngblip') pictProps.type = 'png';
          else if (w === 'jpegblip') pictProps.type = 'jpeg';
          else if (w === 'picwgoal' && p !== undefined) pictProps.width = p;
          else if (w === 'pichgoal' && p !== undefined) pictProps.height = p;
        }
      } else {
        pictData += tok.replace(/\s/g, '');
      }
      continue;
    }

    // Control words
    if (tok.startsWith('\\')) {
      if (tok === '\\*') continue;
      const cwm = tok.match(/^\\([a-zA-Z]+)(-?\d+)?$/);
      if (!cwm) {
        if (tok === '\\\\') html += '\\';
        else if (tok === '\\{') html += '{';
        else if (tok === '\\}') html += '}';
        else if (tok === '\\~') html += '&nbsp;';
        else if (tok === '\\-') html += '&shy;';
        else if (tok === '\\_') html += '&ndash;';
        else if (tok.startsWith("\\'")) {
          const code = parseInt(tok.substring(2), 16);
          if (!isNaN(code)) html += String.fromCharCode(code);
        }
        continue;
      }

      const word = cwm[1];
      // Skip unknown uppercase control words (e.g. \E from SAS)
      if (word !== word.toLowerCase()) continue;

      const param = cwm[2] !== undefined ? parseInt(cwm[2]) : undefined;

      switch (word) {
        case 'b': state.bold = param !== 0; break;
        case 'i': state.italic = param !== 0; break;
        case 'ul': case 'ulw': state.underline = true; break;
        case 'ulnone': state.underline = false; break;
        case 'strike': state.strike = param !== 0; break;
        case 'fs': if (param !== undefined) state.fontSize = param; break;
        case 'f': if (param !== undefined) state.fontIndex = param; break;
        case 'cf': if (param !== undefined) state.foreColor = param; break;
        case 'cb': case 'highlight': if (param !== undefined) state.backColor = param; break;
        case 'super': state.superscript = true; state.subscript = false; break;
        case 'sub': state.subscript = true; state.superscript = false; break;
        case 'nosupersub': state.superscript = false; state.subscript = false; break;
        case 'ql': state.align = 'left'; break;
        case 'qc': state.align = 'center'; break;
        case 'qr': state.align = 'right'; break;
        case 'qj': state.align = 'justify'; break;
        case 'par': html += '<br>'; break;
        case 'line': html += '<br>'; break;
        case 'tab': html += '&emsp;'; break;
        case 'lquote': html += '\u2018'; break;
        case 'rquote': html += '\u2019'; break;
        case 'ldblquote': html += '\u201C'; break;
        case 'rdblquote': html += '\u201D'; break;
        case 'emdash': html += '\u2014'; break;
        case 'endash': html += '\u2013'; break;
        case 'bullet': html += '\u2022'; break;
        case 'u':
          if (param !== undefined) html += String.fromCharCode(param < 0 ? param + 65536 : param);
          break;
        case 'pard': state = { ...defaultState(), align: 'left' }; break;
        case 'plain': state = defaultState(); break;

        case 'trowd':
          if (!inTable) { html += '<table>'; inTable = true; tableStartRow = rowCounter; }
          if (inCell) { html += '</td>'; inCell = false; }
          if (inRow) { html += '</tr>'; }
          cellDefs = [];
          cellIndex = 0;
          pendingCellBorders = {};
          inRow = false;
          break;

        case 'clbrdrt': pendingCellBorders.borderTop = true; break;
        case 'clbrdrb': pendingCellBorders.borderBottom = true; break;
        case 'clbrdrl': pendingCellBorders.borderLeft = true; break;
        case 'clbrdrr': pendingCellBorders.borderRight = true; break;
        case 'brdrw': if (param !== undefined) pendingCellBorders.borderWidth = param; break;
        case 'brdrcf': if (param !== undefined) pendingCellBorders.borderColor = param; break;
        case 'clcbpat': if (param !== undefined) pendingCellBorders.bgColor = param; break;
        case 'clvertalc': pendingCellBorders.vertAlign = 'middle'; break;
        case 'clvertalb': pendingCellBorders.vertAlign = 'bottom'; break;
        case 'clvertalt': pendingCellBorders.vertAlign = 'top'; break;

        case 'cellx':
          if (param !== undefined) {
            cellDefs.push({
              borderTop: pendingCellBorders.borderTop || false,
              borderBottom: pendingCellBorders.borderBottom || false,
              borderLeft: pendingCellBorders.borderLeft || false,
              borderRight: pendingCellBorders.borderRight || false,
              borderWidth: pendingCellBorders.borderWidth || 0,
              borderColor: pendingCellBorders.borderColor || 0,
              bgColor: pendingCellBorders.bgColor || 0,
              vertAlign: pendingCellBorders.vertAlign || 'top',
              rightEdge: param,
            });
            pendingCellBorders = {};
          }
          break;

        case 'intbl': break;

        case 'cell':
          if (inCell) { html += '</td>'; inCell = false; }
          cellIndex++;
          break;

        case 'row':
          if (inCell) { html += '</td>'; inCell = false; }
          if (inRow) { html += '</tr>'; }
          inRow = false;
          cellIndex = 0;
          rowCounter++;
          break;

        case 'page':
          if (inCell) { html += '</td>'; inCell = false; }
          if (inRow) { html += '</tr>'; inRow = false; }
          if (inTable) { html += '</table>'; inTable = false; }
          html += '<hr class="page-break">';
          break;

        case 'pict':
          inPict = true;
          pictData = '';
          pictProps = { width: 0, height: 0, type: 'png' };
          break;

        default: break;
      }
      continue;
    }

    // Plain text
    if (tok.trim().length === 0 && !inTable) {
      if (tok.includes(' ')) html += ' ';
      continue;
    }

    // Start cell if needed
    if (inTable && !inCell) {
      if (!inRow) { html += '<tr>'; inRow = true; }
      const colspan = getColspan(cellDefs, cellIndex, maxCols.get(tableStartRow) || 1);
      html += buildTdOpen(cellDefs, cellIndex, state, colors, colspan);
      inCell = true;
    }

    // Emit styled text
    const styles: string[] = [];
    if (state.bold) styles.push('font-weight:bold');
    if (state.italic) styles.push('font-style:italic');
    const deco: string[] = [];
    if (state.underline) deco.push('underline');
    if (state.strike) deco.push('line-through');
    if (deco.length) styles.push(`text-decoration:${deco.join(' ')}`);
    if (state.fontSize !== 20) styles.push(`font-size:${state.fontSize / 2}pt`);
    if (state.foreColor > 0 && state.foreColor < colors.length) styles.push(`color:${colors[state.foreColor]}`);
    if (state.backColor > 0 && state.backColor < colors.length) styles.push(`background-color:${colors[state.backColor]}`);
    const fontName = fonts.get(state.fontIndex);
    if (fontName) styles.push(`font-family:'${fontName}',serif`);

    let text = escapeHtml(tok);
    if (state.superscript) text = `<sup>${text}</sup>`;
    if (state.subscript) text = `<sub>${text}</sub>`;

    if (styles.length > 0) {
      html += `<span style="${styles.join(';')}">${text}</span>`;
    } else {
      html += text;
    }
  }

  if (inCell) html += '</td>';
  if (inRow) html += '</tr>';
  if (inTable) html += '</table>';

  return html;
}

/** First pass: collect number of cells per row and group rows into tables */
interface RowInfo { numCells: number; tableIndex: number; }

function collectRowInfo(tokens: string[]): RowInfo[] {
  const rows: RowInfo[] = [];
  let cellCount = 0;
  let inTableDef = false;
  let tableIndex = 0;
  let wasInTable = false;
  let skipDepth = 0;
  let groupDepth = 0;
  const skipStarts = new Set(['header', 'footer', 'headerl', 'headerr', 'footerl', 'footerr', 'info', 'stylesheet', 'fonttbl', 'colortbl']);

  for (let i = 0; i < tokens.length; i++) {
    const tok = tokens[i];
    if (tok === '{') {
      groupDepth++;
      if (skipDepth > 0) { skipDepth++; continue; }
      if (i + 1 < tokens.length) {
        const next = tokens[i + 1];
        if (next === '\\*') {
          if (i + 2 < tokens.length) {
            const dest = tokens[i + 2];
            const m = dest.match(/^\\([a-z]+)/);
            if (m && (skipStarts.has(m[1]) || m[1] === 'bkmkstart' || m[1] === 'bkmkend' || m[1] === 'ud' || m[1] === 'fldinst') && m[1] !== 'shppict') {
              skipDepth = 1;
            }
          }
          continue;
        }
        const cwm = next.match(/^\\([a-z]+)/);
        if (cwm && skipStarts.has(cwm[1])) { skipDepth = 1; }
      }
      continue;
    }
    if (tok === '}') { groupDepth--; if (skipDepth > 0) skipDepth--; continue; }
    if (skipDepth > 0) continue;

    const cwm = tok.match(/^\\([a-z]+)(-?\d+)?$/);
    if (!cwm) continue;
    const word = cwm[1];
    const param = cwm[2] !== undefined ? parseInt(cwm[2]) : undefined;

    if (word === 'trowd') {
      if (!wasInTable) { tableIndex = rows.length; wasInTable = true; }
      cellCount = 0;
      inTableDef = true;
    } else if (word === 'cellx') {
      cellCount++;
    } else if (word === 'row') {
      rows.push({ numCells: cellCount, tableIndex });
      inTableDef = false;
    } else if (word === 'page') {
      wasInTable = false;
    }
  }
  return rows;
}

/** Compute max columns per table (keyed by first row index of that table) */
function computeMaxCols(rowInfos: RowInfo[]): Map<number, number> {
  const tableMaxCols = new Map<number, number>();
  for (const ri of rowInfos) {
    const cur = tableMaxCols.get(ri.tableIndex) || 0;
    if (ri.numCells > cur) tableMaxCols.set(ri.tableIndex, ri.numCells);
  }
  return tableMaxCols;
}

function getColspan(cellDefs: CellDef[], cellIndex: number, maxCols: number): number {
  if (cellDefs.length >= maxCols || cellDefs.length <= 1) return cellIndex === 0 && cellDefs.length < maxCols && cellDefs.length === 1 ? maxCols : 1;
  // Only apply colspan to the last cell if row has fewer cells than max
  if (cellIndex === cellDefs.length - 1 && cellDefs.length < maxCols) {
    return maxCols - cellDefs.length + 1;
  }
  return 1;
}

function buildTdOpen(cellDefs: CellDef[], cellIndex: number, state: RtfState, colors: string[], colspan: number): string {
  const def = cellDefs[cellIndex];
  const styles: string[] = [];
  if (def) {
    const prevEdge = cellIndex > 0 ? cellDefs[cellIndex - 1].rightEdge : 0;
    const widthTwips = def.rightEdge - prevEdge;
    if (colspan <= 1) styles.push(`width:${(widthTwips / 1440).toFixed(2)}in`);
    if (def.borderTop) styles.push(`border-top:${Math.max(1, def.borderWidth / 10)}px solid ${colors[def.borderColor] || '#000'}`);
    if (def.borderBottom) styles.push(`border-bottom:${Math.max(1, def.borderWidth / 10)}px solid ${colors[def.borderColor] || '#000'}`);
    if (def.borderLeft) styles.push(`border-left:${Math.max(1, def.borderWidth / 10)}px solid ${colors[def.borderColor] || '#000'}`);
    if (def.borderRight) styles.push(`border-right:${Math.max(1, def.borderWidth / 10)}px solid ${colors[def.borderColor] || '#000'}`);
    if (def.bgColor > 0 && def.bgColor < colors.length) styles.push(`background-color:${colors[def.bgColor]}`);
    styles.push(`vertical-align:${def.vertAlign}`);
  }
  styles.push(`text-align:${state.align}`);
  const colspanAttr = colspan > 1 ? ` colspan="${colspan}"` : '';
  return `<td${colspanAttr} style="${styles.join(';')}">`;
}

function parseFontTable(rtf: string): Map<number, string> {
  const fonts = new Map<number, string>();
  const m = rtf.match(/\{\\fonttbl([\s\S]*?)\}/);
  if (m) {
    const re = /\{\\f(\d+)[^}]*\s([^;{}]+);/g;
    let fm: RegExpExecArray | null;
    while ((fm = re.exec(m[1]))) fonts.set(parseInt(fm[1]), fm[2].trim());
  }
  return fonts;
}

function parseColorTable(rtf: string): string[] {
  const colors: string[] = ['#000000'];
  const m = rtf.match(/\{\\colortbl;?([\s\S]*?)\}/);
  if (m) {
    for (const entry of m[1].split(';')) {
      const r = entry.match(/\\red(\d+)/);
      const g = entry.match(/\\green(\d+)/);
      const b = entry.match(/\\blue(\d+)/);
      if (r && g && b) colors.push(`rgb(${r[1]},${g[1]},${b[1]})`);
    }
  }
  return colors;
}

function tokenize(rtf: string): string[] {
  const tokens: string[] = [];
  let i = 0;
  const len = rtf.length;
  while (i < len) {
    const ch = rtf[i];
    if (ch === '{') { tokens.push('{'); i++; continue; }
    if (ch === '}') { tokens.push('}'); i++; continue; }
    if (ch === '\\') {
      if (i + 1 < len) {
        const next = rtf[i + 1];
        if (next === '{' || next === '}' || next === '\\') { tokens.push('\\' + next); i += 2; continue; }
        if (next === '~' || next === '-' || next === '_') { tokens.push('\\' + next); i += 2; continue; }
        if (next === '*') { tokens.push('\\*'); i += 2; continue; }
        if (next === "'") { tokens.push("\\'" + rtf.substring(i + 2, i + 4)); i += 4; continue; }
        if (next === '\n' || next === '\r') { tokens.push('\\par'); i += 2; if (i < len && rtf[i] === '\n') i++; continue; }
        if (/[a-zA-Z]/.test(next)) {
          let j = i + 1;
          while (j < len && /[a-zA-Z]/.test(rtf[j])) j++;
          let k = j;
          if (k < len && (rtf[k] === '-' || /\d/.test(rtf[k]))) {
            if (rtf[k] === '-') k++;
            while (k < len && /\d/.test(rtf[k])) k++;
          }
          tokens.push(rtf.substring(i, k));
          if (k < len && rtf[k] === ' ') k++;
          i = k;
          continue;
        }
        tokens.push('\\' + next); i += 2; continue;
      }
      i++; continue;
    }
    if (ch === '\r' || ch === '\n') { i++; continue; }
    let j = i;
    while (j < len && rtf[j] !== '{' && rtf[j] !== '}' && rtf[j] !== '\\' && rtf[j] !== '\r' && rtf[j] !== '\n') j++;
    tokens.push(rtf.substring(i, j));
    i = j;
  }
  return tokens;
}

function hexToBase64(hex: string): string {
  const bytes: number[] = [];
  for (let i = 0; i < hex.length; i += 2) {
    const b = parseInt(hex.substring(i, i + 2), 16);
    if (!isNaN(b)) bytes.push(b);
  }
  if (typeof Buffer !== 'undefined') return Buffer.from(bytes).toString('base64');
  let binary = '';
  for (const b of bytes) binary += String.fromCharCode(b);
  return btoa(binary);
}

function escapeHtml(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
