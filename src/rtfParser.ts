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
  leftIndent: number;   // \li in twips
  firstIndent: number;  // \fi in twips
}

interface CellDef {
  borderTop: boolean;
  borderBottom: boolean;
  borderLeft: boolean;
  borderRight: boolean;
  borderTopStyle: string;
  borderBottomStyle: string;
  borderLeftStyle: string;
  borderRightStyle: string;
  borderTopWidth: number;
  borderBottomWidth: number;
  borderLeftWidth: number;
  borderRightWidth: number;
  borderTopColor: number;
  borderBottomColor: number;
  borderLeftColor: number;
  borderRightColor: number;
  bgColor: number;
  vertAlign: string;
  rightEdge: number;
  merged: boolean;       // \clmrg — continuation of horizontal merge
  mergeFirst: boolean;   // \clmgf — first cell in horizontal merge
  padTop: number;
  padRight: number;
  padBottom: number;
  padLeft: number;
}

function defaultState(): RtfState {
  return { bold: false, italic: false, underline: false, strike: false, fontSize: 20, fontIndex: 0, foreColor: 0, backColor: 0, superscript: false, subscript: false, align: 'left', leftIndent: 0, firstIndent: 0 };
}
function cloneState(s: RtfState): RtfState { return { ...s }; }

export function rtfToHtml(rtf: string): string {
  const fonts = parseFontTable(rtf);
  const colors = parseColorTable(rtf);
  const tokens = tokenize(rtf);

  // Two-pass approach: first collect row info for colspan, then render
  const rowInfos = collectRowInfo(tokens);

  // Build column grids per table for edge-based colspan
  const { grids: tableGrids, maxCols } = computeTableGrids(rowInfos);

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
  // Track which border side we're currently defining
  let pendingBorderSide: 'top' | 'bottom' | 'left' | 'right' | null = null;
  let rowCounter = 0;
  let tableStartRow = 0;
  let inMergedCell = false;
  let intbl = false;

  const skipStarts = new Set(['header', 'footer', 'headerl', 'headerr', 'footerl', 'footerr', 'headerf',
    'info', 'stylesheet', 'fonttbl', 'colortbl']);

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
            // Skip all \* destinations except shppict (contains images)
            if (m && m[1] !== 'shppict') { skipDepth = 1; continue; }
          }
          continue;
        }
        // Skip known non-\* destination groups
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
            while (cellIndex < cellDefs.length && cellDefs[cellIndex].merged) { cellIndex++; }
            const grid = tableGrids.get(tableStartRow);
            const colspan = computeCellColspan(cellDefs, cellIndex, grid);
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
        case 'par':
          if (inTable && !intbl) {
            if (inCell) { html += '</td>'; inCell = false; }
            if (inRow) { html += '</tr>'; inRow = false; }
            html += '</table>'; inTable = false;
          }
          html += '<br>'; break;
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
        case 'pard': state = { ...defaultState(), align: 'left' }; intbl = false; break;
        case 'plain': state = defaultState(); break;
        case 'li': if (param !== undefined) state.leftIndent = param; break;
        case 'fi': if (param !== undefined) state.firstIndent = param; break;
        case 'ri': case 'rin0': case 'lin0': break; // right/left indent variants, skip
        case 'keepn': case 'widctlpar': case 'wrapdefault': case 'faauto': case 'adjustright': case 'itap0': break;
        case 'trhdr': case 'trqc': case 'trql': case 'trqr': break; // row-level formatting
        case 'pnhang': break; // hanging indent for numbered paragraphs
        case 'rtlch': case 'ltrch': case 'loch': case 'hich': case 'dbch': case 'cgrid': break; // charset directives
        case 'lang': case 'langfe': case 'langnp': case 'langfenp': case 'alang': break; // language
        case 'af': case 'afs': break; // associated font
        case 'snext0': case 'sbasedon0': case 'sqformat': case 'spriority0': break; // stylesheet props

        case 'trowd':
          if (!inTable) { html += '<table>'; inTable = true; tableStartRow = rowCounter; }
          if (inCell) { html += '</td>'; inCell = false; }
          if (inRow) { html += '</tr>'; }
          cellDefs = [];
          cellIndex = 0;
          pendingCellBorders = {};
          pendingBorderSide = null;
          inRow = false;
          inMergedCell = false;
          break;

        case 'clbrdrt': pendingBorderSide = 'top'; pendingCellBorders.borderTop = true; break;
        case 'clbrdrb': pendingBorderSide = 'bottom'; pendingCellBorders.borderBottom = true; break;
        case 'clbrdrl': pendingBorderSide = 'left'; pendingCellBorders.borderLeft = true; break;
        case 'clbrdrr': pendingBorderSide = 'right'; pendingCellBorders.borderRight = true; break;
        case 'brdrs': case 'brdrdot': case 'brdrdash': case 'brdrdb': case 'brdrth': case 'brdrsh':
        case 'brdrnone': {
          const style = word === 'brdrnone' ? 'none' : word === 'brdrdot' ? 'dotted' : word === 'brdrdash' ? 'dashed' : word === 'brdrdb' ? 'double' : 'solid';
          if (pendingBorderSide === 'top') pendingCellBorders.borderTopStyle = style;
          else if (pendingBorderSide === 'bottom') pendingCellBorders.borderBottomStyle = style;
          else if (pendingBorderSide === 'left') pendingCellBorders.borderLeftStyle = style;
          else if (pendingBorderSide === 'right') pendingCellBorders.borderRightStyle = style;
          if (word === 'brdrnone') {
            if (pendingBorderSide === 'top') pendingCellBorders.borderTop = false;
            else if (pendingBorderSide === 'bottom') pendingCellBorders.borderBottom = false;
            else if (pendingBorderSide === 'left') pendingCellBorders.borderLeft = false;
            else if (pendingBorderSide === 'right') pendingCellBorders.borderRight = false;
          }
          break;
        }
        case 'brdrw':
          if (param !== undefined) {
            if (pendingBorderSide === 'top') pendingCellBorders.borderTopWidth = param;
            else if (pendingBorderSide === 'bottom') pendingCellBorders.borderBottomWidth = param;
            else if (pendingBorderSide === 'left') pendingCellBorders.borderLeftWidth = param;
            else if (pendingBorderSide === 'right') pendingCellBorders.borderRightWidth = param;
          }
          break;
        case 'brdrcf':
          if (param !== undefined) {
            if (pendingBorderSide === 'top') pendingCellBorders.borderTopColor = param;
            else if (pendingBorderSide === 'bottom') pendingCellBorders.borderBottomColor = param;
            else if (pendingBorderSide === 'left') pendingCellBorders.borderLeftColor = param;
            else if (pendingBorderSide === 'right') pendingCellBorders.borderRightColor = param;
          }
          break;
        case 'clcbpat': if (param !== undefined) pendingCellBorders.bgColor = param; break;
        case 'clvertalc': pendingCellBorders.vertAlign = 'middle'; break;
        case 'clvertalb': pendingCellBorders.vertAlign = 'bottom'; break;
        case 'clvertalt': pendingCellBorders.vertAlign = 'top'; break;
        case 'clmgf': pendingCellBorders.mergeFirst = true; pendingCellBorders.merged = false; break;
        case 'clmrg': pendingCellBorders.merged = true; pendingCellBorders.mergeFirst = false; break;
        case 'clpadt': if (param !== undefined) pendingCellBorders.padTop = param; break;
        case 'clpadr': if (param !== undefined) pendingCellBorders.padRight = param; break;
        case 'clpadb': if (param !== undefined) pendingCellBorders.padBottom = param; break;
        case 'clpadl': if (param !== undefined) pendingCellBorders.padLeft = param; break;
        case 'clpadft': case 'clpadfr': case 'clpadfb': case 'clpadfl': break; // padding format flags, ignore
        case 'cltxlrtb': case 'cltxbtlr': break; // cell text flow direction, ignore

        case 'cellx':
          if (param !== undefined) {
            cellDefs.push({
              borderTop: pendingCellBorders.borderTop || false,
              borderBottom: pendingCellBorders.borderBottom || false,
              borderLeft: pendingCellBorders.borderLeft || false,
              borderRight: pendingCellBorders.borderRight || false,
              borderTopStyle: pendingCellBorders.borderTopStyle || 'solid',
              borderBottomStyle: pendingCellBorders.borderBottomStyle || 'solid',
              borderLeftStyle: pendingCellBorders.borderLeftStyle || 'solid',
              borderRightStyle: pendingCellBorders.borderRightStyle || 'solid',
              borderTopWidth: pendingCellBorders.borderTopWidth || 0,
              borderBottomWidth: pendingCellBorders.borderBottomWidth || 0,
              borderLeftWidth: pendingCellBorders.borderLeftWidth || 0,
              borderRightWidth: pendingCellBorders.borderRightWidth || 0,
              borderTopColor: pendingCellBorders.borderTopColor || 0,
              borderBottomColor: pendingCellBorders.borderBottomColor || 0,
              borderLeftColor: pendingCellBorders.borderLeftColor || 0,
              borderRightColor: pendingCellBorders.borderRightColor || 0,
              bgColor: pendingCellBorders.bgColor || 0,
              vertAlign: pendingCellBorders.vertAlign || 'top',
              rightEdge: param,
              merged: pendingCellBorders.merged || false,
              mergeFirst: pendingCellBorders.mergeFirst || false,
              padTop: pendingCellBorders.padTop || 0,
              padRight: pendingCellBorders.padRight || 0,
              padBottom: pendingCellBorders.padBottom || 0,
              padLeft: pendingCellBorders.padLeft || 0,
            });
            pendingCellBorders = {};
            pendingBorderSide = null;
          }
          break;

        case 'intbl': intbl = true; break;

        case 'cell':
          // Emit empty cell if cell was never opened
          if (inTable && !inCell && !inMergedCell) {
            if (!inRow) { html += '<tr>'; inRow = true; }
            while (cellIndex < cellDefs.length && cellDefs[cellIndex].merged) { cellIndex++; }
            const grid2 = tableGrids.get(tableStartRow);
            const cs2 = computeCellColspan(cellDefs, cellIndex, grid2);
            html += buildTdOpen(cellDefs, cellIndex, state, colors, cs2);
            inCell = true;
          }
          if (inCell) { html += '</td>'; inCell = false; }
          inMergedCell = false;
          cellIndex++;
          // Check if the next raw cell is a merged continuation — suppress its content
          if (cellIndex < cellDefs.length && cellDefs[cellIndex].merged) {
            inMergedCell = true;
          }
          break;

        case 'row':
          if (inCell) { html += '</td>'; inCell = false; }
          if (inRow) { html += '</tr>'; }
          inRow = false;
          cellIndex = 0;
          inMergedCell = false;
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

    // Close table if we're no longer in intbl mode
    if (inTable && !intbl) {
      if (inCell) { html += '</td>'; inCell = false; }
      if (inRow) { html += '</tr>'; inRow = false; }
      html += '</table>';
      inTable = false;
    }

    // Suppress text in merged continuation cells
    if (inMergedCell) continue;

    // Start cell if needed
    if (inTable && !inCell) {
      if (!inRow) { html += '<tr>'; inRow = true; }
      // Skip merged continuation cells
      while (cellIndex < cellDefs.length && cellDefs[cellIndex].merged) {
        cellIndex++;
      }
      const grid = tableGrids.get(tableStartRow);
      const colspan = computeCellColspan(cellDefs, cellIndex, grid);
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
    if (state.leftIndent > 0) styles.push(`margin-left:${(state.leftIndent / 1440).toFixed(3)}in`);
    if (state.firstIndent !== 0) styles.push(`text-indent:${(state.firstIndent / 1440).toFixed(3)}in`);

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

/** First pass: collect cell edge positions per row and group rows into tables */
interface RowInfo { edges: number[]; tableIndex: number; }

function collectRowInfo(tokens: string[]): RowInfo[] {
  const rows: RowInfo[] = [];
  let edges: number[] = [];
  let tableIndex = 0;
  let wasInTable = false;
  let skipDepth = 0;
  let groupDepth = 0;
  const skipStarts = new Set(['header', 'footer', 'headerl', 'headerr', 'footerl', 'footerr', 'headerf',
    'info', 'stylesheet', 'fonttbl', 'colortbl']);

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
            if (m && m[1] !== 'shppict') { skipDepth = 1; }
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
      edges = [];
    } else if (word === 'cellx' && param !== undefined) {
      edges.push(param);
    } else if (word === 'row') {
      rows.push({ edges: [...edges], tableIndex });
    } else if (word === 'page') {
      wasInTable = false;
    }
  }
  return rows;
}

/** Build sorted unique column grid edges per table, and compute max grid columns */
function computeTableGrids(rowInfos: RowInfo[]): { grids: Map<number, number[]>; maxCols: Map<number, number> } {
  const edgeSets = new Map<number, Set<number>>();
  for (const ri of rowInfos) {
    let s = edgeSets.get(ri.tableIndex);
    if (!s) { s = new Set(); edgeSets.set(ri.tableIndex, s); }
    for (const e of ri.edges) s.add(e);
  }
  const grids = new Map<number, number[]>();
  const maxCols = new Map<number, number>();
  for (const [ti, s] of edgeSets) {
    const sorted = [...s].sort((a, b) => a - b);
    grids.set(ti, sorted);
    maxCols.set(ti, sorted.length);
  }
  return { grids, maxCols };
}

function computeCellColspan(cellDefs: CellDef[], cellIndex: number, grid: number[] | undefined): number {
  if (cellIndex >= cellDefs.length || !grid || grid.length === 0) return 1;
  const def = cellDefs[cellIndex];
  const leftEdge = cellIndex > 0 ? cellDefs[cellIndex - 1].rightEdge : 0;
  // For merged-first cells, include continuation cells' edges
  let rightEdge = def.rightEdge;
  if (def.mergeFirst) {
    for (let j = cellIndex + 1; j < cellDefs.length; j++) {
      if (cellDefs[j].merged) rightEdge = cellDefs[j].rightEdge;
      else break;
    }
  }
  // Count grid columns this cell spans
  let startCol = grid.indexOf(leftEdge) + 1; // grid column after leftEdge; if leftEdge=0, start at 0
  if (leftEdge === 0) startCol = 0;
  else if (startCol === 0) startCol = 0; // leftEdge not in grid, start at 0
  let endCol = grid.indexOf(rightEdge);
  if (endCol < 0) endCol = grid.length - 1;
  const span = endCol - startCol + 1;
  return span > 1 ? span : 1;
}

function buildTdOpen(cellDefs: CellDef[], cellIndex: number, state: RtfState, colors: string[], colspan: number): string {
  const def = cellDefs[cellIndex];
  const styles: string[] = [];
  if (def) {
    const prevEdge = cellIndex > 0 ? cellDefs[cellIndex - 1].rightEdge : 0;
    const widthTwips = def.rightEdge - prevEdge;
    if (colspan <= 1) styles.push(`width:${(widthTwips / 1440).toFixed(2)}in`);
    if (def.borderTop) styles.push(`border-top:${Math.max(1, def.borderTopWidth / 10)}px ${def.borderTopStyle || 'solid'} ${colors[def.borderTopColor] || '#000'}`);
    if (def.borderBottom) styles.push(`border-bottom:${Math.max(1, def.borderBottomWidth / 10)}px ${def.borderBottomStyle || 'solid'} ${colors[def.borderBottomColor] || '#000'}`);
    if (def.borderLeft) styles.push(`border-left:${Math.max(1, def.borderLeftWidth / 10)}px ${def.borderLeftStyle || 'solid'} ${colors[def.borderLeftColor] || '#000'}`);
    if (def.borderRight) styles.push(`border-right:${Math.max(1, def.borderRightWidth / 10)}px ${def.borderRightStyle || 'solid'} ${colors[def.borderRightColor] || '#000'}`);
    if (def.bgColor > 0 && def.bgColor < colors.length) styles.push(`background-color:${colors[def.bgColor]}`);
    styles.push(`vertical-align:${def.vertAlign}`);
    const pad: string[] = [];
    if (def.padTop) pad.push(`${(def.padTop / 1440 * 72).toFixed(0)}px`); else pad.push('2px');
    if (def.padRight) pad.push(`${(def.padRight / 1440 * 72).toFixed(0)}px`); else pad.push('6px');
    if (def.padBottom) pad.push(`${(def.padBottom / 1440 * 72).toFixed(0)}px`); else pad.push('2px');
    if (def.padLeft) pad.push(`${(def.padLeft / 1440 * 72).toFixed(0)}px`); else pad.push('6px');
    if (def.padTop || def.padRight || def.padBottom || def.padLeft) styles.push(`padding:${pad.join(' ')}`);
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
