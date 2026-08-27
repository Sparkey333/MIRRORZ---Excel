/**
 * The painter.
 *
 * Draw order is Excel's, and it is not arbitrary: fills, then gridlines, then
 * borders, then text, then the selection overlay. Painting borders before
 * gridlines lets a hairline gridline show through a thin border and makes the
 * grid look dirty; painting text before borders lets a medium border eat a
 * descender.
 *
 * Two behaviours here are the ones a spreadsheet user notices instantly if they
 * are missing, and both are more work than they look:
 *
 *   Text overflows into adjacent EMPTY cells and is clipped the moment the
 *   neighbour is occupied. A long label in A1 runs across B1 and C1 until
 *   something is typed into B1, at which point it is cut mid-word. Without this
 *   every label looks truncated; with a naive version, text draws over data.
 *
 *   A number too wide for its column becomes a row of '#'. Text does not - it
 *   overflows or clips. Getting this backwards is immediately wrong to anyone
 *   who has used a spreadsheet.
 *
 * Everything is measured through the shared text cache, and the whole frame runs
 * inside one devicePixelRatio transform so that a 1-device-pixel gridline is
 * exactly one device pixel on a HiDPI screen rather than a 2px blur.
 */

import {
  CellError,
  MAX_COLS,
  MAX_ROWS,
  type CellData,
  type CellStyle,
  type Color,
  type DateSystem,
  type HorizontalAlign,
  type Scalar,
  type Sheet,
  type StyleTable,
  type VerticalAlign,
  colToName,
  packKey,
} from '@mirrorz/core';
import { resolveColor } from '@mirrorz/formats';
import { builtinFormatCode, format, overflowText } from '@mirrorz/formats/numfmt';

import type { AxisIndex, GridLayout, PaneLayout, Rect, Span } from './viewport.js';
import type { GridRange, SelectionState } from './selection.js';
import { type GridTheme, formatColorToHex } from './theme.js';
import { TextMeasureCache, fontHeights, fontString } from './text.js';

/**
 * The slice of CanvasRenderingContext2D the painter uses.
 *
 * Declared structurally rather than taken as CanvasRenderingContext2D so the
 * tests can hand in a recorder without pulling in a canvas implementation. A
 * real CanvasRenderingContext2D satisfies it.
 */
export interface GridRenderingContext {
  font: string;
  fillStyle: string | CanvasGradient | CanvasPattern;
  strokeStyle: string | CanvasGradient | CanvasPattern;
  lineWidth: number;
  textAlign: CanvasTextAlign;
  textBaseline: CanvasTextBaseline;
  globalAlpha: number;
  measureText(text: string): { width: number };
  save(): void;
  restore(): void;
  beginPath(): void;
  closePath(): void;
  rect(x: number, y: number, w: number, h: number): void;
  clip(): void;
  moveTo(x: number, y: number): void;
  lineTo(x: number, y: number): void;
  stroke(): void;
  fill(): void;
  fillRect(x: number, y: number, w: number, h: number): void;
  strokeRect(x: number, y: number, w: number, h: number): void;
  fillText(text: string, x: number, y: number, maxWidth?: number): void;
  translate(x: number, y: number): void;
  rotate(angle: number): void;
  setTransform(a: number, b: number, c: number, d: number, e: number, f: number): void;
  setLineDash(segments: number[]): void;
}

/** Where a cell lands. Satisfied by Viewport, and trivially by a test double. */
export interface CellGeometry {
  xOf(col: number): number;
  yOf(row: number): number;
  readonly rows: AxisIndex;
  readonly cols: AxisIndex;
}

export interface RenderInput {
  sheet: Sheet;
  styles: StyleTable;
  layout: GridLayout;
  geometry: CellGeometry;
  theme: GridTheme;
  measure: TextMeasureCache;
  selection?: SelectionState | null;
  showGridLines?: boolean;
  dateSystem?: DateSystem;
  /** Resolved theme palette from the workbook, for theme-indexed colours. */
  themePalette?: readonly string[];
  zoom?: number;
}

export interface RenderStats {
  /** Cell slots the painter looked at. The bound the huge-sheet test asserts. */
  cellsVisited: number;
  /** Cell slots that produced at least one draw call. */
  cellsPainted: number;
  fills: number;
  gridLines: number;
  borderSegments: number;
  texts: number;
  headerCells: number;
  /** Cells whose text ran into a neighbour. */
  overflowed: number;
  /** Cells rendered as '#' because the number would not fit. */
  hashed: number;
  merges: number;
  measureCalls: number;
}

/** Left and right inset inside a cell, matching Excel's 2-pixel text margin. */
const CELL_PAD = 3;
/** Excel's documented indent step: one level is about three spaces. */
const INDENT_SPACES = 3;
const MAX_OVERFLOW_COLS = 64;
/** textRotation 255 is the file format's spelling of "stacked vertically". */
const STACKED_ROTATION = 255;

export function renderGrid(ctx: GridRenderingContext, input: RenderInput): RenderStats {
  const { layout, theme, measure } = input;
  const stats: RenderStats = {
    cellsVisited: 0,
    cellsPainted: 0,
    fills: 0,
    gridLines: 0,
    borderSegments: 0,
    texts: 0,
    headerCells: 0,
    overflowed: 0,
    hashed: 0,
    merges: 0,
    measureCalls: 0,
  };
  const before = measure.stats.misses;

  ctx.save();
  // One transform for the whole frame: everything below is in CSS pixels and
  // lands on exact device pixels, which is what keeps text crisp at dpr 2.
  ctx.setTransform(layout.dpr, 0, 0, layout.dpr, 0, 0);
  measure.invalidateFont();

  ctx.fillStyle = theme.background;
  ctx.fillRect(0, 0, layout.width, layout.height);
  stats.fills++;

  for (const pane of layout.panes) paintPane(ctx, input, pane, stats);

  if (layout.showHeaders) paintHeaders(ctx, input, stats);
  paintFrozenDividers(ctx, input);

  ctx.restore();
  measure.invalidateFont();
  stats.measureCalls = measure.stats.misses - before;
  return stats;
}

// --- panes ----------------------------------------------------------------

interface MergeInfo {
  readonly range: GridRange;
  readonly anchorKey: number;
}

function paintPane(
  ctx: GridRenderingContext,
  input: RenderInput,
  pane: PaneLayout,
  stats: RenderStats,
): void {
  const { sheet, theme, layout } = input;
  if (pane.rows.length === 0 || pane.cols.length === 0) return;

  const firstRow = (pane.rows[0] as Span).index;
  const lastRow = (pane.rows[pane.rows.length - 1] as Span).index;
  const firstCol = (pane.cols[0] as Span).index;
  const lastCol = (pane.cols[pane.cols.length - 1] as Span).index;

  // Cells hidden under a merge, keyed by packed address. Bounded by the visible
  // area because only the on-screen part of each merge is enumerated.
  const covered = new Map<number, MergeInfo>();
  const visibleMerges: MergeInfo[] = [];
  for (const m of sheet.merges) {
    const r = m.range;
    if (r.end.row < firstRow || r.start.row > lastRow) continue;
    if (r.end.col < firstCol || r.start.col > lastCol) continue;
    const info: MergeInfo = {
      range: { top: r.start.row, left: r.start.col, bottom: r.end.row, right: r.end.col },
      anchorKey: packKey(r.start.row, r.start.col),
    };
    visibleMerges.push(info);
    const rTop = Math.max(firstRow, r.start.row);
    const rBottom = Math.min(lastRow, r.end.row);
    const cLeft = Math.max(firstCol, r.start.col);
    const cRight = Math.min(lastCol, r.end.col);
    for (let row = rTop; row <= rBottom; row++) {
      for (let col = cLeft; col <= cRight; col++) covered.set(packKey(row, col), info);
    }
  }
  stats.merges += visibleMerges.length;

  ctx.save();
  ctx.beginPath();
  ctx.rect(pane.rect.x, pane.rect.y, pane.rect.width, pane.rect.height);
  ctx.clip();

  paintFills(ctx, input, pane, covered, stats);
  if (input.showGridLines !== false) paintGridLines(ctx, input, pane, stats);
  paintMergeBodies(ctx, input, pane, visibleMerges, stats);
  paintBorders(ctx, input, pane, covered, stats);
  paintText(ctx, input, pane, covered, stats);
  paintMergeText(ctx, input, pane, visibleMerges, stats);
  if (input.selection) paintSelection(ctx, input, pane, input.selection);

  ctx.restore();
  void theme;
  void layout;
}

function paintFills(
  ctx: GridRenderingContext,
  input: RenderInput,
  pane: PaneLayout,
  covered: Map<number, MergeInfo>,
  stats: RenderStats,
): void {
  const { sheet, styles, themePalette } = input;
  for (const rowSpan of pane.rows) {
    for (const colSpan of pane.cols) {
      stats.cellsVisited++;
      if (covered.has(packKey(rowSpan.index, colSpan.index))) continue;
      const style = styles.get(sheet.getStyle(rowSpan.index, colSpan.index));
      const fill = fillColorOf(style, themePalette);
      if (!fill) continue;
      ctx.fillStyle = fill;
      ctx.fillRect(colSpan.start, rowSpan.start, colSpan.size, rowSpan.size);
      stats.fills++;
      stats.cellsPainted++;
    }
  }
}

function paintGridLines(
  ctx: GridRenderingContext,
  input: RenderInput,
  pane: PaneLayout,
  stats: RenderStats,
): void {
  const { theme, layout } = input;
  const dpr = layout.dpr;
  ctx.beginPath();
  ctx.strokeStyle = theme.gridLine;
  ctx.lineWidth = 1 / dpr;
  ctx.setLineDash([]);

  const top = pane.rect.y;
  const bottom = Math.min(pane.rect.y + pane.rect.height, lastEdge(pane.rows));
  const left = pane.rect.x;
  const right = Math.min(pane.rect.x + pane.rect.width, lastEdge(pane.cols));

  for (const col of pane.cols) {
    const x = crisp(col.start, dpr);
    ctx.moveTo(x, top);
    ctx.lineTo(x, bottom);
    stats.gridLines++;
  }
  const lastCol = pane.cols[pane.cols.length - 1] as Span;
  const endX = crisp(lastCol.start + lastCol.size, dpr);
  ctx.moveTo(endX, top);
  ctx.lineTo(endX, bottom);
  stats.gridLines++;

  for (const row of pane.rows) {
    const y = crisp(row.start, dpr);
    ctx.moveTo(left, y);
    ctx.lineTo(right, y);
    stats.gridLines++;
  }
  const lastRow = pane.rows[pane.rows.length - 1] as Span;
  const endY = crisp(lastRow.start + lastRow.size, dpr);
  ctx.moveTo(left, endY);
  ctx.lineTo(right, endY);
  stats.gridLines++;

  ctx.stroke();
}

function lastEdge(spans: readonly Span[]): number {
  const last = spans[spans.length - 1];
  return last ? last.start + last.size : 0;
}

function paintMergeBodies(
  ctx: GridRenderingContext,
  input: RenderInput,
  pane: PaneLayout,
  merges: readonly MergeInfo[],
  stats: RenderStats,
): void {
  const { sheet, styles, themePalette, theme, layout } = input;
  for (const m of merges) {
    const rect = mergeRect(input, m.range);
    const style = styles.get(sheet.getStyle(m.range.top, m.range.left));
    const fill = fillColorOf(style, themePalette) ?? theme.background;
    ctx.fillStyle = fill;
    ctx.fillRect(rect.x, rect.y, rect.width, rect.height);
    stats.fills++;
    // A merge is one region, so the internal gridlines must not show through;
    // its outline is redrawn here after the fill has covered them.
    ctx.beginPath();
    ctx.strokeStyle = theme.gridLine;
    ctx.lineWidth = 1 / layout.dpr;
    ctx.setLineDash([]);
    ctx.rect(
      crisp(rect.x, layout.dpr),
      crisp(rect.y, layout.dpr),
      Math.max(0, rect.width),
      Math.max(0, rect.height),
    );
    ctx.stroke();
    stats.gridLines++;
  }
  void pane;
}

function mergeRect(input: RenderInput, range: GridRange): Rect {
  const { geometry } = input;
  const x = geometry.xOf(range.left);
  const y = geometry.yOf(range.top);
  const endX = geometry.xOf(range.right) + geometry.cols.sizeOf(range.right);
  const endY = geometry.yOf(range.bottom) + geometry.rows.sizeOf(range.bottom);
  return { x, y, width: endX - x, height: endY - y };
}

// --- borders --------------------------------------------------------------

function paintBorders(
  ctx: GridRenderingContext,
  input: RenderInput,
  pane: PaneLayout,
  covered: Map<number, MergeInfo>,
  stats: RenderStats,
): void {
  const { sheet, styles, themePalette, layout } = input;
  for (const rowSpan of pane.rows) {
    for (const colSpan of pane.cols) {
      if (covered.has(packKey(rowSpan.index, colSpan.index))) continue;
      const style = styles.get(sheet.getStyle(rowSpan.index, colSpan.index));
      const border = style.border;
      if (!border) continue;
      const x0 = colSpan.start;
      const y0 = rowSpan.start;
      const x1 = colSpan.start + colSpan.size;
      const y1 = rowSpan.start + rowSpan.size;
      if (border.top?.style && border.top.style !== 'none') {
        strokeEdge(ctx, x0, y0, x1, y0, border.top.style, border.top.color, themePalette, layout.dpr);
        stats.borderSegments++;
        stats.cellsPainted++;
      }
      if (border.bottom?.style && border.bottom.style !== 'none') {
        strokeEdge(ctx, x0, y1, x1, y1, border.bottom.style, border.bottom.color, themePalette, layout.dpr);
        stats.borderSegments++;
        stats.cellsPainted++;
      }
      if (border.left?.style && border.left.style !== 'none') {
        strokeEdge(ctx, x0, y0, x0, y1, border.left.style, border.left.color, themePalette, layout.dpr);
        stats.borderSegments++;
        stats.cellsPainted++;
      }
      if (border.right?.style && border.right.style !== 'none') {
        strokeEdge(ctx, x1, y0, x1, y1, border.right.style, border.right.color, themePalette, layout.dpr);
        stats.borderSegments++;
        stats.cellsPainted++;
      }
      if (border.diagonal?.style && border.diagonal.style !== 'none') {
        if (border.diagonalDown) {
          strokeEdge(ctx, x0, y0, x1, y1, border.diagonal.style, border.diagonal.color, themePalette, layout.dpr);
          stats.borderSegments++;
        }
        if (border.diagonalUp) {
          strokeEdge(ctx, x0, y1, x1, y0, border.diagonal.style, border.diagonal.color, themePalette, layout.dpr);
          stats.borderSegments++;
        }
      }
    }
  }
}

interface StrokeSpec {
  width: number;
  dash: number[];
  double: boolean;
}

/** Pixel weights and dash patterns for Excel's fourteen border styles. */
export function borderStroke(style: string): StrokeSpec {
  switch (style) {
    case 'hair':
      return { width: 1, dash: [1, 1], double: false };
    case 'dotted':
      return { width: 1, dash: [1, 2], double: false };
    case 'dashed':
      return { width: 1, dash: [3, 2], double: false };
    case 'dashDot':
      return { width: 1, dash: [4, 2, 1, 2], double: false };
    case 'dashDotDot':
      return { width: 1, dash: [4, 2, 1, 2, 1, 2], double: false };
    case 'medium':
      return { width: 2, dash: [], double: false };
    case 'mediumDashed':
      return { width: 2, dash: [4, 2], double: false };
    case 'mediumDashDot':
      return { width: 2, dash: [5, 2, 1, 2], double: false };
    case 'mediumDashDotDot':
      return { width: 2, dash: [5, 2, 1, 2, 1, 2], double: false };
    case 'slantDashDot':
      return { width: 2, dash: [4, 2, 1, 2], double: false };
    case 'thick':
      return { width: 3, dash: [], double: false };
    case 'double':
      return { width: 1, dash: [], double: true };
    default:
      return { width: 1, dash: [], double: false };
  }
}

function strokeEdge(
  ctx: GridRenderingContext,
  x0: number,
  y0: number,
  x1: number,
  y1: number,
  style: string,
  color: Color | undefined,
  palette: readonly string[] | undefined,
  dpr: number,
): void {
  const spec = borderStroke(style);
  ctx.strokeStyle = resolveColor(color, palette) ?? '#000000';
  ctx.lineWidth = spec.width / dpr > 0 ? spec.width : 1;
  ctx.setLineDash(spec.dash);
  if (spec.double) {
    // Excel's double border is two hairlines with a one-pixel gap.
    const dx = x0 === x1 ? 1 : 0;
    const dy = y0 === y1 ? 1 : 0;
    line(ctx, x0 - dx, y0 - dy, x1 - dx, y1 - dy, dpr);
    line(ctx, x0 + dx, y0 + dy, x1 + dx, y1 + dy, dpr);
  } else {
    line(ctx, x0, y0, x1, y1, dpr);
  }
  ctx.setLineDash([]);
}

function line(
  ctx: GridRenderingContext,
  x0: number,
  y0: number,
  x1: number,
  y1: number,
  dpr: number,
): void {
  ctx.beginPath();
  ctx.moveTo(x0 === x1 ? crisp(x0, dpr) : x0, y0 === y1 ? crisp(y0, dpr) : y0);
  ctx.lineTo(x0 === x1 ? crisp(x1, dpr) : x1, y0 === y1 ? crisp(y1, dpr) : y1);
  ctx.stroke();
}

// --- text -----------------------------------------------------------------

function paintText(
  ctx: GridRenderingContext,
  input: RenderInput,
  pane: PaneLayout,
  covered: Map<number, MergeInfo>,
  stats: RenderStats,
): void {
  const { sheet } = input;
  for (const rowSpan of pane.rows) {
    for (const colSpan of pane.cols) {
      const key = packKey(rowSpan.index, colSpan.index);
      if (covered.has(key)) continue;
      const cell = sheet.cells.get(key);
      if (!cell || cell.value === null) continue;
      const rect: Rect = {
        x: colSpan.start,
        y: rowSpan.start,
        width: colSpan.size,
        height: rowSpan.size,
      };
      if (drawCellText(ctx, input, rowSpan.index, colSpan.index, cell, rect, false, stats)) {
        stats.cellsPainted++;
      }
    }
  }
}

function paintMergeText(
  ctx: GridRenderingContext,
  input: RenderInput,
  pane: PaneLayout,
  merges: readonly MergeInfo[],
  stats: RenderStats,
): void {
  const { sheet } = input;
  for (const m of merges) {
    const cell = sheet.cells.get(m.anchorKey);
    if (!cell || cell.value === null) continue;
    const rect = mergeRect(input, m.range);
    // A merge is a closed box: text never spills out of it.
    if (drawCellText(ctx, input, m.range.top, m.range.left, cell, rect, true, stats)) {
      stats.cellsPainted++;
    }
  }
  void pane;
}

export function formatCodeOf(style: CellStyle): string {
  if (style.numFmt) return style.numFmt;
  if (style.numFmtId !== undefined) return builtinFormatCode(style.numFmtId) ?? 'General';
  return 'General';
}

/**
 * General alignment, which is a property of the VALUE and not of the format:
 * numbers and dates right, logicals and errors centred, everything else left.
 * This is why a number stored as text jumps to the left of its column, the
 * oldest diagnostic in spreadsheet history.
 */
export function generalAlign(value: Scalar): 'left' | 'center' | 'right' {
  if (typeof value === 'number') return 'right';
  if (typeof value === 'boolean') return 'center';
  if (value instanceof CellError) return 'center';
  return 'left';
}

export function effectiveHAlign(
  align: HorizontalAlign | undefined,
  value: Scalar,
): 'left' | 'center' | 'right' | 'fill' | 'justify' | 'distributed' {
  switch (align) {
    case undefined:
    case 'general':
      return generalAlign(value);
    case 'centerContinuous':
      return 'center';
    case 'left':
    case 'center':
    case 'right':
    case 'fill':
    case 'justify':
    case 'distributed':
      return align;
    default:
      return generalAlign(value);
  }
}

function verticalOffset(
  vertical: VerticalAlign | undefined,
  rect: Rect,
  blockHeight: number,
): number {
  switch (vertical) {
    case 'top':
      return rect.y + CELL_PAD;
    case 'center':
    case 'distributed':
    case 'justify':
      return rect.y + (rect.height - blockHeight) / 2;
    default:
      // Excel's default is bottom, which is why a tall row pins its text down.
      return rect.y + rect.height - CELL_PAD - blockHeight;
  }
}

/** Returns true when anything was actually drawn. */
function drawCellText(
  ctx: GridRenderingContext,
  input: RenderInput,
  row: number,
  col: number,
  cell: CellData,
  rect: Rect,
  isMerge: boolean,
  stats: RenderStats,
): boolean {
  const { sheet, styles, theme, measure, themePalette, zoom = 1 } = input;
  const style = styles.get(sheet.getStyle(row, col));
  const align = style.alignment ?? {};
  const value = cell.value;

  const result = format(value, formatCodeOf(style), { dateSystem: input.dateSystem ?? 1900 });
  if (result.text === '' && !result.overflow) return false;

  const font = fontString(
    {
      family: style.font?.name,
      size: style.font?.size,
      bold: style.font?.bold === true,
      italic: style.font?.italic === true,
    },
    { family: theme.defaultFontFamily, size: theme.defaultFontSize },
    zoom,
  );
  const fontSize = (style.font?.size ?? theme.defaultFontSize) * zoom;
  const metrics = fontHeights(fontSize);

  const color =
    formatColorToHex(result.color?.name, theme) ??
    resolveColor(style.font?.color, themePalette) ??
    (value instanceof CellError ? theme.errorText : theme.cellText);

  const indentUnit = indentStep(ctx, measure, font, fontSize);
  const indent = (align.indent ?? 0) * indentUnit;
  const avail = Math.max(0, rect.width - 2 * CELL_PAD - indent);

  const rotation = align.textRotation ?? 0;
  if (rotation !== 0) {
    drawRotated(ctx, input, result.text, font, color, rect, rotation, align.vertical, metrics);
    stats.texts++;
    return true;
  }

  if (align.wrapText) {
    drawWrapped(ctx, input, result.text, font, color, rect, indent, avail, align, value, metrics);
    stats.texts++;
    return true;
  }

  let text = result.text;
  let width = measure.measure(ctx, text, font);

  // Numbers that will not fit become '#'. Text does not: it overflows or clips.
  if ((result.numeric || result.overflow) && (result.overflow || width > avail)) {
    const hashWidth = measure.measure(ctx, '#', font) || fontSize * 0.5;
    text = overflowText(Math.max(1, Math.floor(avail / hashWidth)));
    width = measure.measure(ctx, text, font);
    stats.hashed++;
  }

  const h = effectiveHAlign(align.horizontal, value);
  let boxX = rect.x + CELL_PAD;
  let boxWidth = rect.width - 2 * CELL_PAD;

  if (!isMerge && width > avail) {
    // Overflow into empty neighbours, in the direction the text is aligned.
    const need = width - avail;
    let leftRoom = 0;
    let rightRoom = 0;
    if (h === 'left' || h === 'center') {
      rightRoom = overflowRoom(input, row, col, 1, h === 'center' ? need / 2 : need);
    }
    if (h === 'right' || h === 'center') {
      leftRoom = overflowRoom(input, row, col, -1, h === 'center' ? need / 2 : need);
    }
    if (leftRoom > 0 || rightRoom > 0) {
      boxX -= leftRoom;
      boxWidth += leftRoom + rightRoom;
      stats.overflowed++;
    }
  }

  const clipNeeded = width > boxWidth - indent;
  if (clipNeeded) {
    ctx.save();
    ctx.beginPath();
    ctx.rect(boxX, rect.y, boxWidth, rect.height);
    ctx.clip();
  }

  measure.useFont(ctx, font);
  ctx.fillStyle = color;
  ctx.textBaseline = 'alphabetic';
  ctx.textAlign = 'left';

  const baseline = verticalOffset(align.vertical, rect, metrics.ascent + metrics.descent) + metrics.ascent;
  let x: number;
  switch (h) {
    case 'right':
      x = boxX + boxWidth - indent - width;
      break;
    case 'center':
      x = boxX + (boxWidth - width) / 2;
      break;
    default:
      x = boxX + indent;
      break;
  }
  ctx.fillText(text, x, baseline);
  stats.texts++;

  if (clipNeeded) ctx.restore();
  return true;
}

/**
 * How far text may run past its cell in one direction.
 *
 * Stops at the first neighbour that holds anything, or that belongs to a merge,
 * which is exactly Excel's rule. The 64-column cap keeps a pathological case (a
 * 10,000 character label on an empty row) from walking the sheet.
 */
function overflowRoom(
  input: RenderInput,
  row: number,
  col: number,
  direction: 1 | -1,
  need: number,
): number {
  const { sheet, geometry } = input;
  let room = 0;
  let c = col + direction;
  for (let n = 0; n < MAX_OVERFLOW_COLS && room < need; n++) {
    if (c < 0 || c >= MAX_COLS) break;
    const neighbour = sheet.cells.get(packKey(row, c));
    if (neighbour && neighbour.value !== null) break;
    if (sheet.mergeAt(row, c)) break;
    room += geometry.cols.sizeOf(c);
    c += direction;
  }
  return Math.min(room, Math.max(0, need));
}

function drawWrapped(
  ctx: GridRenderingContext,
  input: RenderInput,
  text: string,
  font: string,
  color: string,
  rect: Rect,
  indent: number,
  avail: number,
  align: { horizontal?: HorizontalAlign; vertical?: VerticalAlign },
  value: Scalar,
  metrics: { ascent: number; descent: number; line: number },
): void {
  const { measure } = input;
  const lines = measure.wrap(ctx, text, font, avail);
  const blockHeight = lines.length * metrics.line;
  const top = verticalOffset(align.vertical, rect, blockHeight);
  const h = effectiveHAlign(align.horizontal, value);

  ctx.save();
  ctx.beginPath();
  ctx.rect(rect.x, rect.y, rect.width, rect.height);
  ctx.clip();
  measure.useFont(ctx, font);
  ctx.fillStyle = color;
  ctx.textBaseline = 'alphabetic';
  ctx.textAlign = 'left';

  for (let i = 0; i < lines.length; i++) {
    const lineText = lines[i] as string;
    const width = measure.measure(ctx, lineText, font);
    let x: number;
    switch (h) {
      case 'right':
        x = rect.x + rect.width - CELL_PAD - indent - width;
        break;
      case 'center':
        x = rect.x + (rect.width - width) / 2;
        break;
      default:
        x = rect.x + CELL_PAD + indent;
        break;
    }
    ctx.fillText(lineText, x, top + i * metrics.line + metrics.ascent);
  }
  ctx.restore();
}

/**
 * Rotated text. The file format's angles are not canvas angles: 1-90 means
 * counter-clockwise degrees, 91-180 means (value - 90) degrees CLOCKWISE, and
 * 255 means the characters are stacked vertically rather than rotated at all.
 */
function drawRotated(
  ctx: GridRenderingContext,
  input: RenderInput,
  text: string,
  font: string,
  color: string,
  rect: Rect,
  rotation: number,
  vertical: VerticalAlign | undefined,
  metrics: { ascent: number; descent: number; line: number },
): void {
  const { measure } = input;
  ctx.save();
  ctx.beginPath();
  ctx.rect(rect.x, rect.y, rect.width, rect.height);
  ctx.clip();
  measure.useFont(ctx, font);
  ctx.fillStyle = color;
  ctx.textBaseline = 'alphabetic';
  ctx.textAlign = 'left';

  if (rotation === STACKED_ROTATION) {
    const chars = [...text];
    const blockHeight = chars.length * metrics.line;
    let y = verticalOffset(vertical, rect, blockHeight) + metrics.ascent;
    for (const ch of chars) {
      const w = measure.measure(ctx, ch, font);
      ctx.fillText(ch, rect.x + (rect.width - w) / 2, y);
      y += metrics.line;
    }
    ctx.restore();
    return;
  }

  const degrees = rotation <= 90 ? -rotation : rotation - 90;
  const radians = (degrees * Math.PI) / 180;
  const width = measure.measure(ctx, text, font);
  // Anchor at the bottom-left for CCW rotation and the top-left for CW, which is
  // where Excel pins rotated text inside the cell.
  const anchorX = rect.x + CELL_PAD;
  const anchorY = degrees < 0 ? rect.y + rect.height - CELL_PAD : rect.y + CELL_PAD + metrics.ascent;
  ctx.translate(anchorX, anchorY);
  ctx.rotate(radians);
  ctx.fillText(text, 0, 0);
  ctx.restore();
  void width;
}

function indentStep(
  ctx: GridRenderingContext,
  measure: TextMeasureCache,
  font: string,
  fontSize: number,
): number {
  const space = measure.measure(ctx, ' ', font);
  return (space > 0 ? space : fontSize * 0.25) * INDENT_SPACES;
}

// --- selection ------------------------------------------------------------

function paintSelection(
  ctx: GridRenderingContext,
  input: RenderInput,
  pane: PaneLayout,
  selection: SelectionState,
): void {
  const { theme, layout, geometry } = input;
  const firstRow = (pane.rows[0] as Span).index;
  const lastRow = (pane.rows[pane.rows.length - 1] as Span).index;
  const firstCol = (pane.cols[0] as Span).index;
  const lastCol = (pane.cols[pane.cols.length - 1] as Span).index;

  const activeRect = cellRectOf(geometry, selection.active.row, selection.active.col);

  ctx.fillStyle = theme.selectionFill;
  for (const range of selection.ranges) {
    const top = Math.max(range.top, firstRow);
    const bottom = Math.min(range.bottom, lastRow);
    const left = Math.max(range.left, firstCol);
    const right = Math.min(range.right, lastCol);
    if (top > bottom || left > right) continue;
    const x = geometry.xOf(left);
    const y = geometry.yOf(top);
    const w = geometry.xOf(right) + geometry.cols.sizeOf(right) - x;
    const h = geometry.yOf(bottom) + geometry.rows.sizeOf(bottom) - y;
    const rect: Rect = { x, y, width: w, height: h };
    // Excel leaves the active cell unshaded so the cursor stays findable inside
    // a large selection; the shade is painted as up to four bands around it.
    const inside =
      selection.active.row >= top &&
      selection.active.row <= bottom &&
      selection.active.col >= left &&
      selection.active.col <= right;
    if (inside) fillAround(ctx, rect, activeRect);
    else ctx.fillRect(rect.x, rect.y, rect.width, rect.height);
  }

  ctx.setLineDash([]);
  ctx.strokeStyle = theme.selectionBorder;
  ctx.lineWidth = 1;
  for (const range of selection.ranges) {
    const x = geometry.xOf(range.left);
    const y = geometry.yOf(range.top);
    const w = geometry.xOf(range.right) + geometry.cols.sizeOf(range.right) - x;
    const h = geometry.yOf(range.bottom) + geometry.rows.sizeOf(range.bottom) - y;
    ctx.strokeRect(crisp(x, layout.dpr), crisp(y, layout.dpr), w, h);
  }

  ctx.strokeStyle = theme.activeCellBorder;
  ctx.lineWidth = 2;
  ctx.strokeRect(activeRect.x, activeRect.y, activeRect.width, activeRect.height);

  const primary = selection.ranges[selection.activeRange] ?? selection.ranges[0];
  if (primary) {
    const hx = geometry.xOf(primary.right) + geometry.cols.sizeOf(primary.right);
    const hy = geometry.yOf(primary.bottom) + geometry.rows.sizeOf(primary.bottom);
    ctx.fillStyle = theme.fillHandle;
    ctx.fillRect(hx - 3, hy - 3, 6, 6);
  }
}

function fillAround(ctx: GridRenderingContext, outer: Rect, hole: Rect): void {
  const holeRight = hole.x + hole.width;
  const holeBottom = hole.y + hole.height;
  const outerRight = outer.x + outer.width;
  const outerBottom = outer.y + outer.height;
  if (hole.y > outer.y) ctx.fillRect(outer.x, outer.y, outer.width, hole.y - outer.y);
  if (holeBottom < outerBottom) {
    ctx.fillRect(outer.x, holeBottom, outer.width, outerBottom - holeBottom);
  }
  if (hole.x > outer.x) ctx.fillRect(outer.x, hole.y, hole.x - outer.x, hole.height);
  if (holeRight < outerRight) {
    ctx.fillRect(holeRight, hole.y, outerRight - holeRight, hole.height);
  }
}

function cellRectOf(geometry: CellGeometry, row: number, col: number): Rect {
  return {
    x: geometry.xOf(col),
    y: geometry.yOf(row),
    width: geometry.cols.sizeOf(col),
    height: geometry.rows.sizeOf(row),
  };
}

// --- headers --------------------------------------------------------------

function paintHeaders(ctx: GridRenderingContext, input: RenderInput, stats: RenderStats): void {
  const { layout, theme, measure, selection } = input;
  const font = fontString({}, { family: theme.defaultFontFamily, size: theme.defaultFontSize });
  const headerFontSize = theme.defaultFontSize;
  const metrics = fontHeights(headerFontSize);

  ctx.save();
  ctx.beginPath();
  ctx.rect(layout.headerWidth, 0, layout.width - layout.headerWidth, layout.headerHeight);
  ctx.clip();
  measure.useFont(ctx, font);
  ctx.textBaseline = 'alphabetic';
  ctx.textAlign = 'left';

  for (const span of layout.colSpans) {
    const selected = selection ? isColSelected(selection, span.index) : false;
    const entire = selection ? isEntireColSelected(selection, span.index) : false;
    ctx.fillStyle = entire
      ? theme.headerBackgroundActive
      : selected
        ? theme.headerBackgroundSelected
        : theme.headerBackground;
    ctx.fillRect(span.start, 0, span.size, layout.headerHeight);
    const label = colToName(span.index);
    const w = measure.measure(ctx, label, font);
    ctx.fillStyle = selected ? theme.headerTextSelected : theme.headerText;
    ctx.fillText(
      label,
      span.start + (span.size - w) / 2,
      (layout.headerHeight + metrics.ascent - metrics.descent) / 2,
    );
    stats.headerCells++;
  }
  ctx.beginPath();
  ctx.strokeStyle = theme.headerLine;
  ctx.lineWidth = 1 / layout.dpr;
  for (const span of layout.colSpans) {
    const x = crisp(span.start + span.size, layout.dpr);
    ctx.moveTo(x, 0);
    ctx.lineTo(x, layout.headerHeight);
  }
  ctx.stroke();
  ctx.restore();

  ctx.save();
  ctx.beginPath();
  ctx.rect(0, layout.headerHeight, layout.headerWidth, layout.height - layout.headerHeight);
  ctx.clip();
  measure.useFont(ctx, font);
  ctx.textBaseline = 'alphabetic';
  ctx.textAlign = 'left';

  for (const span of layout.rowSpans) {
    const selected = selection ? isRowSelected(selection, span.index) : false;
    const entire = selection ? isEntireRowSelected(selection, span.index) : false;
    ctx.fillStyle = entire
      ? theme.headerBackgroundActive
      : selected
        ? theme.headerBackgroundSelected
        : theme.headerBackground;
    ctx.fillRect(0, span.start, layout.headerWidth, span.size);
    const label = String(span.index + 1);
    const w = measure.measure(ctx, label, font);
    ctx.fillStyle = selected ? theme.headerTextSelected : theme.headerText;
    ctx.fillText(
      label,
      (layout.headerWidth - w) / 2,
      span.start + (span.size + metrics.ascent - metrics.descent) / 2,
    );
    stats.headerCells++;
  }
  ctx.beginPath();
  ctx.strokeStyle = theme.headerLine;
  ctx.lineWidth = 1 / layout.dpr;
  for (const span of layout.rowSpans) {
    const y = crisp(span.start + span.size, layout.dpr);
    ctx.moveTo(0, y);
    ctx.lineTo(layout.headerWidth, y);
  }
  ctx.stroke();
  ctx.restore();

  ctx.fillStyle = theme.cornerBackground;
  ctx.fillRect(0, 0, layout.headerWidth, layout.headerHeight);
  stats.headerCells++;

  ctx.beginPath();
  ctx.strokeStyle = theme.headerEdge;
  ctx.lineWidth = 1 / layout.dpr;
  const ex = crisp(layout.headerWidth, layout.dpr);
  const ey = crisp(layout.headerHeight, layout.dpr);
  ctx.moveTo(ex, 0);
  ctx.lineTo(ex, layout.height);
  ctx.moveTo(0, ey);
  ctx.lineTo(layout.width, ey);
  ctx.stroke();
}

function isRowSelected(s: SelectionState, row: number): boolean {
  for (const r of s.ranges) if (row >= r.top && row <= r.bottom) return true;
  return false;
}

function isColSelected(s: SelectionState, col: number): boolean {
  for (const r of s.ranges) if (col >= r.left && col <= r.right) return true;
  return false;
}

function isEntireRowSelected(s: SelectionState, row: number): boolean {
  for (const r of s.ranges) {
    if (row >= r.top && row <= r.bottom && r.left === 0 && r.right === MAX_COLS - 1) return true;
  }
  return false;
}

function isEntireColSelected(s: SelectionState, col: number): boolean {
  for (const r of s.ranges) {
    if (col >= r.left && col <= r.right && r.top === 0 && r.bottom === MAX_ROWS - 1) return true;
  }
  return false;
}

function paintFrozenDividers(ctx: GridRenderingContext, input: RenderInput): void {
  const { layout, theme } = input;
  if (layout.frozenColCount === 0 && layout.frozenRowCount === 0) return;
  ctx.beginPath();
  ctx.strokeStyle = theme.frozenLine;
  ctx.lineWidth = 1 / layout.dpr;
  ctx.setLineDash([]);
  if (layout.frozenColCount > 0) {
    const x = crisp(layout.headerWidth + layout.frozenWidth, layout.dpr);
    ctx.moveTo(x, 0);
    ctx.lineTo(x, layout.height);
  }
  if (layout.frozenRowCount > 0) {
    const y = crisp(layout.headerHeight + layout.frozenHeight, layout.dpr);
    ctx.moveTo(0, y);
    ctx.lineTo(layout.width, y);
  }
  ctx.stroke();
}

// --- helpers --------------------------------------------------------------

export function fillColorOf(
  style: CellStyle,
  palette: readonly string[] | undefined,
): string | undefined {
  const fill = style.fill;
  if (!fill || !fill.pattern || fill.pattern === 'none') return undefined;
  // In a solid fill the *foreground* colour is the visible one. Reading bg here
  // is the classic xlsx mistake and paints every shaded cell white.
  const color = fill.pattern === 'solid' ? (fill.fg ?? fill.bg) : (fill.bg ?? fill.fg);
  return resolveColor(color, palette);
}

/**
 * Snap a coordinate so a one-device-pixel line lands on a pixel centre.
 * Without this every gridline is a two-pixel grey smear at dpr 1 and a soft
 * double line at dpr 2.
 */
export function crisp(value: number, dpr: number): number {
  return Math.round(value * dpr) / dpr + 0.5 / dpr;
}
