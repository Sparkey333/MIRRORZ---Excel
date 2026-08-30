/**
 * A recording 2D context.
 *
 * Deliberately not a canvas implementation: the tests here assert what the
 * painter ASKED for, not what pixels came out, and a real rasteriser would make
 * "did it paint 800 cells or 800,000" impossible to answer. Text widths are a
 * deterministic function of the font size so alignment and overflow maths can be
 * asserted exactly.
 */

import type { GridRenderingContext } from '../src/render.js';

export interface DrawCall {
  op: string;
  args: readonly unknown[];
  /** State at the time of the call, for asserting colours and fonts. */
  fillStyle: string;
  strokeStyle: string;
  font: string;
  lineWidth: number;
}

/** Width of one character as a fraction of the font size. */
export const CHAR_RATIO = 0.5;

export class FakeContext implements GridRenderingContext {
  font = '10px sans-serif';
  fillStyle: string | CanvasGradient | CanvasPattern = '#000000';
  strokeStyle: string | CanvasGradient | CanvasPattern = '#000000';
  lineWidth = 1;
  textAlign: CanvasTextAlign = 'start';
  textBaseline: CanvasTextBaseline = 'alphabetic';
  globalAlpha = 1;

  readonly calls: DrawCall[] = [];
  measureCount = 0;
  fontAssignments = 0;

  private record(op: string, ...args: unknown[]): void {
    this.calls.push({
      op,
      args,
      fillStyle: String(this.fillStyle),
      strokeStyle: String(this.strokeStyle),
      font: this.font,
      lineWidth: this.lineWidth,
    });
  }

  measureText(text: string): { width: number } {
    this.measureCount++;
    const m = /(\d+(?:\.\d+)?)px/.exec(this.font);
    const size = m ? Number(m[1]) : 10;
    const bold = this.font.includes('bold') ? 1.1 : 1;
    return { width: text.length * size * CHAR_RATIO * bold };
  }

  save(): void {
    this.record('save');
  }
  restore(): void {
    this.record('restore');
  }
  beginPath(): void {
    this.record('beginPath');
  }
  closePath(): void {
    this.record('closePath');
  }
  rect(x: number, y: number, w: number, h: number): void {
    this.record('rect', x, y, w, h);
  }
  clip(): void {
    this.record('clip');
  }
  moveTo(x: number, y: number): void {
    this.record('moveTo', x, y);
  }
  lineTo(x: number, y: number): void {
    this.record('lineTo', x, y);
  }
  stroke(): void {
    this.record('stroke');
  }
  fill(): void {
    this.record('fill');
  }
  fillRect(x: number, y: number, w: number, h: number): void {
    this.record('fillRect', x, y, w, h);
  }
  strokeRect(x: number, y: number, w: number, h: number): void {
    this.record('strokeRect', x, y, w, h);
  }
  fillText(text: string, x: number, y: number, maxWidth?: number): void {
    this.record('fillText', text, x, y, maxWidth);
  }
  translate(x: number, y: number): void {
    this.record('translate', x, y);
  }
  rotate(angle: number): void {
    this.record('rotate', angle);
  }
  setTransform(a: number, b: number, c: number, d: number, e: number, f: number): void {
    this.record('setTransform', a, b, c, d, e, f);
  }
  setLineDash(segments: number[]): void {
    this.record('setLineDash', segments.slice());
  }

  ops(op: string): DrawCall[] {
    return this.calls.filter((c) => c.op === op);
  }

  texts(): string[] {
    return this.ops('fillText').map((c) => String(c.args[0]));
  }

  reset(): void {
    this.calls.length = 0;
    this.measureCount = 0;
  }
}

export class FakeCanvas {
  width = 0;
  height = 0;
  style = { width: '', height: '' };
  readonly context = new FakeContext();

  getContext(contextId: '2d'): GridRenderingContext | null {
    return contextId === '2d' ? this.context : null;
  }
}

/** Expected width of a string under FakeContext's measurement rule. */
export function expectedWidth(text: string, fontSize: number, bold = false): number {
  return text.length * fontSize * CHAR_RATIO * (bold ? 1.1 : 1);
}
