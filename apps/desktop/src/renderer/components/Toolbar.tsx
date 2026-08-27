/**
 * The toolbar.
 *
 * Not a ribbon. Copying Microsoft's ribbon would be a trademark problem and a
 * design mistake at the same time: the ribbon's tabs hide two thirds of the
 * commands behind a mode switch, which is why every user eventually memorises
 * keyboard shortcuts instead. This is one row of labelled groups, always
 * visible, with the command palette carrying the long tail.
 *
 * Every control reads its state from the selection rather than from a local
 * toggle, so the bold button is pressed when the selection is bold - including
 * after an undo, which is exactly where a local toggle would go stale.
 */

import { useState } from 'react';
import type { Color } from '@mirrorz/core';
import type { AppController } from '../state/controller.js';
import { useApp, useController, useDerived } from '../state/context.js';
import { adjustDecimals, presetsByCategory, presetSample, validateFormatCode } from '../model/number-formats.js';
import { nextPreference } from '../model/theme.js';
import { FindReplace } from './FindReplace.js';

const FONT_FAMILIES = [
  'Calibri',
  'Arial',
  'Helvetica',
  'Times New Roman',
  'Georgia',
  'Courier New',
  'Consolas',
  'Verdana',
];
const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72];

const TEXT_COLORS = ['#1b1f27', '#c0392b', '#1a7f43', '#1f6feb', '#7c3aed', '#a86a00', '#ffffff'];
const FILL_COLORS = ['#fff3bf', '#d3f9d8', '#d0ebff', '#ffe3e3', '#f3f0ff', '#f1f3f5', '#ffffff'];

function rgb(hex: string): Color {
  return { kind: 'rgb', argb: `FF${hex.replace('#', '').toUpperCase()}` };
}

export function Toolbar({ onOpenFile, onSave, onSaveAs }: {
  onOpenFile?: () => void;
  onSave?: () => void;
  onSaveAs?: () => void;
}) {
  const controller = useController();
  const snapshot = useApp();
  const [formatMenuOpen, setFormatMenuOpen] = useState(false);
  const [customCode, setCustomCode] = useState('');
  const [findOpen, setFindOpen] = useState(false);

  const font = useDerived(() => controller.uniformStyle((s) => s.font));
  const numFmt = useDerived(() => controller.uniformStyle((s) => s.numFmt));
  const alignment = useDerived(() => controller.uniformStyle((s) => s.alignment));

  const customValidation = customCode === '' ? null : validateFormatCode(customCode);

  return (
    <div className="mz-toolbar" role="toolbar" aria-label="Main toolbar">
      <ToolbarGroup label="File">
        <button type="button" onClick={onOpenFile} disabled={!onOpenFile}>
          Open
        </button>
        <button type="button" onClick={onSave} disabled={!onSave}>
          Save
        </button>
        <button type="button" onClick={onSaveAs} disabled={!onSaveAs}>
          Save as
        </button>
      </ToolbarGroup>

      <ToolbarGroup label="Clipboard">
        <button type="button" onClick={() => void copySelection(controller)} aria-label="Copy">
          Copy
        </button>
        <button type="button" onClick={() => void cutSelection(controller)} aria-label="Cut">
          Cut
        </button>
        <button type="button" onClick={() => void pasteFromClipboard(controller)} aria-label="Paste">
          Paste
        </button>
      </ToolbarGroup>

      <ToolbarGroup label="Font">
        <select
          aria-label="Font family"
          value={font?.name ?? 'Calibri'}
          onChange={(e) => controller.applyStyle({ font: { name: e.target.value } }, 'Change font')}
        >
          {FONT_FAMILIES.map((family) => (
            <option key={family} value={family}>
              {family}
            </option>
          ))}
        </select>
        <select
          aria-label="Font size"
          value={font?.size ?? 11}
          onChange={(e) =>
            controller.applyStyle({ font: { size: Number(e.target.value) } }, 'Change font size')
          }
        >
          {FONT_SIZES.map((size) => (
            <option key={size} value={size}>
              {size}
            </option>
          ))}
        </select>
        <button
          type="button"
          aria-label="Bold"
          aria-pressed={font?.bold === true}
          className={font?.bold ? 'mz-active' : undefined}
          onClick={() => controller.applyStyle({ font: { bold: !font?.bold } }, 'Bold')}
        >
          <strong>B</strong>
        </button>
        <button
          type="button"
          aria-label="Italic"
          aria-pressed={font?.italic === true}
          className={font?.italic ? 'mz-active' : undefined}
          onClick={() => controller.applyStyle({ font: { italic: !font?.italic } }, 'Italic')}
        >
          <em>I</em>
        </button>
        <button
          type="button"
          aria-label="Underline"
          aria-pressed={font?.underline === 'single'}
          className={font?.underline === 'single' ? 'mz-active' : undefined}
          onClick={() =>
            controller.applyStyle(
              { font: { underline: font?.underline === 'single' ? 'none' : 'single' } },
              'Underline',
            )
          }
        >
          <u>U</u>
        </button>
        <ColorPicker
          label="Text colour"
          colors={TEXT_COLORS}
          onPick={(hex) => controller.applyStyle({ font: { color: rgb(hex) } }, 'Text colour')}
        />
        <ColorPicker
          label="Fill colour"
          colors={FILL_COLORS}
          onPick={(hex) =>
            controller.applyStyle({ fill: { pattern: 'solid', fg: rgb(hex) } }, 'Fill colour')
          }
        />
      </ToolbarGroup>

      <ToolbarGroup label="Alignment">
        {(['left', 'center', 'right'] as const).map((horizontal) => (
          <button
            key={horizontal}
            type="button"
            aria-label={`Align ${horizontal}`}
            aria-pressed={alignment?.horizontal === horizontal}
            className={alignment?.horizontal === horizontal ? 'mz-active' : undefined}
            onClick={() => controller.applyStyle({ alignment: { horizontal } }, `Align ${horizontal}`)}
          >
            <span className="mz-align-glyph" data-align={horizontal} aria-hidden="true" />
            <span className="mz-visually-hidden">{horizontal}</span>
          </button>
        ))}
        <button
          type="button"
          aria-label="Wrap text"
          aria-pressed={alignment?.wrapText === true}
          className={alignment?.wrapText ? 'mz-active' : undefined}
          onClick={() =>
            controller.applyStyle({ alignment: { wrapText: !alignment?.wrapText } }, 'Wrap text')
          }
        >
          Wrap
        </button>
      </ToolbarGroup>

      <ToolbarGroup label="Number">
        <div className="mz-menu-anchor">
          <button
            type="button"
            aria-haspopup="menu"
            aria-expanded={formatMenuOpen}
            onClick={() => setFormatMenuOpen((open) => !open)}
          >
            {numFmt ?? 'General'}
          </button>
          {formatMenuOpen ? (
            <div className="mz-menu" role="menu" aria-label="Number format">
              {presetsByCategory().map((group) => (
                <div key={group.category} className="mz-menu-group">
                  <div className="mz-menu-group-label">{group.category}</div>
                  {group.presets.map((preset) => (
                    <button
                      key={preset.id}
                      type="button"
                      role="menuitemradio"
                      aria-checked={(preset.code ?? undefined) === numFmt}
                      onClick={() => {
                        controller.setNumberFormat(preset.code);
                        setFormatMenuOpen(false);
                      }}
                    >
                      <span className="mz-menu-item-label">{preset.label}</span>
                      <span className="mz-menu-item-sample">{presetSample(preset)}</span>
                    </button>
                  ))}
                </div>
              ))}
              <div className="mz-menu-group">
                <div className="mz-menu-group-label">custom</div>
                <input
                  aria-label="Custom format code"
                  value={customCode}
                  placeholder="#,##0.00;[Red]-#,##0.00"
                  onChange={(e) => setCustomCode(e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key !== 'Enter' || !customValidation?.valid) return;
                    controller.setNumberFormat(customCode);
                    setFormatMenuOpen(false);
                  }}
                />
                {customValidation ? (
                  <div
                    className={customValidation.valid ? 'mz-menu-hint' : 'mz-menu-hint mz-invalid'}
                    role="status"
                  >
                    {customValidation.valid
                      ? `Sample: ${customValidation.preview}`
                      : customValidation.problem}
                  </div>
                ) : null}
              </div>
            </div>
          ) : null}
        </div>
        <button
          type="button"
          aria-label="Add decimal place"
          onClick={() => controller.setNumberFormat(adjustDecimals(numFmt, 1))}
        >
          .0&#8594;
        </button>
        <button
          type="button"
          aria-label="Remove decimal place"
          onClick={() => controller.setNumberFormat(adjustDecimals(numFmt, -1))}
        >
          &#8592;.0
        </button>
      </ToolbarGroup>

      <ToolbarGroup label="Cells">
        <button type="button" onClick={() => controller.insertRows(controller.selectionRange().start.row)}>
          Insert row
        </button>
        <button type="button" onClick={() => controller.deleteRows(controller.selectionRange().start.row)}>
          Delete row
        </button>
        <button type="button" onClick={() => controller.insertCols(controller.selectionRange().start.col)}>
          Insert column
        </button>
        <button type="button" onClick={() => controller.deleteCols(controller.selectionRange().start.col)}>
          Delete column
        </button>
      </ToolbarGroup>

      <ToolbarGroup label="Data">
        <button
          type="button"
          aria-label="Sort ascending"
          onClick={() => controller.sortSelection(controller.selectionRange().start.col, 'asc')}
        >
          A&#8594;Z
        </button>
        <button
          type="button"
          aria-label="Sort descending"
          onClick={() => controller.sortSelection(controller.selectionRange().start.col, 'desc')}
        >
          Z&#8594;A
        </button>
        <button type="button" onClick={() => setFindOpen((open) => !open)} aria-expanded={findOpen}>
          Find
        </button>
      </ToolbarGroup>

      <ToolbarGroup label="View">
        <button
          type="button"
          aria-pressed={snapshot.panels.history}
          className={snapshot.panels.history ? 'mz-active' : undefined}
          onClick={() => controller.togglePanel('history')}
        >
          History
        </button>
        <button
          type="button"
          aria-pressed={snapshot.panels.inspector}
          className={snapshot.panels.inspector ? 'mz-active' : undefined}
          onClick={() => controller.togglePanel('inspector')}
        >
          Inspector
        </button>
        <button
          type="button"
          aria-label={`Theme: ${snapshot.theme}`}
          onClick={() => controller.setTheme(nextPreference(snapshot.theme))}
        >
          {snapshot.theme === 'system' ? 'Auto' : snapshot.theme === 'dark' ? 'Dark' : 'Light'}
        </button>
      </ToolbarGroup>

      {findOpen ? <FindReplace onClose={() => setFindOpen(false)} /> : null}
    </div>
  );
}

function ToolbarGroup({ label, children }: { label: string; children: React.ReactNode }) {
  return (
    <div className="mz-toolbar-group" role="group" aria-label={label}>
      <div className="mz-toolbar-buttons">{children}</div>
      <div className="mz-toolbar-group-label">{label}</div>
    </div>
  );
}

function ColorPicker({
  label,
  colors,
  onPick,
}: {
  label: string;
  colors: readonly string[];
  onPick: (hex: string) => void;
}) {
  const [open, setOpen] = useState(false);
  return (
    <div className="mz-menu-anchor">
      <button type="button" aria-label={label} aria-expanded={open} onClick={() => setOpen((o) => !o)}>
        {label === 'Fill colour' ? '■' : 'A'}
      </button>
      {open ? (
        <div className="mz-swatches" role="menu" aria-label={label}>
          {colors.map((hex) => (
            <button
              key={hex}
              type="button"
              role="menuitem"
              aria-label={`${label} ${hex}`}
              style={{ background: hex }}
              onClick={() => {
                onPick(hex);
                setOpen(false);
              }}
            />
          ))}
        </div>
      ) : null}
    </div>
  );
}

/**
 * Clipboard operations use the tab-separated text convention every spreadsheet
 * speaks, so copy and paste work with Excel, Numbers and a text editor without
 * a private format.
 */
async function copySelection(controller: AppController): Promise<void> {
  const range = controller.selectionRange();
  const lines: string[] = [];
  for (let r = range.start.row; r <= range.end.row; r++) {
    const cells: string[] = [];
    for (let c = range.start.col; c <= range.end.col; c++) {
      cells.push(controller.displayText({ sheet: controller.getSnapshot().activeSheet, row: r, col: c }));
    }
    lines.push(cells.join('\t'));
  }
  await navigator.clipboard?.writeText(lines.join('\n'));
}

async function cutSelection(controller: AppController): Promise<void> {
  await copySelection(controller);
  controller.clearSelection();
}

/**
 * A paste is proposed, never applied directly: the import review is the whole
 * point, and a paste is the commonest way data gets silently mangled.
 */
async function pasteFromClipboard(controller: AppController): Promise<void> {
  const text = await navigator.clipboard?.readText();
  if (!text) return;
  const rows = text.replace(/\r\n?/g, '\n').split('\n').map((line) => line.split('\t'));
  const width = rows.reduce((max, row) => Math.max(max, row.length), 0);
  controller.proposeImport({
    anchor: controller.activeAddr(),
    rows,
    overrides: Array.from({ length: width }, () => 'auto' as const),
    headerRow: false,
    source: 'paste',
  });
}
