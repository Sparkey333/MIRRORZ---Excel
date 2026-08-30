/**
 * The application layout.
 *
 * Toolbar, formula bar, then a row of grid and side panels, then the tab strip
 * and status bar. The panels are siblings of the grid rather than overlays
 * because a floating panel over a spreadsheet covers the cells being inspected,
 * which is precisely the wrong thing for a panel whose whole job is to talk
 * about a cell.
 */

import { useCallback, useState } from 'react';
import { useApp, useController } from './state/context.js';
import { useTheme } from './state/useTheme.js';
import { useKeyboard } from './state/useKeyboard.js';
import { useDocumentState, useShell } from './state/useShell.js';
import {
  isDelimitedName,
  readDelimited,
  resolveHost,
  serializeWorkbook,
} from './state/host.js';
import { Toolbar } from './components/Toolbar.js';
import { FormulaBar } from './components/FormulaBar.js';
import { GridHost } from './components/GridHost.js';
import { SheetTabs } from './components/SheetTabs.js';
import { SheetExplorer } from './components/SheetExplorer.js';
import { HistoryPanel } from './components/HistoryPanel.js';
import { FormulaInspector } from './components/FormulaInspector.js';
import { CommandPalette } from './components/CommandPalette.js';
import { ImportReviewDialog } from './components/ImportReviewDialog.js';
import { StatusBar } from './components/StatusBar.js';

export interface AppProps {
  /** Called when a workbook file is opened, since that replaces the controller. */
  onOpenWorkbook?: (file: { name: string; data: Uint8Array }) => void;
}

export function App({ onOpenWorkbook }: AppProps = {}) {
  const controller = useController();
  const snapshot = useApp();
  useTheme(snapshot.theme);
  useKeyboard(controller);
  const [host] = useState(() => resolveHost());

  /**
   * Take in a file, whichever way it arrived.
   *
   * The Open button and the operating system reach exactly the same code: a
   * double-clicked spreadsheet has to behave identically to one picked from the
   * dialog, and the only way to guarantee that is for there to be one path.
   */
  const acceptFile = useCallback(
    (file: { name: string; data: Uint8Array }) => {
      // A delimited file is data going into the current workbook and goes
      // through the import review; a workbook file replaces the document.
      if (isDelimitedName(file.name)) {
        const { rows } = readDelimited(file);
        const width = rows.reduce((max, row) => Math.max(max, row.length), 0);
        controller.proposeImport({
          anchor: controller.activeAddr(),
          rows,
          overrides: Array.from({ length: width }, () => 'auto' as const),
          headerRow: true,
          source: 'csv',
          fileName: file.name,
        });
        return;
      }
      onOpenWorkbook?.(file);
    },
    [controller, onOpenWorkbook],
  );

  const openFile = useCallback(async () => {
    const file = await host.openFile();
    if (file) acceptFile(file);
  }, [acceptFile, host]);

  const save = useCallback(
    async (as: boolean) => {
      const name = snapshot.fileName.endsWith('.xlsx') ? snapshot.fileName : `${snapshot.fileName}.xlsx`;
      const bytes = serializeWorkbook(controller.workbook);
      const saved = as ? await host.saveFileAs(name, bytes) : await host.saveFile(name, bytes);
      if (saved) {
        if (typeof saved === 'string') controller.setFileName(saved);
        controller.markSaved();
      }
    },
    [controller, host, snapshot.fileName],
  );

  // The native menu, files the operating system pushes at us, and the autosave
  // clock. Main has been sending all three since it was written; this is the
  // end that receives them.
  useShell(controller, {
    openFile: () => void openFile(),
    save: (as) => void save(as),
    acceptFile,
    serialize: () => serializeWorkbook(controller.workbook),
  });

  // Main gates the unsaved-changes prompt and the autosave ticker on this, so a
  // window whose renderer never reports it closes on unsaved work in silence.
  useDocumentState(snapshot.dirty, snapshot.fileName);

  return (
    <div className="mz-app" data-theme={snapshot.theme}>
      <Toolbar
        onOpenFile={() => void openFile()}
        onSave={() => void save(false)}
        onSaveAs={() => void save(true)}
      />
      <FormulaBar />
      <div className="mz-body">
        {snapshot.panels.explorer ? <SheetExplorer /> : null}
        <main className="mz-main">
          <GridHost />
        </main>
        {snapshot.panels.inspector ? <FormulaInspector /> : null}
        {snapshot.panels.history ? <HistoryPanel /> : null}
      </div>
      <SheetTabs />
      <StatusBar />
      <CommandPalette />
      <ImportReviewDialog />
    </div>
  );
}
