/**
 * Renderer entry point.
 *
 * The controller lives in React state so that opening a workbook can replace it
 * wholesale: a new file is a new Document, and reusing the old one would carry
 * the previous file's undo history into it.
 */

import { StrictMode, useState } from 'react';
import { createRoot } from 'react-dom/client';
import { App } from './App.js';
import { AppProvider, createController } from './state/context.js';
import { openWorkbook } from './state/host.js';
import type { AppController } from './state/controller.js';
import './styles/app.css';

function Root() {
  const [controller, setController] = useState<AppController>(() => createController());

  return (
    <AppProvider controller={controller}>
      <App
        onOpenWorkbook={(file) => {
          const opened = openWorkbook(file);
          controller.dispose();
          setController(opened.controller);
          if (opened.warnings.length > 0) {
            opened.controller.setMessage(
              `Opened with ${opened.warnings.length} warning${opened.warnings.length === 1 ? '' : 's'}`,
            );
          }
        }}
      />
    </AppProvider>
  );
}

const container = document.getElementById('root');
if (container) {
  createRoot(container).render(
    <StrictMode>
      <Root />
    </StrictMode>,
  );
}
