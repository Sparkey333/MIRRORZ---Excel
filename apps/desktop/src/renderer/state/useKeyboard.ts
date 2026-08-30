/**
 * The global keyboard map.
 *
 * Shortcuts are bound once, at the window, and every one of them dispatches a
 * command from the same catalogue the palette and toolbar use. Binding them to
 * individual components is how an application ends up with a shortcut that works
 * only when the right thing happens to have focus.
 *
 * Nothing fires while the user is typing in a field, except the ones that have
 * to: escape, and the palette itself.
 */

import { useEffect } from 'react';
import type { AppController } from './controller.js';

const TEXT_INPUTS = new Set(['INPUT', 'TEXTAREA']);

function isTyping(target: EventTarget | null): boolean {
  if (!(target instanceof HTMLElement)) return false;
  return TEXT_INPUTS.has(target.tagName) || target.isContentEditable;
}

export function useKeyboard(controller: AppController): void {
  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent): void => {
      const mod = event.ctrlKey || event.metaKey;
      const typing = isTyping(event.target);

      if (mod && event.key.toLowerCase() === 'p') {
        event.preventDefault();
        controller.setPalette(true, event.shiftKey ? 'command' : 'all');
        return;
      }
      if (event.key === 'Escape') {
        const snapshot = controller.getSnapshot();
        if (snapshot.paletteOpen) controller.setPalette(false);
        else if (snapshot.pendingImport) controller.cancelImport();
        return;
      }
      if (typing) return;

      if (mod && event.key.toLowerCase() === 'z' && !event.shiftKey) {
        event.preventDefault();
        controller.undo();
        return;
      }
      if ((mod && event.key.toLowerCase() === 'y') || (mod && event.shiftKey && event.key.toLowerCase() === 'z')) {
        event.preventDefault();
        controller.redo();
        return;
      }
      if (mod && event.shiftKey && event.key.toLowerCase() === 'h') {
        event.preventDefault();
        controller.togglePanel('history');
        return;
      }
      if (mod && event.shiftKey && event.key.toLowerCase() === 'i') {
        event.preventDefault();
        controller.togglePanel('inspector');
        return;
      }
      if (mod && event.shiftKey && event.key.toLowerCase() === 'e') {
        event.preventDefault();
        controller.togglePanel('explorer');
        return;
      }
      if (event.key === 'F9') {
        event.preventDefault();
        controller.recalculateAll();
        return;
      }
      if (event.key === 'Delete' || event.key === 'Backspace') {
        event.preventDefault();
        controller.clearSelection();
        return;
      }
      if (event.key.startsWith('Arrow')) {
        event.preventDefault();
        const { active } = controller.getSnapshot().selection;
        const drow = event.key === 'ArrowDown' ? 1 : event.key === 'ArrowUp' ? -1 : 0;
        const dcol = event.key === 'ArrowRight' ? 1 : event.key === 'ArrowLeft' ? -1 : 0;
        controller.selectCell(Math.max(0, active.row + drow), Math.max(0, active.col + dcol));
      }
    };

    window.addEventListener('keydown', onKeyDown);
    return () => window.removeEventListener('keydown', onKeyDown);
  }, [controller]);
}
