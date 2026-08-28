/**
 * The shell channel, from the renderer's side.
 *
 * Main has always sent menu commands, pushed files and run the autosave clock,
 * and the renderer listened to none of it - so the native menu did nothing, a
 * double-clicked file opened a blank window, and closing a dirty window threw
 * the work away without a prompt, because the prompt is gated on a flag nobody
 * ever set. These tests hold that contract from this end.
 */

import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { act } from '@testing-library/react';
import { App } from '../src/renderer/App.js';
import { runMenuCommand, bytesToBase64, base64ToBytes } from '../src/renderer/state/useShell.js';
import type { AppController } from '../src/renderer/state/controller.js';
import { createController } from '../src/renderer/state/context.js';
import { enter, renderWith } from './render.js';

type CommandListener = (command: string, arg?: string) => void;
type OpenListener = (file: { name: string; data: Uint8Array }) => void;

interface FakeShell {
  commands: CommandListener[];
  opens: OpenListener[];
  states: { dirty: boolean; displayName?: string }[];
  journals: string[];
}

let fake: FakeShell;

beforeEach(() => {
  fake = { commands: [], opens: [], states: [], journals: [] };
  window.mirrorzShell = {
    onCommand(callback) {
      fake.commands.push(callback);
      return () => {
        fake.commands = fake.commands.filter((c) => c !== callback);
      };
    },
    setDocumentState(state) {
      fake.states.push(state);
    },
    async autosave(payload) {
      fake.journals.push(payload);
      return true;
    },
  };
  window.mirrorz = {
    onOpenRequest(callback) {
      fake.opens.push(callback);
      return () => {
        fake.opens = fake.opens.filter((c) => c !== callback);
      };
    },
  };
});

afterEach(() => {
  delete window.mirrorzShell;
  delete window.mirrorz;
});

function send(command: string, arg?: string): void {
  act(() => {
    for (const listener of [...fake.commands]) listener(command, arg);
  });
}

describe('the shell channel is subscribed', () => {
  it('reports unsaved changes to main, which is what gates the close prompt', () => {
    const { controller } = renderWith(() => <App />);
    expect(fake.states.at(-1)).toEqual({ dirty: false, displayName: 'Untitled' });

    act(() => controller.commitEntry('42', { sheet: 'Sheet1', row: 0, col: 0 }));
    expect(fake.states.at(-1)?.dirty).toBe(true);

    act(() => controller.markSaved());
    expect(fake.states.at(-1)?.dirty).toBe(false);
  });

  it('runs a menu command', () => {
    const { controller } = renderWith(() => <App />, (c) => enter(c, 0, 0, '42'));
    expect(fake.commands.length).toBeGreaterThan(0);

    send('edit.undo');
    expect(controller.cellAt({ sheet: 'Sheet1', row: 0, col: 0 })).toBeUndefined();
    send('edit.redo');
    expect(controller.cellAt({ sheet: 'Sheet1', row: 0, col: 0 })?.value).toBe(42);
  });

  it('answers the autosave tick with a journal rather than an error message', () => {
    const { controller } = renderWith(() => <App />, (c) => enter(c, 0, 0, '42'));
    send('app.autosaveTick');
    expect(fake.journals).toHaveLength(1);
    expect(fake.journals[0]!.length).toBeGreaterThan(0);
    // The tick is not a menu item and must not be reported as an unknown one.
    expect(controller.getSnapshot().message ?? '').not.toContain('Unrecognised');
  });

  it('takes in a file the shell pushes at it', () => {
    const received: { name: string }[] = [];
    renderWith(() => <App onOpenWorkbook={(file) => received.push({ name: file.name })} />);
    expect(fake.opens.length).toBeGreaterThan(0);
    act(() => {
      for (const listener of fake.opens) listener({ name: 'Book.xlsx', data: new Uint8Array([1]) });
    });
    expect(received).toEqual([{ name: 'Book.xlsx' }]);
  });

  it('says so when a menu item is not built yet, rather than doing nothing', () => {
    const { controller } = renderWith(() => <App />);
    send('data.removeDuplicates');
    expect(controller.getSnapshot().message).toBe('That is not built yet');
  });

  it('unsubscribes when the app unmounts', () => {
    const { view } = renderWith(() => <App />);
    expect(fake.commands.length).toBe(1);
    view.unmount();
    expect(fake.commands.length).toBe(0);
    expect(fake.opens.length).toBe(0);
  });
});

describe('runMenuCommand', () => {
  let controller: AppController;
  const actions = {
    openFile: vi.fn(),
    save: vi.fn(),
    acceptFile: vi.fn(),
    serialize: () => new Uint8Array(0),
  };

  beforeEach(() => {
    controller = createController(undefined, { now: () => 1 });
    actions.openFile.mockClear();
    actions.save.mockClear();
  });

  it('routes save and save-as to different calls', () => {
    expect(runMenuCommand(controller, actions, 'file.save')).toBe('ran');
    expect(actions.save).toHaveBeenCalledWith(false);
    expect(runMenuCommand(controller, actions, 'file.saveAs')).toBe('ran');
    expect(actions.save).toHaveBeenCalledWith(true);
  });

  it('opens find for both find and replace', () => {
    runMenuCommand(controller, actions, 'edit.find');
    expect(controller.getSnapshot().findOpen).toBe(true);
    controller.setFind(false);
    runMenuCommand(controller, actions, 'edit.replace');
    expect(controller.getSnapshot().findOpen).toBe(true);
  });

  it('inserts below the selection, not above it', () => {
    controller.commitEntry('a', { sheet: 'Sheet1', row: 0, col: 0 });
    controller.commitEntry('b', { sheet: 'Sheet1', row: 1, col: 0 });
    controller.selectCell(0, 0);
    runMenuCommand(controller, actions, 'insert.rowsBelow');
    expect(controller.cellAt({ sheet: 'Sheet1', row: 0, col: 0 })?.value).toBe('a');
    expect(controller.cellAt({ sheet: 'Sheet1', row: 2, col: 0 })?.value).toBe('b');
  });

  it('applies a number format from the menu', () => {
    controller.commitEntry('0.5', { sheet: 'Sheet1', row: 0, col: 0 });
    controller.selectCell(0, 0);
    runMenuCommand(controller, actions, 'format.numberPercent');
    expect(controller.displayText({ sheet: 'Sheet1', row: 0, col: 0 })).toContain('%');
  });

  it('separates an unbuilt item from an unrecognised one', () => {
    expect(runMenuCommand(controller, actions, 'data.textToColumns')).toBe('unimplemented');
    expect(runMenuCommand(controller, actions, 'nonsense.command')).toBe('unknown');
  });

  it('leaves window-level commands to main', () => {
    // `file.new` and `file.close` are about the window, not the document.
    expect(runMenuCommand(controller, actions, 'file.close')).toBe('unknown');
  });
});

describe('journal encoding', () => {
  it('round-trips bytes, including ones no text encoding survives', () => {
    const bytes = new Uint8Array([0, 1, 127, 128, 200, 255, 0x50, 0x4b]);
    expect([...base64ToBytes(bytesToBase64(bytes))]).toEqual([...bytes]);
  });

  it('handles a payload larger than one chunk', () => {
    const bytes = new Uint8Array(0x8000 * 2 + 17);
    for (let i = 0; i < bytes.length; i++) bytes[i] = i % 256;
    expect(base64ToBytes(bytesToBase64(bytes)).length).toBe(bytes.length);
  });
});
