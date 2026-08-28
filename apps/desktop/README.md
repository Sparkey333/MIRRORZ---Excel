# MIRRORZ Sheets — desktop shell

The Electron application: a main process, a preload bridge and the React
renderer. The spreadsheet itself lives in `packages/*`; this workspace is the
window, the file dialogs, the menu and the crash recovery around it.

---

## Running it in development

Two terminals, because the renderer is served by Vite with hot reload while the
shell is a compiled bundle:

```sh
# terminal 1 - the renderer, on http://localhost:5273
npm run dev --workspace @mirrorz/desktop

# terminal 2 - build the main process and preload, then launch Electron
npm run dev:electron --workspace @mirrorz/desktop
```

`dev:electron` rebuilds `dist/main` and `dist/preload` first. Changes to the
renderer are hot-reloaded; changes to anything under `src/main` or `src/preload`
need that command run again, because the main process is not reloadable.

Point the shell at a different dev server with `MIRRORZ_DEV_SERVER`. A packaged
build ignores the variable entirely and always loads from disk, so it cannot be
talked into loading a renderer from the network.

Open a file at launch the way the operating system would:

```sh
npm run dev:electron --workspace @mirrorz/desktop -- /absolute/path/Budget.xlsx
```

### Tests and typecheck

```sh
npm test --workspace @mirrorz/desktop        # renderer (jsdom) and shell (node)
npm run test:main --workspace @mirrorz/desktop
npm run typecheck --workspace @mirrorz/desktop
```

The shell's tests never launch Electron. Everything with a decision in it -
path safety, IPC validation, the recent-files store, the autosave journal, the
menu template, argv parsing - lives in a module that imports nothing from
Electron at runtime, and is tested directly. What is left in `index.ts`,
`windows.ts` and `ipc.ts` is wiring.

---

## Building installers

```sh
npm run pack --workspace @mirrorz/desktop     # unpacked directory, fastest check
npm run dist --workspace @mirrorz/desktop     # installers for the host platform
npm run dist:win --workspace @mirrorz/desktop
npm run dist:mac --workspace @mirrorz/desktop
npm run dist:linux --workspace @mirrorz/desktop
```

Output lands in `apps/desktop/release/`.

| Platform | Targets | Notes |
| --- | --- | --- |
| Windows | NSIS installer, portable `.exe` | x64; the portable build needs no installation and no administrator |
| macOS | `.dmg` | x64 and arm64 |
| Linux | AppImage, `.deb` | x64 |

Each platform must be built on itself. A `.dmg` needs macOS tooling, NSIS needs
Windows, and a cross-built AppImage works right up until it does not. The
release workflow uses one runner per platform for exactly this reason.

### Unsigned builds work, on purpose

Nothing in `electron-builder.yml` references a certificate, and
`mac.identity: null` disables signing outright rather than letting
electron-builder search the keychain and fail on a machine with no key. Building
the application you just cloned must not require a paid identity.

What users see with an unsigned build:

- **Windows** — SmartScreen shows "Windows protected your PC" on first run.
  More info, then Run anyway. The warning fades as a signed binary accumulates
  reputation, and never fades for an unsigned one.
- **macOS** — Gatekeeper refuses to open the app: right-click, Open, then Open
  again; on recent versions, also System Settings > Privacy & Security > Open
  Anyway. An unsigned, un-notarized `.dmg` downloaded through a browser carries
  the quarantine attribute and is refused outright until that is done once.
- **Linux** — no signing infrastructure to speak of; the AppImage needs
  `chmod +x` and the `.deb` installs normally.

### What a signed release additionally needs

| Platform | Needed | How it is supplied |
| --- | --- | --- |
| Windows | An OV or EV code-signing certificate. OV is roughly $200–400/year; EV comes on a hardware token or in a cloud HSM and clears SmartScreen immediately. | `CSC_LINK` (base64 `.pfx` or a URL) and `CSC_KEY_PASSWORD` in the environment. For an EV token, use a signing service or a self-hosted runner - a USB token cannot be plugged into a GitHub runner. |
| macOS | An Apple Developer Program membership ($99/year), a **Developer ID Application** certificate, and notarization by Apple. | `CSC_LINK` + `CSC_KEY_PASSWORD`, then `APPLE_ID`, `APPLE_APP_SPECIFIC_PASSWORD` and `APPLE_TEAM_ID` for notarization. In `electron-builder.yml` set `mac.identity` to the certificate name (remove `identity: null`), `hardenedRuntime: true`, and `mac.notarize: true`. Notarization is required for Gatekeeper to accept the app without the right-click dance. |
| Linux | Nothing for the AppImage. A `.deb` in an apt repository would want a GPG key for the repository metadata rather than the package. | — |

The release workflow already sets `CSC_IDENTITY_AUTO_DISCOVERY=false`; remove
that line on the macOS job when real credentials exist, and add the secrets to
the repository. No other change to the workflow is needed.

### Icons

`build/` holds the installer artwork and is currently empty; electron-builder
falls back to the Electron icon. `build/README.md` lists the exact files and
sizes a release should add.

---

## How it is put together

```
src/main/        the main process, one module per concern
  index.ts       lifecycle only: ready, open-file, activate, quit
  windows.ts     one BrowserWindow per document, and the security posture
  ipc.ts         the IPC handlers
  files.ts       reading and writing the user's documents
  validate.ts    every check applied to anything crossing IPC
  paths.ts       path safety and the format allowlists
  recent.ts      the Recent Files list
  autosave.ts    the crash-recovery journal
  menu.ts        the native menu, built as plain data
  cli.ts         argv and the macOS open-file queue
  channels.ts    the channel names, shared with the preload
  store.ts       a JSON settings file, and the seam tests inject through
src/preload/     the contextBridge, and nothing else
src/renderer/    the React application
```

### One window per document

Not tabs. Two complaints about the single-window model are structural rather
than cosmetic: two workbooks with the same file name become indistinguishable
once they are tabs, and putting one workbook on each of two monitors is either
impossible or needs a second instance with its own settings. A window per
document is something the operating system's own window management already
knows how to arrange.

A second launch does not start a second instance - it forwards its argv to the
first through the single-instance lock, so there is one recent-files list and
one idea of what is open. Opening a file that is already open focuses that
window instead of opening a second, racing copy.

### Security

The application opens files from strangers, so the renderer is treated as
hostile:

- `nodeIntegration: false`, `contextIsolation: true`, `sandbox: true`,
  `webSecurity: true`, no `webview` tag, no experimental features. These flags
  are exported as `SECURE_WEB_PREFERENCES` and asserted in tests, so relaxing
  one to make something convenient work fails the suite.
- A strict CSP: `default-src 'none'`, `script-src 'self'`, no `unsafe-eval`, no
  remote origins. `'unsafe-inline'` is allowed for **styles only**, because
  per-cell formatting is inline style attributes and there is no nonce
  mechanism for those. The dev-server relaxation exists only when a dev server
  does.
- Navigation is refused to anything but our own origin. `window.open` is always
  denied; `http`, `https` and `mailto` links are handed to the operating
  system's browser or mail client, and every other scheme - `file:`,
  `javascript:`, whatever some other application registered - is dropped.
- All permission requests (camera, microphone, location, notifications) are
  denied. A spreadsheet does not need them, and denying by default means a
  future feature has to ask explicitly.
- The spellchecker is off, at the window and at the session. This is a network
  decision rather than a typography one: Chromium downloads its Hunspell
  dictionaries from Google's CDN, from the browser process, where no CSP
  applies. Left on the default it would be the one component of this
  application that talks to the internet, and it is asserted off in the tests.
- Nothing in the renderer can reach `fs` or `child_process`. Every file
  operation goes through the preload, and every argument is validated **in the
  main process**, because a preload shares a process with the page it serves.

### The preload surface

`window.mirrorz` is exactly seven functions:

| Function | Does |
| --- | --- |
| `openFile()` | Native open dialog; resolves `null` when cancelled |
| `openPath(path)` | Opens one file, as the Recent Files menu does |
| `saveFile(bytes, path)` | Overwrites a file the user already chose |
| `saveFileAs(bytes, suggestedName)` | Save dialog seeded with a sanitized name |
| `recentFiles()` | The recent list |
| `onOpenRequest(cb)` | Files the OS wants opened here; returns an unsubscribe |
| `getVersion()` | Application, Electron, Chrome and Node versions |

`window.mirrorzShell` is separate and deliberately small: `onCommand` for menu
items and shell events, `autosave(payload)` for the journal, and
`setDocumentState` so the window can show a modified indicator and prompt before
closing. It is a second namespace rather than three more functions on the first
so the file surface stays exactly the seven above, and nothing on it can read or
write a file the user did not choose.

`window.mirrorzHost` is a compatibility shim for the renderer's existing
`FileHost` interface, expressed entirely in terms of those seven functions.

### File associations

Double-clicking a `.xlsx`, `.xlsm`, `.xltx`, `.xltm`, `.xls`, `.csv`, `.tsv` or
`.ods` opens MIRRORZ. Three different mechanisms deliver that intent and all
three end in the same validation: `open-file` on macOS (queued, because it
arrives before the app is ready), `process.argv` on Windows and Linux, and the
forwarded argv of a second launch.

`.xls` and `.ods` are registered as **Viewer** rather than Editor, because the
format layer reads them and does not write them. Offering to save into a format
that would silently become something else is worse than not offering it.

### Crash recovery

The user's file is written when, and only when, they save it. What is persisted
on a timer is the command log, into a journal file **beside** the document
(`.Budget.xlsx.mirrorz-journal.json`), or in the application's recovery
directory when that folder is not writable or the document has never been saved.

On a clean close the journal is deleted. So a journal still on disk at startup
means the previous run did not shut down cleanly - no heuristics, no lock file
to go stale, nothing to clean up after a normal quit. What is found is
classified before anyone is asked about it:

- **recover** — offered.
- **source-missing** — the file was moved or deleted; offered anyway.
- **superseded** — the file on disk is newer than the journal, so the user
  already saved this work. Discarded silently; offering it would invite
  overwriting good work with old work.
- **stale** — older than 30 days. Discarded.

The autosave clock lives in the main process, so the cadence is one policy in
one place, and the write is rate-limited there too: a renderer in a tight loop
cannot use it to hammer the disk.

---

## Dependencies

| Package | Licence | Why |
| --- | --- | --- |
| `electron` | MIT | The shell — pinned to an exact version, because electron-builder cannot resolve a range and refuses to package |
| `electron-builder` | MIT | Installers |
| `vite` | MIT | Bundles the renderer, and the main and preload processes |
| `react`, `react-dom` | MIT | The renderer |

No dependency here is GPL, LGPL or AGPL.
