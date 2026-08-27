/**
 * The renderer's cage: navigation policy and Content-Security-Policy.
 *
 * A spreadsheet is a document format with hyperlinks in it, and the files come
 * from strangers. The threat is concrete: a cell containing a link, or an HTML
 * fragment pasted from a browser, persuading the window to navigate somewhere
 * and thereby giving remote content the renderer's privileges. The answer is
 * that this application has exactly one origin it will ever display, and every
 * other URL is either handed to the operating system's browser or dropped.
 *
 * Both policies are plain functions over strings so the decision table can be
 * tested directly, rather than only being observable by driving a real window.
 */

/** What the renderer is allowed to load, as a CSP header value. */
export interface CspOptions {
  /** In development the Vite dev server also needs its websocket for HMR. */
  devServer?: string | undefined;
}

/**
 * `'unsafe-inline'` for styles only, and never for scripts.
 *
 * The renderer sets inline `style` attributes for per-cell formatting - a
 * spreadsheet's whole job - and there is no nonce mechanism for style
 * attributes. Scripts get no such exemption: `script-src 'self'` is what makes
 * an injected `<script>` from a pasted document inert.
 */
export function contentSecurityPolicy(opts: CspOptions = {}): string {
  const dev = opts.devServer;
  const connect = ["'self'"];
  const script = ["'self'"];
  if (dev) {
    connect.push(dev, dev.replace(/^http/, 'ws'));
    // The dev server serves modules over http and evaluates them; production
    // never carries this relaxation because production never has a dev server.
    script.push(dev, "'unsafe-eval'");
  }
  return [
    "default-src 'none'",
    `script-src ${script.join(' ')}`,
    `style-src 'self' 'unsafe-inline'${dev ? ` ${dev}` : ''}`,
    `img-src 'self' data: blob:`,
    `font-src 'self' data:`,
    `connect-src ${connect.join(' ')}`,
    "object-src 'none'",
    "frame-src 'none'",
    "worker-src 'self' blob:",
    "base-uri 'none'",
    "form-action 'none'",
    "frame-ancestors 'none'",
  ].join('; ');
}

export interface NavigationPolicy {
  /** The origin the packaged renderer is served from; `file://` has none. */
  devServer?: string | undefined;
}

/**
 * Origins the window may navigate to. In production that is the `file:` URL of
 * our own bundle; in development it is additionally the dev server.
 */
export function isAllowedNavigation(target: string, policy: NavigationPolicy = {}): boolean {
  let url: URL;
  try {
    url = new URL(target);
  } catch {
    return false;
  }
  if (url.protocol === 'file:') return true;
  if (policy.devServer) {
    try {
      const dev = new URL(policy.devServer);
      if (url.protocol === dev.protocol && url.host === dev.host) return true;
    } catch {
      return false;
    }
  }
  return false;
}

export type ExternalDecision =
  | { action: 'external'; url: string }
  | { action: 'deny'; reason: string };

/**
 * What to do when the page asks for a new window or clicks a link out.
 *
 * Never open a new Electron window: an attacker-controlled page inside our
 * process model is the thing we are avoiding. http and https go to the user's
 * browser, which has its own sandbox and its own idea of trust. Everything else
 * - `file:`, `javascript:`, custom schemes registered by other applications,
 * which is a well-worn path to local code execution - is dropped.
 */
export function decideExternalOpen(target: string): ExternalDecision {
  let url: URL;
  try {
    url = new URL(target);
  } catch {
    return { action: 'deny', reason: 'not a URL' };
  }
  if (url.protocol === 'http:' || url.protocol === 'https:') {
    return { action: 'external', url: url.toString() };
  }
  if (url.protocol === 'mailto:') return { action: 'external', url: url.toString() };
  return { action: 'deny', reason: `scheme not permitted: ${url.protocol}` };
}

/**
 * The BrowserWindow webPreferences this application will accept.
 *
 * Exported as data, and asserted in tests, because these four flags are the
 * whole security posture and a future edit that flips one of them should fail a
 * test rather than ship.
 */
export const SECURE_WEB_PREFERENCES = {
  nodeIntegration: false,
  nodeIntegrationInWorker: false,
  nodeIntegrationInSubFrames: false,
  contextIsolation: true,
  sandbox: true,
  webSecurity: true,
  allowRunningInsecureContent: false,
  experimentalFeatures: false,
  webviewTag: false,
  spellcheck: true,
} as const;
