/**
 * The window's cage.
 *
 * These are the tests that should fail loudly if someone relaxes a flag to make
 * something convenient work. The threat model is a spreadsheet that arrived by
 * email containing a link, or a fragment of HTML pasted from a browser: neither
 * may be able to navigate the window anywhere, open a second Electron window,
 * or load a script from outside our own bundle.
 */

import { describe, expect, it } from 'vitest';
import {
  SECURE_WEB_PREFERENCES,
  contentSecurityPolicy,
  decideExternalOpen,
  isAllowedNavigation,
} from '../../src/main/security.js';

describe('webPreferences', () => {
  it('keeps Node out of the renderer', () => {
    expect(SECURE_WEB_PREFERENCES.nodeIntegration).toBe(false);
    expect(SECURE_WEB_PREFERENCES.nodeIntegrationInWorker).toBe(false);
    expect(SECURE_WEB_PREFERENCES.nodeIntegrationInSubFrames).toBe(false);
  });

  it('isolates the preload context from the page', () => {
    expect(SECURE_WEB_PREFERENCES.contextIsolation).toBe(true);
  });

  it('runs the renderer in the OS sandbox', () => {
    expect(SECURE_WEB_PREFERENCES.sandbox).toBe(true);
  });

  it('leaves the web security model intact', () => {
    expect(SECURE_WEB_PREFERENCES.webSecurity).toBe(true);
    expect(SECURE_WEB_PREFERENCES.allowRunningInsecureContent).toBe(false);
    expect(SECURE_WEB_PREFERENCES.experimentalFeatures).toBe(false);
    expect(SECURE_WEB_PREFERENCES.webviewTag).toBe(false);
  });

  /**
   * Not a nicety. Chromium's spellchecker downloads its Hunspell dictionaries
   * from Google's CDN, from the browser process, where no Content-Security-
   * Policy applies - so leaving the default on is a silent outbound request in
   * an application that promises it makes none. This is the assertion that
   * fails if somebody turns it back on.
   */
  it('makes no outbound request for spellcheck dictionaries', () => {
    expect(SECURE_WEB_PREFERENCES.spellcheck).toBe(false);
  });
});

describe('contentSecurityPolicy', () => {
  const production = contentSecurityPolicy();

  it('denies everything that is not explicitly allowed', () => {
    expect(production).toContain("default-src 'none'");
  });

  it('never allows inline or evaluated script in production', () => {
    expect(production).toContain("script-src 'self'");
    expect(production).not.toContain("script-src 'self' 'unsafe-inline'");
    expect(production).not.toContain("'unsafe-eval'");
  });

  it('allows inline styles, which per-cell formatting requires', () => {
    expect(production).toContain("style-src 'self' 'unsafe-inline'");
  });

  it('allows data and blob images, for icons and generated charts', () => {
    expect(production).toContain('img-src');
    expect(production).toContain('data:');
  });

  it('blocks plugins, frames and framing', () => {
    expect(production).toContain("object-src 'none'");
    expect(production).toContain("frame-src 'none'");
    expect(production).toContain("frame-ancestors 'none'");
  });

  it('blocks form submission and base-tag hijacking', () => {
    expect(production).toContain("form-action 'none'");
    expect(production).toContain("base-uri 'none'");
  });

  it('permits no network connection beyond our own origin', () => {
    expect(production).toContain("connect-src 'self'");
    expect(production).not.toContain('https://');
  });

  it('relaxes only for the dev server, and only in development', () => {
    const dev = contentSecurityPolicy({ devServer: 'http://localhost:5273' });
    expect(dev).toContain('http://localhost:5273');
    expect(dev).toContain('ws://localhost:5273');
    expect(dev).toContain("'unsafe-eval'");
    expect(production).not.toContain('localhost');
  });
});

describe('isAllowedNavigation', () => {
  it('allows our own bundle', () => {
    expect(isAllowedNavigation('file:///opt/mirrorz/dist/renderer/index.html')).toBe(true);
  });

  it('blocks the open web', () => {
    expect(isAllowedNavigation('https://example.test/')).toBe(false);
    expect(isAllowedNavigation('http://example.test/')).toBe(false);
  });

  it('blocks javascript: and data: navigations', () => {
    expect(isAllowedNavigation('javascript:alert(1)')).toBe(false);
    expect(isAllowedNavigation('data:text/html,<script>alert(1)</script>')).toBe(false);
  });

  it('blocks a string that is not a URL at all', () => {
    expect(isAllowedNavigation('')).toBe(false);
    expect(isAllowedNavigation('/etc/passwd')).toBe(false);
  });

  it('allows the dev server when one is configured', () => {
    const policy = { devServer: 'http://localhost:5273' };
    expect(isAllowedNavigation('http://localhost:5273/index.html', policy)).toBe(true);
  });

  it('fences file navigation to the directory the bundle was loaded from', () => {
    // A local page inherits this window's preload, so "any file: URL" is not a
    // policy - it is a route from a linked document to the IPC surface.
    const policy = { rendererFile: '/opt/mirrorz/dist/renderer/index.html' };
    expect(isAllowedNavigation('file:///opt/mirrorz/dist/renderer/index.html', policy)).toBe(true);
    expect(isAllowedNavigation('file:///opt/mirrorz/dist/renderer/help.html', policy)).toBe(true);
    expect(isAllowedNavigation('file:///etc/passwd', policy)).toBe(false);
    expect(isAllowedNavigation('file:///home/someone/Downloads/evil.html', policy)).toBe(false);
    expect(isAllowedNavigation('file:///opt/mirrorz/dist/main/index.cjs', policy)).toBe(false);
  });

  it('does not accept an escaped traversal back out of the bundle', () => {
    const policy = { rendererFile: '/opt/mirrorz/dist/renderer/index.html' };
    expect(
      isAllowedNavigation('file:///opt/mirrorz/dist/renderer/%2e%2e/main/index.cjs', policy),
    ).toBe(false);
    expect(isAllowedNavigation('file:///opt/mirrorz/dist/renderer/../../etc/passwd', policy)).toBe(
      false,
    );
  });

  it('does not allow a lookalike of the dev server host', () => {
    const policy = { devServer: 'http://localhost:5273' };
    expect(isAllowedNavigation('http://localhost:5274/', policy)).toBe(false);
    expect(isAllowedNavigation('http://localhost.example.test:5273/', policy)).toBe(false);
    expect(isAllowedNavigation('https://localhost:5273/', policy)).toBe(false);
  });
});

describe('decideExternalOpen', () => {
  it('sends a web link to the user browser', () => {
    expect(decideExternalOpen('https://example.test/help')).toEqual({
      action: 'external',
      url: 'https://example.test/help',
    });
  });

  it('sends a mailto link to the user mail client', () => {
    expect(decideExternalOpen('mailto:someone@example.test').action).toBe('external');
  });

  it('refuses a scheme registered by some other application', () => {
    expect(decideExternalOpen('ms-msdt:/id PCWDiagnostic').action).toBe('deny');
    expect(decideExternalOpen('smb://server/share').action).toBe('deny');
    expect(decideExternalOpen('vscode://file/etc/passwd').action).toBe('deny');
  });

  it('refuses file: so a hyperlink cannot launch a local binary', () => {
    expect(decideExternalOpen('file:///bin/sh').action).toBe('deny');
  });

  it('refuses javascript: and data:', () => {
    expect(decideExternalOpen('javascript:alert(1)').action).toBe('deny');
    expect(decideExternalOpen('data:text/html,x').action).toBe('deny');
  });

  it('refuses a string that is not a URL', () => {
    expect(decideExternalOpen('not a url').action).toBe('deny');
  });
});
