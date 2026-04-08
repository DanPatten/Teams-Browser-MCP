import fs from 'node:fs';
import path from 'node:path';
import { chromium } from 'playwright';
import { log } from './logger.js';
import { Paths } from './paths.js';
import { AuthState } from './authState.js';

const TEAMS_URL = 'https://teams.microsoft.com/v2/';

// Authoritative logged-in DOM markers. If any one of these is visible,
// Teams is on the app shell rather than the sign-in flow. Both the
// initial auth check and the interactive poll loop use this list.
const LOGGED_IN_SELECTORS = [
  "[data-tid='app-bar-2a']",
  'div#app-bar',
  "div[role='navigation'][aria-label*='App']",
];

// If the URL contains any of these, we're on a sign-in page.
const LOGIN_URL_HOSTS = ['login.microsoftonline.com', 'login.live.com'];

// Visible selectors that prove we're sitting on a sign-in form.
const LOGIN_FORM_SELECTORS = [
  "input[name='loginfmt']",
  "input[type='email']",
];

// Budget for the quick "are we still authed?" check on every getPage().
// Generous because a false negative here triggers a full headed re-sign-in,
// while a false positive just causes the calling tool to fail with a clear
// selector error. Bias toward declaring authed.
const QUICK_AUTH_CHECK_MS = 3000;

// One-time warm-up after a fresh page is established. The auth markers
// (LOGGED_IN_SELECTORS) appear early in Teams' boot — before the search
// box, chat list, etc. are usable. We wait for one of these "actually
// ready" markers before handing the page to the first tool, then never
// wait again for the lifetime of the session.
const UI_READY_SELECTORS = [
  "[data-tid='AUTOSUGGEST_INPUT']", // Global search box in the persistent header.
];
const UI_READY_TIMEOUT_MS = 15_000;

// Budget for the initial "logged in vs login page" decision after the
// very first navigate.
const INITIAL_DECIDE_MS = 20_000;

// Interactive sign-in poll loop.
const SIGN_IN_TIMEOUT_MS = 5 * 60 * 1000;
const SIGN_IN_POLL_INTERVAL_MS = 2000;

const ABORT_FILE = path.join(Paths.root, 'abort-auth');

// Hostname suffixes that are allowed to load in the Playwright context.
// Everything else (ads, tracking, 3rd-party CDNs, generic web) is
// aborted at the network layer so this Chrome window can't be used as a
// general-purpose browser — the point of the lock is that it exists
// solely to talk to Teams.
//
// Keep this list generous enough that the Microsoft sign-in flow still
// works (which pulls from aadcdn/msauth/live.com) and Teams v2 features
// don't break (SharePoint attachments, Azure storage for avatars, the
// various MS CDNs). If Teams ever starts pulling from a new host, add
// the suffix here.
const ALLOWED_HOST_SUFFIXES = [
  'microsoft.com',
  'microsoftonline.com',
  'office.com',
  'office.net',
  'sharepoint.com',
  'skype.com',
  'skypeforbusiness.com',
  'live.com',
  'live.net',
  'cloud.microsoft',
  'msauth.net',
  'msauthimages.net',
  'msftauth.net',
  'msftauthimages.net',
  'azureedge.net',
  'windows.net',
  'trafficmanager.net',
  'msecnd.net',
  'onmicrosoft.com',
  'gfx.ms',
  'msocdn.com',
  'msedge.net',
  'bing.com', // Teams search sometimes calls out to Bing suggest.
];

function isAllowedUrl(url) {
  try {
    const u = new URL(url);
    // Never block non-http (data:, blob:, about:, chrome:, ws:, etc.)
    if (u.protocol !== 'http:' && u.protocol !== 'https:') return true;
    const host = u.hostname.toLowerCase();
    return ALLOWED_HOST_SUFFIXES.some(
      (suf) => host === suf || host.endsWith('.' + suf),
    );
  } catch {
    // If URL can't be parsed, err on the side of letting it through.
    return true;
  }
}

/**
 * Long-lived Playwright session. Owns the entire auth lifecycle: we
 * launch chromium once, re-use the same page across tool calls, and
 * every getPage() verifies the session is still authed before handing
 * the page back.
 *
 * All cookie scraping / native-browser launching is gone. The user
 * signs in directly inside this Playwright Chrome window; the
 * context's storageState is persisted to state.json so the next
 * process start replays it.
 */
class TeamsSession {
  constructor() {
    this._browser = null;
    this._context = null;
    this._page = null;
    this._headed = false;
    this._lock = Promise.resolve();
  }

  async getPage() {
    // Simple mutex: chain each ensureSession onto the previous one so
    // callers serialize through the init path.
    const next = this._lock.then(() => this._ensureSession());
    this._lock = next.catch(() => {}); // don't poison the lock on failure
    return next;
  }

  async forceReauthenticate() {
    const next = this._lock.then(async () => {
      await this._disposeSession();
      AuthState.delete();
      return this._ensureSession();
    });
    this._lock = next.catch(() => {});
    return next;
  }

  async _disposeSession() {
    if (this._context) {
      try { await this._context.close(); } catch { /* ignore */ }
    }
    if (this._browser) {
      try { await this._browser.close(); } catch { /* ignore */ }
    }
    this._context = null;
    this._page = null;
    this._browser = null;
    this._headed = false;
  }

  async _ensureSession() {
    // Fast path: browser up, page alive, and still authed.
    if (this._browser && this._page && !this._page.isClosed()) {
      const authed = await this._isAuthed(QUICK_AUTH_CHECK_MS);
      if (authed) {
        log.info('session', 'reuse-existing-page', { authed: true, headed: this._headed });
        await this._persistStorageState();
        return this._page;
      }
      log.warn('session', 'existing-page-not-authed; switching to headed sign-in', { headed: this._headed });
      await this._doInteractiveSignIn();
      return this._page;
    }

    // Cold start path: no browser yet.
    if (AuthState.exists()) {
      // Try the happy path: launch headless and replay cached state.
      log.info('session', 'cold-start:headless-with-cached-state');
      await this._launchBrowser({ headed: false });

      await log.span('browser', 'goto', async () => {
        await this._page.goto(TEAMS_URL, { waitUntil: 'domcontentloaded' });
      }, { url: TEAMS_URL });

      const authed = await this._waitForAuthOrLoginPage(INITIAL_DECIDE_MS);
      if (authed) {
        log.info('session', 'initial-state:already-authed');
        await this._persistStorageState();
        return this._page;
      }
      log.warn('session', 'cached-state-stale; switching to headed sign-in');
      // Fall through to interactive sign-in.
    } else {
      log.info('session', 'cold-start:no-cached-state; headed sign-in needed');
    }

    await this._doInteractiveSignIn();
    return this._page;
  }

  /**
   * Tear down whatever session we have, launch a *headed* Chrome so the
   * user can sign in, run the interactive sign-in loop, persist the
   * resulting storage state, and then relaunch headless so subsequent
   * tool calls don't show a window.
   */
  async _doInteractiveSignIn() {
    await this._disposeSession();

    log.info('session', 'launch:headed');
    await this._launchBrowser({ headed: true });
    await log.span('browser', 'goto', async () => {
      await this._page.goto(TEAMS_URL, { waitUntil: 'domcontentloaded' });
    }, { url: TEAMS_URL });

    await this._runSignInLoop();
    await this._persistStorageState();

    log.info('session', 'relaunch:headless-after-auth');
    await this._disposeSession();
    await this._launchBrowser({ headed: false });
    await log.span('browser', 'goto', async () => {
      await this._page.goto(TEAMS_URL, { waitUntil: 'domcontentloaded' });
    }, { url: TEAMS_URL });

    const authed = await this._waitForAuthOrLoginPage(INITIAL_DECIDE_MS);
    if (!authed) {
      log.error('session', 'headless-replay-failed', null, {});
      throw new Error(
        'Sign-in completed but the headless replay of the saved storage ' +
        'state did not land on an authenticated Teams session. Try calling ' +
        '`authenticate` again.',
      );
    }
    await this._persistStorageState();
  }

  async _launchBrowser({ headed = false } = {}) {
    // Clear any stale abort sentinel so we don't immediately bail on
    // the first sign-in loop.
    try {
      if (fs.existsSync(ABORT_FILE)) fs.unlinkSync(ABORT_FILE);
    } catch { /* ignore */ }

    await log.span('browser', headed ? 'launch:headed' : 'launch:headless', async () => {
      this._browser = await chromium.launch({
        channel: 'chrome',
        headless: !headed,
      });
      this._headed = headed;
    });

    await log.span('browser', 'newContext', async () => {
      const opts = { viewport: { width: 1280, height: 900 } };
      if (AuthState.exists()) {
        opts.storageState = Paths.stateFile;
        log.info('browser', 'newContext:using-stored-state', { path: Paths.stateFile });
      } else {
        log.info('browser', 'newContext:fresh');
      }
      this._context = await this._browser.newContext(opts);
    });

    // Install a network-level allowlist: only Microsoft-owned hosts
    // (plus auth/CDN infra Teams depends on) can load. Everything else
    // is aborted so this Chrome window can't be used as a general
    // browser and third-party tracking/ads/etc. are cut off. Blocked
    // host aggregation is logged every 20 blocks so the log doesn't
    // drown in spam.
    const blockedCounts = new Map();
    let blockedSinceLog = 0;
    await log.span('browser', 'install-url-allowlist', async () => {
      await this._context.route('**/*', (route) => {
        const url = route.request().url();
        if (isAllowedUrl(url)) {
          return route.continue();
        }
        try {
          const host = new URL(url).hostname;
          blockedCounts.set(host, (blockedCounts.get(host) || 0) + 1);
          blockedSinceLog++;
          if (blockedSinceLog >= 20) {
            log.info('browser', 'url-allowlist:blocked-sample', {
              total: blockedSinceLog,
              topHosts: Array.from(blockedCounts.entries())
                .sort((a, b) => b[1] - a[1])
                .slice(0, 8)
                .map(([h, n]) => `${h}=${n}`),
            });
            blockedSinceLog = 0;
          }
        } catch { /* ignore URL parse errors */ }
        return route.abort();
      });
      log.info('browser', 'url-allowlist:installed', {
        allowedSuffixCount: ALLOWED_HOST_SUFFIXES.length,
      });
    });

    this._page = await log.span('browser', 'newPage', async () => this._context.newPage());
  }

  /**
   * Quick yes/no authed check. Used on every getPage() call and inside
   * the sign-in poll loop. Races all logged-in selectors against the
   * full budget so the selector that actually becomes visible first
   * wins — dividing the budget across selectors leads to false
   * negatives when Teams' DOM happens to be a few hundred ms slower
   * than expected.
   */
  async _isAuthed(totalBudgetMs) {
    return log.span('session', 'isAuthed', async () => {
      if (!this._page || this._page.isClosed()) {
        log.info('session', 'isAuthed:result', { authed: false, reason: 'no-page', headed: this._headed });
        return false;
      }
      let authed = false;
      try {
        await Promise.any(
          LOGGED_IN_SELECTORS.map((sel) =>
            this._page.waitForSelector(sel, { state: 'visible', timeout: totalBudgetMs }),
          ),
        );
        authed = true;
      } catch {
        authed = false;
      }
      log.info('session', 'isAuthed:result', { authed, headed: this._headed, budgetMs: totalBudgetMs });
      return authed;
    }, { budgetMs: totalBudgetMs, headed: this._headed });
  }

  /**
   * After the initial navigate, loop until we decide one of:
   *   - logged in → return true
   *   - on a login page → return false
   * Bounded by totalTimeoutMs.
   */
  async _waitForAuthOrLoginPage(totalTimeoutMs) {
    const deadline = Date.now() + totalTimeoutMs;
    while (Date.now() < deadline) {
      if (await this._isAuthed(800)) return true;
      if (this._looksLikeLoginPage()) return false;
      if (await this._hasVisibleLoginForm()) return false;
      await sleep(500);
    }
    // Timed out without a clear decision. Treat as "needs sign-in" —
    // the poll loop is the safer bet than erroring.
    log.warn('session', 'wait-auth-or-login:timeout; assuming needs-sign-in');
    return false;
  }

  _looksLikeLoginPage() {
    try {
      const url = this._page.url() || '';
      return LOGIN_URL_HOSTS.some((h) => url.includes(h));
    } catch {
      return false;
    }
  }

  async _hasVisibleLoginForm() {
    for (const sel of LOGIN_FORM_SELECTORS) {
      try {
        await this._page.waitForSelector(sel, { state: 'visible', timeout: 200 });
        return true;
      } catch { /* try next */ }
    }
    return false;
  }

  /**
   * Interactive sign-in loop. Focuses the Playwright window, prints a
   * stderr banner, and polls the logged-in selectors every 2s. Supports
   * the abort-file escape hatch. Throws on timeout.
   */
  async _runSignInLoop() {
    try { await this._page.bringToFront(); } catch { /* ignore */ }
    printUserBanner();
    log.info('auth', 'sign-in-loop:begin', {
      timeoutMs: SIGN_IN_TIMEOUT_MS,
      abortFile: ABORT_FILE,
    });

    const deadline = Date.now() + SIGN_IN_TIMEOUT_MS;
    let pollCount = 0;

    while (Date.now() < deadline) {
      pollCount++;
      log.info('auth', 'poll:tick', {
        n: pollCount,
        remainingMs: deadline - Date.now(),
      });

      if (checkAbortFile()) {
        log.warn('auth', 'aborted-via-file', { pollCount });
        throw new Error(
          `Sign-in aborted by user (abort file at ${ABORT_FILE}). ` +
          'Call the `authenticate` tool again to retry.',
        );
      }

      // Refresh focus — a single bringToFront doesn't always stick on
      // Windows if other windows have been interacted with since.
      try { await this._page.bringToFront(); } catch { /* ignore */ }

      if (await this._isAuthed(800)) {
        log.info('auth', 'sign-in-loop:captured', { pollCount });
        return;
      }

      await sleep(SIGN_IN_POLL_INTERVAL_MS);
    }

    log.error('auth', 'sign-in-loop:timeout', null, { pollCount });
    throw new Error(
      'Timed out waiting for Teams sign-in. Complete the login in the ' +
      'Playwright Chrome window, then call the `authenticate` tool again.',
    );
  }

  async _persistStorageState() {
    try {
      await log.span(
        'session',
        'persist-storage-state',
        () => this._context.storageState({ path: Paths.stateFile }),
        { path: Paths.stateFile },
      );
    } catch (err) {
      log.warn('session', 'persist:failed', { err: err.message });
    }
  }
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function checkAbortFile() {
  if (fs.existsSync(ABORT_FILE)) {
    try { fs.unlinkSync(ABORT_FILE); } catch { /* ignore */ }
    return true;
  }
  return false;
}

function printUserBanner() {
  const lines = [
    '',
    '==================================================================',
    '  Teams Browser MCP — sign-in required',
    '==================================================================',
    '  A Playwright-controlled Chrome window has been opened and',
    '  brought to the front. Sign in to Microsoft Teams inside THAT',
    '  window. This process will auto-detect your session (polling',
    `  every ${SIGN_IN_POLL_INTERVAL_MS / 1000}s) and continue automatically.`,
    '',
    `  Timeout: ${Math.round(SIGN_IN_TIMEOUT_MS / 60000)} minutes`,
    '',
    '  To abort without waiting for the timeout, create this file:',
    `    ${ABORT_FILE}`,
    '==================================================================',
    '',
  ];
  for (const line of lines) {
    process.stderr.write(line + '\n');
  }
}

export const session = new TeamsSession();
