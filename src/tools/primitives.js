import { session } from '../teamsSession.js';
import { log } from '../logger.js';

// Primitive browser tools that operate on the long-lived authed Teams
// page owned by `session`. These intentionally mirror Playwright MCP's
// surface (navigate / query / click / type / press / wait / evaluate)
// but are scoped to the single Teams tab and piggyback on the auth +
// URL-allowlist machinery in teamsSession.js.
//
// All selectors are CSS. There is no accessibility-tree / ref
// abstraction here — calling LLMs are expected to pass CSS selectors
// directly. See TEAMS_GUIDE.md (exposed via the `teams_guide` tool)
// for navigation hints.

const TEAMS_HOST_RE = /(^|\.)teams\.microsoft\.com$/i;

// MCP servers don't always enforce inputSchema before the handler runs,
// so callers can land here with required string args missing or of the
// wrong type. Without this guard a missing `selector` gets coerced to
// the literal string "undefined" by document.querySelectorAll and
// silently returns 0 matches — turning a misuse into a fake-success.
function requireString(toolName, fieldName, value) {
  if (typeof value !== 'string' || value.trim() === '') {
    throw new Error(
      `${toolName}: '${fieldName}' is required and must be a non-empty string ` +
      `(got ${value === undefined ? 'undefined' : JSON.stringify(value)})`,
    );
  }
}

function assertTeamsUrl(url) {
  let u;
  try {
    u = new URL(url);
  } catch {
    throw new Error(`Invalid URL: ${url}`);
  }
  if (u.protocol !== 'https:' || !TEAMS_HOST_RE.test(u.hostname)) {
    throw new Error(
      `teams_navigate only allows https://teams.microsoft.com/* (got ${url})`,
    );
  }
  return u.toString();
}

// In-page query payload. Receives the selector + limit, returns a
// trimmed serialization of each match so the calling LLM can decide
// what to do next without an evaluate round-trip.
function queryInPage({ selector, limit }) {
  let nodes;
  try {
    nodes = document.querySelectorAll(selector);
  } catch (err) {
    return { error: 'invalid-selector', message: err.message };
  }
  const out = [];
  const cap = Math.min(nodes.length, limit);
  for (let i = 0; i < cap; i++) {
    const el = nodes[i];
    const attrs = {};
    for (const a of el.attributes) {
      // Skip noisy/long attrs.
      if (a.name === 'style' || a.name === 'class') continue;
      if (a.value && a.value.length <= 200) attrs[a.name] = a.value;
    }
    if (el.getAttribute('class')) {
      attrs.class = el.getAttribute('class').slice(0, 200);
    }
    const rect = el.getBoundingClientRect();
    out.push({
      tag: el.tagName,
      text: (el.innerText || el.textContent || '').trim().slice(0, 240),
      attrs,
      visible: rect.width > 0 && rect.height > 0,
    });
  }
  return { total: nodes.length, returned: out.length, matches: out };
}

// React-friendly value setter so controlled inputs notice the change.
function typeInPage({ selector, text, submit }) {
  const el = document.querySelector(selector);
  if (!el) return { ok: false, error: 'no-match', selector };
  if (typeof el.focus === 'function') el.focus();
  const proto =
    el.tagName === 'TEXTAREA'
      ? HTMLTextAreaElement.prototype
      : HTMLInputElement.prototype;
  const setter = Object.getOwnPropertyDescriptor(proto, 'value')?.set;
  if (setter) {
    setter.call(el, text);
  } else {
    el.value = text;
  }
  el.dispatchEvent(new Event('input', { bubbles: true }));
  el.dispatchEvent(new Event('change', { bubbles: true }));
  return { ok: true, submit: !!submit };
}

export const teamsNavigateTool = {
  name: 'teams_navigate',
  description:
    'Navigate the persistent authed Teams page to a teams.microsoft.com URL. ' +
    'Only https://teams.microsoft.com/* is allowed. Returns the final URL ' +
    'and document title after navigation.',
  inputSchema: {
    type: 'object',
    properties: {
      url: { type: 'string', description: 'Absolute https://teams.microsoft.com/* URL' },
      waitUntil: {
        type: 'string',
        enum: ['load', 'domcontentloaded', 'networkidle'],
        default: 'domcontentloaded',
      },
    },
    required: ['url'],
    additionalProperties: false,
  },
  async handler({ url, waitUntil = 'domcontentloaded' }) {
    return log.span('tool.teams_navigate', 'tool', async () => {
      requireString('teams_navigate', 'url', url);
      const target = assertTeamsUrl(url);
      const page = await session.getPage();
      await page.goto(target, { waitUntil });
      const finalUrl = page.url();
      const title = await page.title();
      log.info('tool.teams_navigate', 'done', { finalUrl, title });
      return {
        content: [
          { type: 'text', text: JSON.stringify({ url: finalUrl, title }) },
        ],
      };
    });
  },
};

export const teamsQueryTool = {
  name: 'teams_query',
  description:
    'DOM query — NOT a content/message search. Runs ' +
    'document.querySelectorAll(selector) on the Teams page and returns a ' +
    'trimmed serialization (tag, text, attributes, visibility) of each ' +
    'match. Takes a CSS `selector` (e.g. "[data-tid=\'chat-pane-message\']"), ' +
    'NOT a free-text query string. To search Teams content, use ' +
    'teams_type to type into the search box and then teams_query to read ' +
    'the result list. Default limit is 20 matches.',
  inputSchema: {
    type: 'object',
    properties: {
      selector: { type: 'string' },
      limit: { type: 'integer', minimum: 1, maximum: 200, default: 20 },
    },
    required: ['selector'],
    additionalProperties: false,
  },
  async handler({ selector, limit = 20 }) {
    return log.span('tool.teams_query', 'tool', async () => {
      requireString('teams_query', 'selector', selector);
      const page = await session.getPage();
      const result = await page.evaluate(queryInPage, { selector, limit });
      log.info('tool.teams_query', 'result', {
        selector,
        total: result?.total,
        returned: result?.returned,
        error: result?.error || null,
      });
      return { content: [{ type: 'text', text: JSON.stringify(result) }] };
    }, { selector });
  },
};

export const teamsClickTool = {
  name: 'teams_click',
  description:
    'Click an element by CSS selector on the Teams page using Playwright ' +
    '(real mouse event, not in-page el.click). If multiple match, use nth ' +
    '(0-indexed) to pick one. Waits for the element to be visible up to ' +
    'timeoutMs (default 5000).',
  inputSchema: {
    type: 'object',
    properties: {
      selector: { type: 'string' },
      nth: { type: 'integer', minimum: 0, default: 0 },
      timeoutMs: { type: 'integer', minimum: 100, maximum: 60_000, default: 5000 },
    },
    required: ['selector'],
    additionalProperties: false,
  },
  async handler({ selector, nth = 0, timeoutMs = 5000 }) {
    return log.span('tool.teams_click', 'tool', async () => {
      requireString('teams_click', 'selector', selector);
      const page = await session.getPage();
      const locator = page.locator(selector).nth(nth);
      await locator.waitFor({ state: 'visible', timeout: timeoutMs });
      await locator.click({ timeout: timeoutMs });
      log.info('tool.teams_click', 'clicked', { selector, nth });
      return {
        content: [{ type: 'text', text: JSON.stringify({ ok: true, selector, nth }) }],
      };
    }, { selector, nth });
  },
};

export const teamsTypeTool = {
  name: 'teams_type',
  description:
    'Type text into an input/textarea/contenteditable on the Teams page. ' +
    'Uses a React-friendly value setter so controlled components register ' +
    'the change. Set submit=true to press Enter after typing.',
  inputSchema: {
    type: 'object',
    properties: {
      selector: { type: 'string' },
      text: { type: 'string' },
      submit: { type: 'boolean', default: false },
      clear: {
        type: 'boolean',
        default: true,
        description: 'Clear the field before typing (default true)',
      },
    },
    required: ['selector', 'text'],
    additionalProperties: false,
  },
  async handler({ selector, text, submit = false, clear = true }) {
    return log.span('tool.teams_type', 'tool', async () => {
      requireString('teams_type', 'selector', selector);
      requireString('teams_type', 'text', text);
      const page = await session.getPage();
      const locator = page.locator(selector).first();
      await locator.waitFor({ state: 'visible', timeout: 5000 });

      // Try Playwright's fill first (best for plain inputs / textareas).
      // Fall back to the React-friendly in-page setter for controlled
      // components and contenteditables that fill can't handle.
      let usedFallback = false;
      try {
        if (clear) await locator.fill('');
        await locator.fill(text);
      } catch (err) {
        log.info('tool.teams_type', 'fill-failed-fallback', { msg: err.message });
        usedFallback = true;
        const res = await page.evaluate(typeInPage, { selector, text, submit });
        if (!res.ok) {
          throw new Error(`teams_type fallback failed: ${res.error}`);
        }
      }

      if (submit) {
        await page.keyboard.press('Enter');
      }
      log.info('tool.teams_type', 'done', { selector, submit, usedFallback });
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify({ ok: true, selector, submit, usedFallback }),
          },
        ],
      };
    }, { selector });
  },
};

export const teamsPressKeyTool = {
  name: 'teams_press_key',
  description:
    'Press a key on the Teams page (e.g. "Enter", "Escape", "ArrowDown", ' +
    '"Control+K"). Uses Playwright keyboard.press syntax.',
  inputSchema: {
    type: 'object',
    properties: { key: { type: 'string' } },
    required: ['key'],
    additionalProperties: false,
  },
  async handler({ key }) {
    return log.span('tool.teams_press_key', 'tool', async () => {
      requireString('teams_press_key', 'key', key);
      const page = await session.getPage();
      await page.keyboard.press(key);
      return {
        content: [{ type: 'text', text: JSON.stringify({ ok: true, key }) }],
      };
    }, { key });
  },
};

export const teamsWaitForTool = {
  name: 'teams_wait_for',
  description:
    'Wait for an element matching the CSS selector to reach the given ' +
    'state ("visible" by default). Returns when the wait succeeds, or ' +
    'errors on timeout.',
  inputSchema: {
    type: 'object',
    properties: {
      selector: { type: 'string' },
      state: {
        type: 'string',
        enum: ['attached', 'detached', 'visible', 'hidden'],
        default: 'visible',
      },
      timeoutMs: { type: 'integer', minimum: 100, maximum: 60_000, default: 10_000 },
    },
    required: ['selector'],
    additionalProperties: false,
  },
  async handler({ selector, state = 'visible', timeoutMs = 10_000 }) {
    return log.span('tool.teams_wait_for', 'tool', async () => {
      requireString('teams_wait_for', 'selector', selector);
      const page = await session.getPage();
      await page.waitForSelector(selector, { state, timeout: timeoutMs });
      return {
        content: [
          { type: 'text', text: JSON.stringify({ ok: true, selector, state }) },
        ],
      };
    }, { selector, state });
  },
};

export const teamsEvaluateTool = {
  name: 'teams_evaluate',
  description:
    'Escape hatch: run a JavaScript expression in the Teams page and ' +
    'return its JSON-serialized result. The expression is evaluated as ' +
    '`(() => (<expression>))()`, so it can be a single expression or an ' +
    'IIFE-style block. Use sparingly — prefer teams_query / teams_click / ' +
    'teams_type when possible.',
  inputSchema: {
    type: 'object',
    properties: { expression: { type: 'string' } },
    required: ['expression'],
    additionalProperties: false,
  },
  async handler({ expression }) {
    return log.span('tool.teams_evaluate', 'tool', async () => {
      requireString('teams_evaluate', 'expression', expression);
      const page = await session.getPage();
      const wrapped = `(() => (${expression}))()`;
      const result = await page.evaluate(wrapped);
      let serialized;
      try {
        serialized = JSON.stringify(result);
      } catch {
        serialized = JSON.stringify({ unserializable: String(result) });
      }
      log.info('tool.teams_evaluate', 'done', {
        bytes: serialized?.length || 0,
      });
      return { content: [{ type: 'text', text: serialized ?? 'undefined' }] };
    });
  },
};
