import fs from 'node:fs';
import { Paths } from './paths.js';

/**
 * Tee logger: every line goes to BOTH stderr (so Claude Code's MCP log
 * viewer picks it up) and the rolling file at %LOCALAPPDATA%\TeamsBrowserMcp\
 * logs\mcp.log (so we have a trail even after the MCP process exits).
 *
 * Every write is sync with O_APPEND so a `Get-Content -Wait` on the log
 * file sees output in real time. The old C# server only logged on
 * exceptions; that left 60-180s holes in the trace when tools hung. We
 * deliberately overcorrect here — verbose is cheap, silence is expensive.
 */

// Open once, keep the fd around. Node guarantees atomic per-write append
// semantics when the file is opened O_APPEND on POSIX; on Windows the OS
// serializes appends to the same fd which is what we want.
const fd = fs.openSync(Paths.logFile, 'a');

function formatDetails(details) {
  if (!details || typeof details !== 'object') return '';
  const parts = [];
  for (const [k, v] of Object.entries(details)) {
    if (v === undefined) continue;
    let out;
    if (v === null) out = 'null';
    else if (typeof v === 'string') {
      // Only quote strings that contain whitespace or special chars.
      out = /[\s"=]/.test(v) ? JSON.stringify(v) : v;
    } else if (typeof v === 'number' || typeof v === 'boolean') {
      out = String(v);
    } else if (v instanceof Error) {
      out = JSON.stringify(`${v.name}: ${v.message}`);
    } else {
      try { out = JSON.stringify(v); } catch { out = String(v); }
    }
    parts.push(`${k}=${out}`);
  }
  return parts.length ? ' ' + parts.join(' ') : '';
}

function write(level, component, event, details) {
  const ts = new Date().toISOString();
  const line = `[${ts}] [${level}] [${component}] ${event}${formatDetails(details)}\n`;
  try { fs.writeSync(fd, line); } catch { /* never break on logging */ }
  try { process.stderr.write(line); } catch { /* same */ }
}

export const log = {
  info(component, event, details) { write('info', component, event, details); },
  warn(component, event, details) { write('warn', component, event, details); },
  error(component, event, err, details) {
    const d = { ...(details || {}) };
    if (err) {
      d.err = err.message || String(err);
      if (err.stack) d.stack = err.stack.split('\n').slice(0, 6).join(' | ');
    }
    write('error', component, event, d);
  },

  /**
   * Wraps an async function with enter/exit/elapsed logs. If the wrapped
   * fn throws, logs the error and rethrows — so callers still see failures.
   * Use this around EVERY non-trivial async operation so any hang produces
   * an enter-without-exit pair pointing straight at the culprit.
   */
  async span(component, event, fn, details) {
    const started = Date.now();
    write('info', component, `${event}:enter`, details);
    try {
      const result = await fn();
      const elapsed = Date.now() - started;
      write('info', component, `${event}:exit`, { ok: true, elapsedMs: elapsed });
      return result;
    } catch (err) {
      const elapsed = Date.now() - started;
      write('error', component, `${event}:exit`, {
        ok: false,
        elapsedMs: elapsed,
        err: err?.message || String(err),
        stack: err?.stack?.split('\n').slice(0, 6).join(' | '),
      });
      throw err;
    }
  },
};
