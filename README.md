# Teams Browser MCP

A Node.js MCP server that lets an LLM drive Microsoft Teams through a real,
user-authenticated Chrome session. Playwright owns the entire auth lifecycle:
on first use it opens a Chrome window, you sign in directly inside it, and the
resulting session (`storageState`) is persisted so future process starts skip
the sign-in entirely. If Teams ever kicks you back to the login page, the
Playwright window pops to the front and polls until you're signed in again.

No tray app, no second process, no build step, no cookie scraping — the MCP
server is plain JavaScript running on Node, and owns the auth lifecycle itself.

**Design philosophy:** Teams' DOM drifts constantly across releases, so this
MCP intentionally **does not ship hardcoded scrapers** for "search",
"list channels", etc. Instead it exposes a small set of selector-based
primitives — `teams_query`, `teams_click`, `teams_type`, `teams_navigate`,
`teams_wait_for`, `teams_press_key`, `teams_evaluate` — and ships a navigation
guide (`TEAMS_GUIDE.md`, also exposed via the `teams_guide` tool) that the
calling LLM reads to drive Teams the same way a human would.

## Layout

| Path | Purpose |
|---|---|
| `src/index.js` | MCP server bootstrap + tool registration (stdio transport). |
| `src/logger.js` | Tee logger — every line goes to stderr AND `%LOCALAPPDATA%\TeamsBrowserMcp\logs\mcp.log`. |
| `src/teamsSession.js` | Gated Playwright lifecycle; owns the full auth flow (cold start, warm start, mid-session expiry). URL allowlist restricts the browser to Microsoft-owned hosts. |
| `src/authState.js` | Thin wrapper around `state.json` (exists / delete); Playwright does the actual read/write. |
| `src/tools/authenticate.js` | The `authenticate` tool. |
| `src/tools/primitives.js` | The 7 primitive browser tools. |
| `src/tools/teamsGuide.js` | Returns `TEAMS_GUIDE.md` to calling LLMs. |
| `TEAMS_GUIDE.md` | Navigation guide: how to drive Teams with the primitives, useful starting selectors, common workflows. |

Chrome currently launches **headed** so you can watch what's happening. Flip
`headless: true` in `src/teamsSession.js` when you're ready to background it.

## Install

> **Hand this section to your AI assistant** (Claude Code, Cursor, Cline,
> Windsurf, etc.). It contains everything an agent needs to clone, install,
> and register the MCP — you don't need to follow it manually.

**Prompt for your AI assistant:**

> Install the Teams Browser MCP from `https://github.com/DanPatten/Teams-Browser-MCP`
> and register it with my MCP client.
>
> Steps you should perform:
>
> 1. Clone the repo to a sensible local path (ask me where if unsure).
> 2. Run `npm install` in the cloned directory. Requires Node 20+.
> 3. Ensure Google Chrome is available. If not, run `npx playwright install chrome`.
>    The MCP launches Chrome with `channel: 'chrome'` on purpose — bundled
>    Chromium triggers Teams' bot detection.
> 4. Register the MCP with whichever MCP client I use. The server entry point
>    is `node <repo-path>/src/index.js`. It speaks stdio. Examples:
>    - **Claude Code:** `claude mcp add teams --scope user node <repo-path>/src/index.js`
>    - **Claude Desktop / generic JSON config:**
>      ```json
>      {
>        "mcpServers": {
>          "teams": {
>            "command": "node",
>            "args": ["<repo-path>/src/index.js"]
>          }
>        }
>      }
>      ```
>    - For other clients (Cursor, Cline, Windsurf, etc.), use their standard
>      MCP config format with the same `command` + `args`.
> 5. After registering, restart / reconnect the MCP client so it picks up the
>    new server.
> 6. Tell me to call the `authenticate` tool once. The first call will open a
>    Chrome window — I sign in to Teams inside that window, the session is
>    persisted to `%LOCALAPPDATA%\TeamsBrowserMcp\state.json`, and future
>    calls reuse it.
> 7. **Important:** before you (the AI) start using this MCP for real work,
>    call the `teams_guide` tool once and read the returned guide. It tells
>    you how to drive Teams via the primitive tools. Do not assume selectors
>    — always probe with `teams_query` first.

## Verifying the running MCP is on your latest edits

There is no compiled binary. On every `/mcp` reconnect, Claude Code respawns
`node src/index.js`, which reads the source files fresh — so "did my edit
land?" has a trivial answer: yes, as long as you reconnected. The server
prints a **boot** line on start that includes the entry file mtime; tail the
log to confirm:

```pwsh
Get-Content -Wait "$env:LOCALAPPDATA\TeamsBrowserMcp\logs\mcp.log"
```

You'll see lines like:

```
[2026-04-08T02:02:31.771Z] [info] [server] boot pid=22976 node=v24.11.0 ... entryMtime=2026-04-08T02:00:44.559Z version=0.2.0
```

## First use

The MCP handshake is instant — Playwright doesn't launch until a tool needs
it. On first use:

1. Call `authenticate` (or any other tool). The server launches a Playwright
   Chrome window and navigates to Teams.
2. **If `state.json` already has a valid session** (warm start), Teams loads
   straight into the app shell and the tool returns within seconds.
3. **Otherwise** (cold start, or expired tokens), the Playwright window is
   brought to the front on the Teams sign-in page and a banner is printed to
   stderr telling you to sign in *inside that window*. Polling runs every 2s;
   within ~2s of the logged-in DOM markers appearing, Playwright saves the
   refreshed `storageState` to disk and your tool call completes.

If Teams ever bounces you back to the login page mid-session (tokens
expired, signed out elsewhere), the next tool call will pop the same
Playwright window to the front and run the sign-in loop again — no tear-down,
no second window.

You can also call `authenticate` explicitly to delete `state.json` and force
a fresh sign-in (for example, to switch accounts).

### Aborting a sign-in

If you want to bail out of the sign-in poll loop without waiting for the 5
minute timeout, create this sentinel file:

```pwsh
New-Item "$env:LOCALAPPDATA\TeamsBrowserMcp\abort-auth"
```

Within ~2s the in-flight tool call will throw and the file is cleaned up.

## Tools

- `authenticate` — delete saved state and re-run the interactive sign-in flow.
- `teams_guide` — returns `TEAMS_GUIDE.md`. The calling LLM should read this
  first.
- `teams_navigate` — navigate the page within `https://teams.microsoft.com/*`.
- `teams_query` — `document.querySelectorAll(selector)`; returns trimmed
  tag/text/attrs/visibility for each match. The primary exploration tool.
- `teams_click` — Playwright click on a CSS selector (with `nth`).
- `teams_type` — React-friendly fill into an input/textarea (`submit: true`
  presses Enter after).
- `teams_press_key` — keyboard input (`Enter`, `Escape`, `Control+K`, etc.).
- `teams_wait_for` — wait for a selector to reach a state.
- `teams_evaluate` — escape hatch: run a JS expression in the page.

## Re-authenticating

To force a fresh login (tokens expired, or you signed out elsewhere), either
delete `%LOCALAPPDATA%\TeamsBrowserMcp\state.json` or call the `authenticate`
tool. It tears down the existing Playwright context, deletes state.json, and
restarts the auth flow from scratch.

## Logging

Every non-trivial operation is wrapped in a `log.span(...)` call that emits
enter + exit + elapsed-ms lines. A hang is impossible to lose: an entered
span without a matching exit points you straight at the culprit. Log lines
go to both stderr (so Claude Code's `/mcp logs` sees them) and the file
above. Every Playwright call, every DOM `evaluate`, every cookie-profile
scan, every auth poll tick is logged.
