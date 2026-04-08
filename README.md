# Teams Browser MCP

An MCP server that gives your AI assistant access to Microsoft Teams. No
Microsoft Graph API. No webhooks. No tenant admin. No app registration. You
sign in once in a normal Chrome window and it just works.

## Why this exists

Every other "Teams integration" path forces you to either beg your IT
department for Graph API permissions, register an Azure app, configure
webhooks, or hand a third party your credentials. For most people that's a
dead end — and even when it isn't, you've now widened your attack surface
with a long-lived API token sitting in some config file.

This MCP takes a different route: it talks to Teams the same way you do, in
a real browser, signed in as you. There is no separate identity, no extra
permissions to grant, and no token to leak. Your AI sees exactly what you
see in Teams — nothing more — and the session lives entirely in a sandboxed
Chrome profile on your own machine. If you can open Teams in a browser, this
works. If your admin revokes your account, this stops working. That's the
whole security model, and it's the same one you already trust.

## How it works

When the MCP starts, it launches a dedicated Chrome window via Playwright
and points it at `teams.microsoft.com`. The first time you use it, you sign
in inside that window — normally, with SSO and MFA and whatever else your
organization requires. Playwright captures the resulting browser session
(cookies, local storage, the works) and saves it to a file under
`%LOCALAPPDATA%\TeamsBrowserMcp\` so the next launch picks up where you
left off. If Teams ever expires your session, the same window pops back to
the front and waits for you to sign in again.

That window is locked down: a network-level allowlist blocks every host
that isn't part of the Microsoft / Teams ecosystem, so the browser can't be
hijacked into a general-purpose web agent and third-party trackers can't
load.

Once you're signed in, the MCP exposes a small set of primitive tools —
query the DOM, click an element, type into a field, wait for something to
appear, navigate within Teams. Your AI uses these the way a human would:
look at the page, find the thing it wants, click it. There are deliberately
**no** hardcoded "search Teams" or "list channels" tools, because Teams'
DOM changes constantly and baked-in scrapers rot within weeks. The repo
ships a `TEAMS_GUIDE.md` (also returned by the `teams_guide` tool) that
teaches the calling AI how to navigate Teams with the primitives.

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

## Logging

Every non-trivial operation is wrapped in a `log.span(...)` call that emits
enter + exit + elapsed-ms lines. A hang is impossible to lose: an entered
span without a matching exit points you straight at the culprit. Log lines
go to both stderr (so Claude Code's `/mcp logs` sees them) and the file
above. Every Playwright call, every DOM `evaluate`, every cookie-profile
scan, every auth poll tick is logged.
