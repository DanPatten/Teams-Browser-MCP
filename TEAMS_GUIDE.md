# Teams Browser MCP — Navigation Guide

This MCP server gives you primitive control over a long-lived,
authenticated Microsoft Teams (`teams.microsoft.com/v2`) browser tab.
It does **not** ship hardcoded selectors for high-level operations like
"search" or "list channels" — you drive Teams yourself, the same way a
human would, using the primitives below.

## Tools

| Tool | Purpose |
|---|---|
| `authenticate` | Force a fresh sign-in. Call once at the start of a session, or after a tool errors with auth-related text. Blocks until the user signs in inside the Playwright Chrome window. |
| `teams_navigate` | Navigate the page to a `https://teams.microsoft.com/*` URL. |
| `teams_query` | `document.querySelectorAll(selector)` and return tag/text/attrs/visibility for each match. **Always start here** when you don't know the DOM. |
| `teams_click` | Real Playwright click on a CSS selector (with `nth` for disambiguation). |
| `teams_type` | Type into an input/textarea. React-friendly — controlled inputs notice. Set `submit: true` to press Enter after typing. |
| `teams_press_key` | Send a key to the page (`Enter`, `Escape`, `Control+K`, etc.). |
| `teams_wait_for` | Wait for a selector to become visible/hidden/attached/detached. |
| `teams_evaluate` | Escape hatch — run a JS expression and get its result back. |
| `teams_guide` | Returns this document. |

## How to drive Teams

**The golden rule: probe before you act.** Teams v2 is a Fluent UI
React app and its DOM is unstable across releases. Do not assume any
selector. Use `teams_query` to discover the current shape, then click
or type.

### A typical exploration loop

1. `teams_query` with a broad selector (`[data-tid]`, `[role="tree"] [role="treeitem"]`, `input`, `[aria-label*="Search"]` etc.) to find candidates.
2. Inspect the returned `text`, `attrs.aria-label`, `attrs.data-tid`, and `tag` to confirm you have the right element.
3. Re-query with a narrower selector that uniquely identifies the target.
4. `teams_click` / `teams_type` against that narrow selector.
5. `teams_wait_for` something the next step depends on (e.g. a results list, a chat pane).
6. Repeat from step 1 against the new state.

If a selector matches multiple elements, prefer narrowing the selector
over using `nth` — `nth` is fragile across re-renders.

### Useful starting selectors

These have been stable enough to be worth trying first, but **always
verify with `teams_query` before clicking**:

- App bar (left vertical rail with Activity / Chat / Teams / Calendar): `[data-tid='app-bar-2a']`, `div#app-bar`, or `[role='navigation'][aria-label*='App']`
- Global search box: `[data-tid='AUTOSUGGEST_INPUT']`
- Logged-in shell markers (handy for `teams_wait_for` after navigate): `[data-tid='app-bar-2a']`, `div#app-bar`
- Left rail tree (chats or teams list): `[role='tree']`, `[role='treeitem']`
- Chat / channel message in the open conversation pane: `[data-tid='chat-pane-message']`
- Search results — message hits: `[data-tid='all-messages'] > li[role='group']` (each `<li>` has `tabindex='0'` and an `aria-label` starting "Message from…" — that LI is what you click)
- Thread/replies pane after clicking a search-result message: `[data-tid='channel-replies-pane-message']`, `[data-tid='message-pane-body']`

If any of the above stops matching, fall back to `teams_query` with
broader selectors and rediscover.

### Common workflows

#### Open the global search box and run a query

```
teams_click   { selector: "[data-tid='AUTOSUGGEST_INPUT']" }
teams_type    { selector: "[data-tid='AUTOSUGGEST_INPUT']", text: "your query", submit: true }
teams_wait_for { selector: "[data-tid='all-messages']" }
teams_query   { selector: "[data-tid='all-messages'] > li", limit: 20 }
```

Then inspect the results, pick one, and click it:

```
teams_click   { selector: "[data-tid='all-messages'] > li[role='group']", nth: 0 }
teams_wait_for { selector: "[data-tid='message-pane-body']" }
teams_query   { selector: "[data-tid='channel-replies-pane-message']", limit: 50 }
```

#### Open a specific channel from the teams list

1. Click the Teams tab in the app bar (`teams_query` for the button by its `aria-label`/`title` containing "Teams").
2. `teams_query` for `[role='treeitem']` to find your team / channel.
3. `teams_click` the channel `treeitem`.
4. `teams_wait_for` the conversation pane.

#### Read recent messages in the open conversation

```
teams_query { selector: "[data-tid='chat-pane-message']", limit: 50 }
```

Each match's `text` contains the rendered message body; `attrs.aria-label`
often includes the author and timestamp.

## Auth and lifecycle

- The browser is launched on first use. If there's no saved session,
  `authenticate` (or any other tool) will open a Chrome window and
  block while you sign in.
- A network-level allowlist restricts the browser to Microsoft-owned
  hosts — no general web browsing.
- The session is reused across all tool calls. Storage state is
  persisted between MCP server restarts.
- If a tool starts failing because the session expired, call
  `authenticate` to force a fresh login.

## Don'ts

- **No screenshots.** This MCP intentionally has no screenshot tool.
  Use `teams_query` and DOM-based reasoning instead.
- **Don't navigate off teams.microsoft.com.** `teams_navigate` will
  reject other hosts, and the network allowlist will block third-party
  resources anyway.
- **Don't cache selectors across sessions.** The DOM drifts. Re-probe.
