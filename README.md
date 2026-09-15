# outlook-mac-mcp

MCP server for Microsoft Outlook, built for **personal Microsoft accounts**
(Hotmail, Outlook.com, Live) that run on the consumer infrastructure rather than
Exchange Online — and are therefore not supported by the official
"Claude for Outlook" add-in nor by the Microsoft 365 connector.

Backend: **Microsoft Graph API** with delegated, user-consented, least-privilege scopes.
Python, transport `stdio`, no telemetry.

> Status: **v1 in development.** Nothing below is validated yet — see
> [Validated environments](#validated-environments).

## Why not AppleScript

Community MCP servers for Outlook on Mac drive the app through AppleScript. That
path only works with **Legacy Outlook**. In August 2026 Microsoft cancelled the
planned AppleScript support in New Outlook, and the transition to New Outlook
ends in October 2026. This project deliberately does not depend on it.

The reference environment for this project runs **New Outlook** (the
"Legacy Outlook" toggle is present in the Outlook menu), where AppleScript
automation is unavailable. This was the deciding factor.

## Technical requirements

### Runtime

| Component | Requirement | Notes |
|---|---|---|
| macOS | 13 Ventura or later | Keychain used for token storage |
| Python | 3.12+ | Managed with `uv`; dependencies pinned in `uv.lock` |
| MCP client | Any stdio client (Claude Desktop, Claude Code) | Tested against Claude Desktop |
| Network | HTTPS to `login.microsoftonline.com` and `graph.microsoft.com` only | No other outbound calls |

### Microsoft account and app registration

| Item | Requirement |
|---|---|
| Account type | Personal Microsoft account (Hotmail / Outlook.com / Live) or work/school |
| App registration | One-time registration in Microsoft Entra ID (free; a personal account needs its own tenant, created via the Azure Portal) |
| Supported account types | **"Accounts in any organizational directory and personal Microsoft accounts"** |
| Authentication | Device-code flow (MSAL, public client, **no client secret**) |
| Authority | `https://login.microsoftonline.com/consumers` |
| Scopes requested by v1 | `Mail.Read`, `Calendars.Read`, `Contacts.Read`, `offline_access`, `User.Read` |
| Admin consent | Not required — all v1 scopes are user-consentable |

The app registration may list broader permissions (e.g. `Calendars.ReadWrite`)
than the server requests. The token only ever carries the scopes the current
version asks for; write scopes enter the code in later versions, behind a
config flag that is off by default, and trigger a new consent prompt.

### Not required

- Microsoft Outlook desktop app (the server talks to Graph, not to the app)
- Exchange Online / Microsoft 365 business tenant
- Accessibility or Automation permissions on macOS
- Node.js, Bun, or any non-Python runtime

## Validated environments

| OS | Python | MCP client | Account type | Outlook app installed | Status |
|---|---|---|---|---|---|
| macOS 27.0 (Golden Gate), Apple M5 | 3.12 | Claude Desktop | Personal (`@hotmail.com`) | Outlook for Mac 16.112.4 (26090911), New Outlook, M365 Subscription — not used by the server | **pending** |

A row moves to **validated** only after every v1 tool returns correct results
on that environment. Contributions of new rows are welcome — please include
exact versions.

## Scope

- **v1 (read-only):** unread emails, emails by folder, search by term, list
  folders, today's / upcoming events, contact search.
- **v2:** create event, mark as read — behind a config flag, off by default.
- **v3:** send email — behind a config flag, off by default, with explicit
  per-call confirmation.

## Security model

- Least privilege: the token can only do what the requested scopes allow, even
  if the server process is compromised.
- Tokens are stored in the macOS Keychain, never in files or environment
  variables.
- Email content is treated as untrusted input. The server never acts on
  instructions found inside message bodies.
- Structured JSON logs go to `stderr` (stdout is the MCP transport). Logs
  contain metadata and counts only — never subjects, bodies, addresses, or
  contact names.

## Architecture

```
domain/           entities and rules — no Graph, no MCP
application/      use cases depending on ports (interfaces)
infrastructure/   Graph API adapter — the only place that knows Microsoft
interface/mcp/    tool registration and input/output translation only
```

Dependencies point inward. Swapping the backend touches `infrastructure/` only.

## Development

```
uv sync
uv run ruff check .
uv run ruff format --check .
uv run mypy
uv run pytest
```

## Installation

Coming with v1.

## License

MIT
