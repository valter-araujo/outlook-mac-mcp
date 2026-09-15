# outlook-mac-mcp

MCP server for Microsoft Outlook, built for **personal Microsoft accounts**
(Hotmail, Outlook.com, Live) that run on the consumer infrastructure rather than
Exchange Online — and are therefore not supported by the official
"Claude for Outlook" add-in nor by the Microsoft 365 connector.

Backend: **Microsoft Graph API** with delegated, user-consented, least-privilege scopes.
Python, transport `stdio`, no telemetry.

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
| Scopes requested by default | `Mail.Read`, `Calendars.Read`, `Contacts.Read`, `offline_access`, `User.Read` |
| Scope added by `OUTLOOK_MCP_ENABLE_CALENDAR_WRITE=true` | `Calendars.ReadWrite` (see [Optional: calendar write](#optional-calendar-write)) |
| Admin consent | Not required — every scope above is user-consentable |

`offline_access`, `openid` and `profile` are reserved scopes: MSAL appends them
to every request on its own and raises if they are passed explicitly. They are
consented at the registration but never appear in the list the code asks for —
that consent is what yields the refresh token behind the Keychain cache.

The app registration may list broader permissions than the server requests. The
token only ever carries the scopes the running configuration asks for: the write
scope is requested only behind a config flag that is off by default, and turning it
on triggers a new sign-in and consent prompt.

### Not required

- Microsoft Outlook desktop app (the server talks to Graph, not to the app)
- Exchange Online / Microsoft 365 business tenant
- Accessibility or Automation permissions on macOS
- Node.js, Bun, or any non-Python runtime

## Validated environments

| OS | Python | MCP client | Account type | Outlook app installed | Status |
|---|---|---|---|---|---|
| macOS 27.0 (Golden Gate), Apple M5 | 3.14.7 | Claude Desktop | Personal (`@hotmail.com`) | Outlook for Mac 16.112.4 (26090911), New Outlook, M365 Subscription — not used by the server | **validated** 2026-09-14 |

A row moves to **validated** only after every v1 tool shipped so far returns correct
results on that environment, against a real mailbox. The row above covers
`list_unread_emails` reading a Hotmail inbox through Graph. Contributions of new rows
are welcome — please include exact versions.

## Scope

- **v1 (read-only):** unread emails, emails by folder, search by term, read one
  email by id (body included), list folders, today's / upcoming events, contact
  search.
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

### Integration tests

Tests marked `@pytest.mark.integration` reach the real Microsoft Graph API and are
deselected by default. They need a completed sign-in (see
[Installation](#3-sign-in-once)) and the client id in the environment:

```
export OUTLOOK_MCP_CLIENT_ID=<your client id>
uv run pytest -m integration
```

They read the signed-in mailbox but never write to it, and they provision their own
fixtures: no particular message has to exist, and a test skips rather than fails if the
mailbox yields nothing to work with.

One of them is load-bearing rather than incidental. `test_search_query_live.py` checks
that Graph parses a search term as literal text instead of as a query, by sending the
same string in the safe form and in the form this project used to send, and comparing
result counts against a sender taken from the mailbox itself. A unit test cannot show
this: it can only assert the bytes we send, and every unit test passed while terms were
in fact being executed as queries. Run it after touching anything about how a query is
quoted, escaped or encoded.

## Installation

### 1. Register the app in Microsoft Entra ID

Follow the table under [Microsoft account and app registration](#microsoft-account-and-app-registration).
Note the **Application (client) ID** — it is not a secret, but it identifies your
registration and does not belong in this repository.

### 2. Install

```sh
git clone https://github.com/valter-araujo/outlook-mac-mcp.git
cd outlook-mac-mcp
uv sync
```

### 3. Sign in once

The MCP server never starts a sign-in of its own: the device-code flow blocks for
minutes waiting for a browser, and a stdio server has nowhere to show the code,
because stdout carries the MCP protocol. Authenticate out of band instead:

```sh
export OUTLOOK_MCP_CLIENT_ID=<your client id>
uv run outlook-mac-mcp sign-in
```

The code and the URL are printed to stderr. Open the URL, enter the code, and approve
the requested scopes. The resulting token is stored in the macOS Keychain under the
service `outlook-mac-mcp`; it is never written to a file or an environment variable.

Repeat this only when the refresh token expires or you revoke consent.

### 4. Point Claude Desktop at the server

Edit `~/Library/Application Support/Claude/claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "outlook-mac-mcp": {
      "command": "uv",
      "args": ["--directory", "/absolute/path/to/outlook-mac-mcp", "run", "outlook-mac-mcp"],
      "env": {
        "OUTLOOK_MCP_CLIENT_ID": "<your client id>"
      }
    }
  }
}
```

Three things this configuration has to get right:

- **`OUTLOOK_MCP_CLIENT_ID` must be in the `env` block.** Claude Desktop launches the
  server as a GUI process, which does not inherit your shell environment: an `export` in
  `~/.zshrc` is invisible to it. The `env` block is the only thing the server sees.
- **Use an absolute path** in `--directory`. The launched process does not start in your
  project directory.
- **`uv` must be on the launcher's `PATH`,** which is likewise not your shell's. If the
  server fails to start, replace `"command": "uv"` with the absolute path from
  `which uv` (typically `~/.local/bin/uv`).

Restart Claude Desktop. `list_unread_emails` should appear in the tool list.

### Optional: log level

`OUTLOOK_MCP_LOG_LEVEL` sets the verbosity (default `INFO`). Structured JSON logs go to
stderr and carry metadata only — tool name, correlation id, duration, outcome and item
count, never subjects, addresses or bodies. Add it to the same `env` block.

### Optional: time zone

Calendar tools express every time in one zone and decide what "today" means from it.
By default that is the machine's time zone, as reported by the operating system. To
override it, add `OUTLOOK_MCP_TIMEZONE` with an IANA name (for example
`America/Sao_Paulo` or `Europe/Lisbon`) to the same `env` block. The name is validated
at startup: an unknown one stops both the server and `sign-in` with a
`ConfigurationError`. The resolved zone is logged at startup.

### Optional: calendar write

Off by default. Setting `OUTLOOK_MCP_ENABLE_CALENDAR_WRITE=true` in the `env` block
does two things: the server requests the `Calendars.ReadWrite` scope in addition to the
read scopes, and it registers the write tools. Accepted values are `true`, `false`,
`1` and `0`; anything else stops the server with a `ConfigurationError`.

Because the scope set changes, the token already in the Keychain no longer satisfies
it and every tool fails with "no usable cached credential" until you sign in again.
Run the sign-in with the flag set in the same shell, and approve the new consent
prompt, which now lists calendar write access:

```sh
export OUTLOOK_MCP_CLIENT_ID=<your client id>
export OUTLOOK_MCP_ENABLE_CALENDAR_WRITE=true
uv run outlook-mac-mcp sign-in
```

Turning the flag off again does not shrink the consent already granted; revoke it at
<https://account.live.com/consent/Manage> if you want the write permission gone.

With the flag on, two tools appear. Creating an event is always a two-step handshake:

1. `preview_event` takes the full details (subject, start, end with UTC offsets,
   optional location, all-day flag, up to 50 attendee addresses), validates them, and
   returns a one-line summary plus an opaque token. Nothing is written.
2. `create_event` takes only that token and performs the real change. A token works
   exactly once and only within the same server process; an unknown or already used
   token is refused, and a server restart discards every pending draft.

Both tool descriptions instruct the client to show the summary and get the user's
agreement before creating, and to get explicit confirmation of every detail before even
previewing when any of it came from email content. Email is untrusted input; a message
can be written to talk a model into putting something on your calendar, and the user,
never the email, decides.

### Troubleshooting

| Symptom | Cause |
|---|---|
| Tool fails with "no usable cached credential" | Step 3 was not completed, or the token was revoked. Run `sign-in` again. |
| Server fails to start with `ConfigurationError` | `OUTLOOK_MCP_CLIENT_ID` is missing from the `env` block, or `OUTLOOK_MCP_TIMEZONE` is not an IANA time zone name. |
| Server does not appear in Claude Desktop | `uv` is not on the launcher's `PATH`; use an absolute path for `command`. |

## License

MIT
