## Outlook MCP

Local stdio MCP server for Outlook and Microsoft Graph, built as a Bun-first standalone repo.

### What It Exposes

- authenticated user identity lookup
- mail folder listing
- recent message listing
- Outlook conversation listing via `conversationId`
- local mail cache and search keyed by `conversationId`
- message search with Microsoft Graph `$search`
- full message fetch
- calendar listing and event CRUD
- send mail
- reply / reply-all
- Graph subscription list/create/renew/delete
- local Graph notification receiver scaffold and notification log inspection

### Runtime

- primary runtime: Bun
- package manager: Bun
- API surface: Microsoft Graph
- auth model: delegated OAuth user auth
- local state: XDG-style per-user state dir by default

### Identity Model

- this MCP is user-scoped only
- `OUTLOOK_USER_TOKEN` can override OAuth for debugging, but the normal path is the stored refreshable OAuth token
- Outlook conversation continuity should use `conversationId`, not individual message ids

### Calendar

The first calendar slice includes:

- `list_calendars`
- `list_events`
- `get_event`
- `create_event`
- `update_event`
- `delete_event`

`list_events` supports both collection mode and calendar-view mode. If you pass
`startDateTime` and `endDateTime`, it uses Graph calendar view.

### Local Mail Cache

This repo now has an explicit local mail cache backed by Bun SQLite:

- `sync_mail_folder`
- `search_cached_messages`
- `list_cached_conversation_messages`

The cache is keyed by `conversationId`, so Outlook thread continuity stays explicit.
Unlike the Webex MCP cleanup, this repo keeps live Graph search and cached search as
separate tools on purpose. There is no silent fallback between them.

`sync_mail_folder` is intentionally folder-scoped because Microsoft Graph message delta
tracks changes per folder.

### Install

```bash
cd outlook-mcp
bun install
```

### Microsoft App Setup

1. Create an Azure app registration.
2. Add a redirect URI that matches your env file.
   Example: `http://localhost:8776/oauth/callback`
3. Grant the delegated scopes in `.env.example`.
4. Copy `.env.example` to `.env.local` and fill in:
   - `OUTLOOK_CLIENT_ID`
   - `OUTLOOK_CLIENT_SECRET`
   - `OUTLOOK_TENANT_ID`
   - `OUTLOOK_REDIRECT_URI`
5. Run the OAuth login helper once.

### OAuth And Reauth

```bash
cd outlook-mcp
bun run auth:login
```

Useful commands:

```bash
bun run auth:status
bun run auth:refresh
bun run auth:clear
```

Notes:

- normal access-token expiry does not require a full reauth; `server.mjs` refreshes via the stored `refresh_token`
- if you change scopes, revoke consent, or the refresh token expires, run `auth:login` again
- when `--env-file` is passed, that file is authoritative over Bun auto-loaded env
- `auth:status` reports whether auth is coming from the token file or `OUTLOOK_USER_TOKEN`

### Subscription Notes

This repo can manage Graph subscriptions and now includes a separate notification receiver scaffold.

The MCP side exposes:

- `list_subscriptions`
- `create_subscription`
- `renew_subscription`
- `delete_subscription`
- `receiver_status`
- `list_received_notifications`

The receiver side is a separate process:

```bash
cd outlook-mcp
bun run start:receiver
```

It handles Microsoft Graph validation requests and appends received notification payloads
to a local NDJSON log file.

`create_subscription` can use the receiver scaffold defaults when you omit `notificationUrl`.

Important:

- Microsoft Graph notification URLs normally need a publicly reachable HTTPS endpoint
- this local receiver is a scaffold, so in practice you would front it with a tunnel or reverse proxy when creating live subscriptions
- if you expose it publicly, set and use `OUTLOOK_SUBSCRIPTION_CLIENT_STATE`; the receiver validates that value before accepting notifications

### Run

```bash
cd outlook-mcp
cp .env.example .env.local
bun run auth:login
bun run start
bun run start:receiver
```

### Codex Config

Add this to `~/.codex/config.toml`:

```toml
[mcp_servers.outlook]
command = "bun"
args = ["/absolute/path/to/outlook-mcp/server.mjs", "--env-file", "/absolute/path/to/outlook-mcp/.env.local"]
```

Then restart Codex.

### Useful Env Vars

- `OUTLOOK_OAUTH_TOKEN_FILE` overrides the token file location
- `OUTLOOK_MCP_INDEX_DB` overrides the SQLite index path
- `OUTLOOK_USER_TOKEN` forces a direct bearer token override for debugging
- `OUTLOOK_RECEIVER_HOST`, `OUTLOOK_RECEIVER_PORT`, and `OUTLOOK_RECEIVER_PATH` configure the local notification receiver
- `OUTLOOK_SUBSCRIPTION_CLIENT_STATE` is used as the default subscription integrity token and is validated by the receiver
- `OUTLOOK_NOTIFICATION_URL` and `OUTLOOK_LIFECYCLE_NOTIFICATION_URL` let you pin explicit Graph subscription URLs
- `OUTLOOK_NOTIFICATION_LOG_FILE` overrides the receiver log path
- `OUTLOOK_NOTIFICATION_MAX_BODY_BYTES` and `OUTLOOK_NOTIFICATION_MAX_LOG_BYTES` bound receiver memory/disk usage
