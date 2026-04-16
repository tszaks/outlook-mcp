# Outlook MCP Server

A Microsoft Graph powered MCP server for Outlook mail, calendar, and contacts with built-in OAuth onboarding.

## What It Covers

- Mail: folders, read/search, threads, send, draft, move, read/unread, delete
- Calendar: calendars, list events, create events, update events, delete events
- Contacts: list/search, create, update, delete
- Account onboarding: connect Outlook through OAuth from MCP tools

## Requirements

- Node.js 20+
- A Microsoft Entra app registration
- Microsoft Graph delegated permissions for:
  - `openid`
  - `profile`
  - `offline_access`
  - `User.Read`
  - `Mail.ReadWrite`
  - `Mail.Send`
  - `Calendars.ReadWrite`
  - `Contacts.ReadWrite`

## Azure App Setup

1. Go to Azure Portal > Microsoft Entra ID > App registrations.
2. Create a new app registration.
3. Mark it as a public client / native app.
4. Add a redirect URI such as `http://localhost`.
5. Add the delegated Microsoft Graph permissions listed above.
6. Grant admin consent if your tenant requires it.
7. Save the `Application (client) ID`.

## Credentials Format

Use either a JSON file or pass JSON directly to `begin_account_auth`.

```json
{
  "clientId": "YOUR-APP-CLIENT-ID",
  "tenantId": "common",
  "redirectUri": "http://localhost"
}
```

`tenantId` can be:

- `common` for personal + work accounts
- a specific tenant ID for one Microsoft 365 tenant

## Install

```bash
cd /Users/tyler/Projects/MCP-Servers/outlook-mcp
npm install
npm run build
```

## MCP Config

```json
{
  "mcpServers": {
    "outlook": {
      "command": "node",
      "args": [
        "/Users/tyler/Projects/MCP-Servers/outlook-mcp/dist/index.js"
      ]
    }
  }
}
```

Optional custom config directory:

```json
{
  "mcpServers": {
    "outlook": {
      "command": "node",
      "args": [
        "/Users/tyler/Projects/MCP-Servers/outlook-mcp/dist/index.js"
      ],
      "env": {
        "OUTLOOK_MCP_CONFIG_DIR": "/custom/path/.outlook-mcp"
      }
    }
  }
}
```

## Account Onboarding Flow

1. Call `begin_account_auth` with:
   - `account_id`
   - `email`
   - `credentials_json` or `credentials_path`
2. Open the returned `auth_url`.
3. Sign in to Microsoft.
4. Copy the full redirected URL from the browser.
5. Pass that full URL into `finish_account_auth.authorization_code`.

The server stores config by default in:

```text
~/.outlook-mcp/
├── accounts.json
└── accounts/
    └── <account-id>/
        ├── credentials.json
        ├── token-cache.json
        ├── pending-auth.json
        └── meta.json
```

## Tool Summary

### Accounts

- `list_accounts`
- `begin_account_auth`
- `finish_account_auth`

### Mail

- `list_mail_folders`
- `read_emails`
- `get_email_thread`
- `send_email`
- `create_draft`
- `update_draft`
- `move_email`
- `mark_as_read`
- `delete_email`

### Calendar

- `list_calendars`
- `list_events`
- `create_event`
- `update_event`
- `delete_event`

### Contacts

- `list_contacts`
- `create_contact`
- `update_contact`
- `delete_contact`

## Notes

- `read_emails` aggregates across enabled accounts when `account` is omitted.
- `list_events` uses the default calendar when `calendar_id` is omitted.
- `list_contacts` does lightweight client-side filtering when `query` is provided.
