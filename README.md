<h1 align="center">outlook-mcp</h1>

<p align="center">
  A Microsoft Outlook MCP server for mail, calendar and contacts, with built-in OAuth sign-in.
</p>

<p align="center">
  <a href="https://github.com/tszaks/outlook-mcp/actions/workflows/ci.yml"><img src="https://github.com/tszaks/outlook-mcp/actions/workflows/ci.yml/badge.svg" alt="CI"></a>
  <a href="https://github.com/tszaks/outlook-mcp/releases/latest"><img src="https://img.shields.io/github/v/release/tszaks/outlook-mcp" alt="Latest release"></a>
  <a href="LICENSE"><img src="https://img.shields.io/github/license/tszaks/outlook-mcp" alt="MIT license"></a>
  <img src="https://img.shields.io/badge/node-20%2B-brightgreen" alt="Node 20 or later">
</p>

Example conversation with an MCP client:

```text
You:    What is on my calendar tomorrow, and did anyone email me about it?
Client: (calls list_events, then search_emails)
        You have a 10:00 design review. Two emails mention it. The newest one
        asks to move it to 11:00. Want me to draft a reply?
You:    Yes, say 11:00 works.
Client: (calls create_draft, shows you the text, waits for you before send_draft)
```

outlook-mcp lets an AI client read and manage an Outlook mailbox through Microsoft Graph. It covers mail, calendar, contacts, categories and attachments. You sign in from inside your MCP client, so there is no separate login script. It works with personal Microsoft accounts and with Microsoft 365 work or school accounts.

It is for people who want an assistant to triage mail and manage a calendar without giving that assistant a mailbox password. You register your own Microsoft Entra app, so the tokens stay on your machine and the server talks only to Microsoft.

## Features

- **41 tools** across accounts, mail, drafts, calendar, contacts, categories and attachments.
- **Built-in OAuth onboarding.** `begin_account_auth` returns a sign-in URL; `finish_account_auth` completes it. Uses the authorization code flow with PKCE.
- **Several accounts at once.** Every tool takes an `account` id. `read_emails` reads across all enabled accounts when you leave it out.
- **Attachment text extraction.** PDF, DOCX, XLSX, PPTX and plain text are read in Node. Images go through Apple Vision OCR on macOS.
- **Guarded deletes.** Tools that permanently delete something refuse to run unless `confirm` is `true` and a `reason` is given.
- **Plain files, no database.** Config and tokens live in one folder you can inspect or delete.

## Install

You need Node.js 20 or later and a Microsoft Entra app registration (next section).

```bash
git clone https://github.com/tszaks/outlook-mcp.git
cd outlook-mcp
npm install
npm run build
```

`npm install` also runs `scripts/build-ocr.sh`. On a Mac with `swiftc` it compiles the OCR helper. Without `swiftc` it skips the step, and image attachments are saved to disk without extracted text.

### Azure app setup

1. Go to Azure Portal > Microsoft Entra ID > App registrations.
2. Create a new app registration.
3. Mark it as a public client / native app.
4. Add a redirect URI such as `http://localhost`.
5. Add these delegated Microsoft Graph permissions:
   - `openid`
   - `profile`
   - `offline_access`
   - `User.Read`
   - `Mail.ReadWrite`
   - `Mail.Send`
   - `Calendars.ReadWrite`
   - `Contacts.ReadWrite`
   - `MailboxSettings.ReadWrite` (used by `block_sender`, `unblock_sender` and `list_blocked_senders`)
6. Grant admin consent if your tenant requires it.
7. Save the `Application (client) ID`.

### Connect it to your client

Replace `/path/to/outlook-mcp` with the folder you cloned.

**Claude Code**

```bash
claude mcp add outlook -- node /path/to/outlook-mcp/dist/index.js
```

**Codex** (`~/.codex/config.toml`)

```toml
[mcp_servers.outlook]
command = "node"
args = ["/path/to/outlook-mcp/dist/index.js"]
```

**Claude Desktop** (`claude_desktop_config.json`)

```json
{
  "mcpServers": {
    "outlook": {
      "command": "node",
      "args": ["/path/to/outlook-mcp/dist/index.js"]
    }
  }
}
```

To keep config somewhere other than `~/.outlook-mcp`, set `OUTLOOK_MCP_CONFIG_DIR`. For example, with Claude Code: `claude mcp add outlook -e OUTLOOK_MCP_CONFIG_DIR=/custom/path/.outlook-mcp -- node /path/to/outlook-mcp/dist/index.js`. In the Codex and Claude Desktop files, add an `env` entry (`env = { OUTLOOK_MCP_CONFIG_DIR = "..." }` in TOML, `"env": { ... }` in JSON).

## Usage

### Sign in to an account

1. Call `begin_account_auth` with `account_id`, `email`, and either `credentials_json` or `credentials_path`.
2. Open the returned `auth_url` and sign in to Microsoft.
3. Copy the full URL the browser lands on after sign-in.
4. Pass that full URL to `finish_account_auth` as `authorization_code`.

The credentials are this JSON, with your own client ID:

```json
{
  "clientId": "YOUR-APP-CLIENT-ID",
  "tenantId": "common",
  "redirectUri": "http://localhost"
}
```

`tenantId` is `common` for personal and work accounts, or a specific tenant ID for one Microsoft 365 tenant. You can also pass `scopes` to override the default list.

### Tools

| Tool | What it does |
|---|---|
| `list_accounts` | List configured accounts and their sign-in status. |
| `begin_account_auth` | Start OAuth sign-in for an account. |
| `finish_account_auth` | Finish sign-in with the auth code or full redirected URL. |
| `list_mail_folders` | List mail folders. |
| `read_emails` | Read recent mail from one account, or all enabled accounts. |
| `search_emails` | Full-text search across mail. |
| `get_email_thread` | Get a conversation by `conversation_id` or from a `message_id`. |
| `send_email` | Send an email. |
| `create_draft` | Create a draft. |
| `update_draft` | Update a draft. |
| `list_drafts` | List drafts, newest edit first. |
| `search_drafts` | Search within drafts. |
| `send_draft` | Send an existing draft. |
| `delete_drafts` | Delete drafts. Needs `confirm=true` and `reason`. |
| `move_email` | Move a message to another folder. |
| `mark_as_read` | Mark messages read or unread. |
| `archive_emails` | Move messages to Archive. |
| `trash_emails` | Move messages to Deleted Items. |
| `delete_email` | Delete an email. Needs `confirm=true` and `reason`. |
| `mute_thread` | Move a whole conversation to Deleted Items. |
| `unsubscribe_from_email` | Unsubscribe using the message's `List-Unsubscribe` header. |
| `block_sender` | Add a sender to the blocked list (Focused Inbox override). |
| `unblock_sender` | Remove a sender from the blocked list. |
| `list_blocked_senders` | List blocked senders. |
| `get_attachment` | Save one attachment to disk and return extracted text. |
| `get_all_attachments` | Save every attachment on a message and return extracted text. |
| `list_categories` | List categories. |
| `create_category` | Create a category with a color. |
| `delete_category` | Delete a category. Needs `confirm=true` and `reason`. |
| `add_categories` | Add categories to a message. |
| `remove_categories` | Remove categories from a message, or all of them. |
| `list_calendars` | List calendars. |
| `list_events` | List events. Uses the default calendar if `calendar_id` is omitted. |
| `get_event` | Get one event with full body and attendees. |
| `create_event` | Create an event. |
| `update_event` | Update an event. |
| `delete_event` | Delete an event. Needs `confirm=true` and `reason`. |
| `list_contacts` | List or search contacts (filtering is done client-side). |
| `create_contact` | Create a contact. |
| `update_contact` | Update a contact. |
| `delete_contact` | Delete a contact. Needs `confirm=true` and `reason`. |

## Configuration

| Setting | Default | What it does |
|---|---|---|
| `OUTLOOK_MCP_CONFIG_DIR` | `~/.outlook-mcp` | Folder for accounts, credentials and token caches. |
| `MCP_ATTACHMENTS_DIR` | `~/Downloads/mcp-attachments` | Folder where attachments are saved. |

Files under the config folder:

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

Account ids may use letters, numbers, underscores and hyphens only.

## Using with agents

- Ask the agent to show you a draft (`create_draft`) before it calls `send_draft`. `send_email` sends at once and has no confirm step.
- Deletes of single items, drafts and categories need `confirm=true` and a `reason`. An agent should ask you before it sets them.
- `read_emails` aggregates across enabled accounts when `account` is omitted. Name the account when it matters.
- Extracted attachment text is capped (text files at 500 KB) and flagged with `textTruncated`.

## Privacy and permissions

The permissions you grant are the limit of what this server can do. With the scopes above it can read, send, move and delete your mail, create, change and delete calendar events and contacts, and manage categories and the blocked-sender list. Remove a scope in Entra to take that ability away.

- **Where data goes.** The server calls `login.microsoftonline.com` (sign-in) and `graph.microsoft.com` (your data). It has no telemetry and no other service. Mail content is returned to your MCP client, so it is also sent to whatever model that client uses.
- **One exception.** `unsubscribe_from_email` sends a GET request to the `http(s)` link in the message's `List-Unsubscribe` header, or sends an email to its `mailto:` address. That link comes from the sender, so use the tool only on mail you trust.
- **Local storage.** Credentials and tokens are plain JSON files in the config folder. Protect that folder, and do not commit it. The repo `.gitignore` already skips these file names.
- **Attachments on disk.** Downloaded attachments stay in the attachments folder until you delete them. Filenames are sanitized to block path traversal.

## FAQ

**Do I need a paid Azure account?** You need an app registration in a Microsoft Entra tenant. Registering an app is free. Work or school tenants may require admin consent.

**Why do I paste a URL back instead of getting redirected?** The redirect URI is `http://localhost`, and no local server is listening there. The browser shows an error page, but the address bar holds the code. Copy that full URL into `finish_account_auth`.

**Does OCR work on Linux or Windows?** No. OCR uses a Swift helper that needs macOS. Everything else works anywhere Node 20 runs.

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md), [SECURITY.md](SECURITY.md) and the [Code of Conduct](CODE_OF_CONDUCT.md). Coding agents should read [AGENTS.md](AGENTS.md). The original design notes are in [docs/plans](docs/plans/2026-03-12-outlook-mcp-design.md).

## License

[MIT](LICENSE)
