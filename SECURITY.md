# Security policy

## Supported versions

Only the latest release of outlook-mcp gets security fixes.

| Version | Supported |
|---|---|
| 1.2.x (latest) | Yes |
| Older | No |

## Reporting a vulnerability

Please report security problems privately, not in a public issue.

Use GitHub private vulnerability reporting: open the **Security** tab of this repository and choose **Report a vulnerability**. If that is not available to you, email tyler@szakacsmedia.com.

Please include:

- What you found and what an attacker could do with it.
- Steps to reproduce.
- Your Node.js version and outlook-mcp version.

You can expect a first reply within a few days. Once a fix is released we will credit you in the changelog unless you prefer to stay anonymous.

## Sensitive surface

This server is given access to a Microsoft mailbox. With the permissions it asks for, a connected AI client can read, send, move and delete mail, change calendar events and edit contacts. These areas matter most:

- Anything that leaks or logs OAuth tokens, `credentials.json` or `token-cache.json` (kept under `~/.outlook-mcp/` by default).
- Anything that lets a tool act on an account the user did not name, or skip the `confirm=true` check on a destructive tool.
- Attachment handling: filename sanitizing and writes outside the attachments folder (`~/Downloads/mcp-attachments/` by default).
- OAuth state handling in the sign-in flow.

The server talks only to Microsoft (`login.microsoftonline.com` and `graph.microsoft.com`). See the [README](README.md#privacy-and-permissions).
