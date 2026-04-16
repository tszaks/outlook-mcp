# Outlook MCP Design

## Goal

Create a simple, reusable Outlook MCP server for Greg that covers mail, calendar, contacts, and account onboarding without building a one-off custom flow.

## Approach

Use Microsoft Graph with delegated OAuth permissions and a public-client desktop app flow. Mirror the existing Gmail MCP structure so the server feels familiar and stays easy to maintain.

## Main Pieces

- `src/config.ts`: local account config and file layout
- `src/accounts.ts`: account resolution and health checks
- `src/graph-client.ts`: Microsoft Graph auth, token cache handling, and API helpers
- `src/index.ts`: MCP tool registration and argument validation

## Scope

- Mail tools for reading, searching, threading, sending, drafting, moving, marking read/unread, and deleting
- Calendar tools for listing calendars and creating/updating/deleting events
- Contact tools for listing, creating, updating, and deleting contacts
- Built-in account onboarding through MCP tools

## KISS Choices

- Use one consistent OAuth flow with `http://localhost`
- Keep storage file-based like the Gmail MCP
- Use direct Graph REST calls instead of a bigger SDK layer
- Start with broad but practical Outlook coverage rather than every Microsoft Graph edge case
