# outlook-mcp

An MCP server (stdio) for Microsoft Outlook mail, calendar and contacts over Microsoft Graph, with built-in OAuth onboarding. TypeScript, ES modules, Node 20+.

This file is for AI agents that work **on** this code. If you are an agent that **uses** the server, read the "Using with agents" section of [README.md](README.md).

## Build and test
- Install: `npm install` (also runs `scripts/build-ocr.sh`, which needs macOS `swiftc` and skips itself otherwise).
- Typecheck: `npm run typecheck`
- Build: `npm run build` (output in `dist/`, git-ignored)
- Test: `npm test` (builds, then runs `test/*.test.mjs` with `node --test`). Tests never call Microsoft.
- Run by hand: `OUTLOOK_MCP_CONFIG_DIR=/tmp/outlook-dev npm start`. Always point the config dir at a temp folder so a test run never touches real accounts or tokens.

## Code map
- `src/index.ts`: the MCP server. One `ListToolsRequestSchema` handler that declares every tool and its schema, and one `CallToolRequestSchema` handler that implements them. A new tool needs an entry in both.
- `src/graph-client.ts`: scopes (`OUTLOOK_SCOPES`), OAuth with MSAL (auth URL, code exchange, token cache), and every Graph REST call (`graphRequest`).
- `src/accounts.ts`: account lookup and health checks.
- `src/config.ts`: config folder layout (`OUTLOOK_MCP_CONFIG_DIR`, default `~/.outlook-mcp`) and `accounts.json`.
- `src/attachments.ts`: attachment saving (`MCP_ATTACHMENTS_DIR`), filename sanitizing, text extraction (pdf-parse, mammoth, xlsx, officeparser, OCR).
- `vendor/ocr/ocr.swift`: Apple Vision OCR helper, compiled to `vendor/ocr/ocr-bin`.
- `test/`: Node test runner tests against `dist/`.
- `docs/plans/`: original design notes.

## Conventions
- Match the surrounding code: naming, comment density, idiom. Import local modules with the `.js` suffix.
- Call Graph directly through `graph-client.ts`. No Graph SDK, no new dependencies without discussion.
- A tool that permanently deletes data calls `requireConfirm` and lists `confirm` and `reason` as required.
- Every tool takes an `account` id except the account onboarding tools. Keep that shape.
- When you add or change a tool, update the tools table in `README.md` and the tool count in its Features list, and add a line to `CHANGELOG.md`.

## Do not
- Do not log, return or commit tokens, client secrets, `credentials.json` or `token-cache.json`.
- Do not add a network destination other than `login.microsoftonline.com` and `graph.microsoft.com` without discussion.
- Do not write to stdout from server code: stdout carries the MCP protocol. Log with `console.error`.
- Do not run the server against a real mailbox while testing a change that sends, moves or deletes mail.
