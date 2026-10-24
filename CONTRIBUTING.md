# Contributing to outlook-mcp

Thanks for helping. outlook-mcp is a small server with a narrow job, so the bar for changes is: does it keep the server **simple**, and does it keep the user's mailbox and tokens safe?

By taking part you agree to follow the [Code of Conduct](CODE_OF_CONDUCT.md).

## Before you start

- For bugs, open an issue with the template. Remove tokens and private mail text first.
- For features, open an issue first and describe the problem you want to solve.
- Small fixes (typos, clear bugs) can go straight to a pull request.
- Security problems go through [SECURITY.md](SECURITY.md), not a public issue.

## Setup

You need Node.js 20 or later. A Mac with `swiftc` also builds the OCR helper; other systems skip it.

```sh
git clone https://github.com/tszaks/outlook-mcp.git
cd outlook-mcp
npm install
npm run build
```

Run the server by hand with `npm start` (it speaks MCP over stdio). Use `OUTLOOK_MCP_CONFIG_DIR=/tmp/outlook-dev` so a test run does not touch your real accounts.

## Run the tests

```sh
npm run typecheck
npm test
```

`npm test` builds first, then runs `test/*.test.mjs` with the Node test runner. Tests do not call Microsoft.

## Code style

- **Match the surrounding code.** Same naming, same comment density, same idiom as the file you are editing. There is no formatter to run, so read the neighbors.
- TypeScript, ES modules. Import local files with the `.js` suffix.
- Call Microsoft Graph directly through `graph-client.ts`. Do not add a Graph SDK.
- No new dependencies without discussion.

## Safety rules

- A tool that permanently deletes data must call `requireConfirm` and list `confirm` and `reason` as required in its schema.
- Never log or return tokens, client secrets or the contents of `credentials.json` or `token-cache.json`.
- Add a new network destination only after discussion. Today the server talks only to Microsoft.

## Pull request checklist

- [ ] `npm run typecheck` and `npm test` pass.
- [ ] New behavior has a test.
- [ ] The change is small and focused.
- [ ] A new or changed tool is listed in the README tools table.
- [ ] A line added under `[Unreleased]` in [CHANGELOG.md](CHANGELOG.md) for anything a user would notice.
- [ ] No tokens, credentials or private mail anywhere in the diff.

## Using AI agents to contribute

Fine, as long as you read and understand what you submit and the checklist above is true. [AGENTS.md](AGENTS.md) tells coding agents how this repository works.

## License

By contributing you agree that your contribution is licensed under the [MIT License](LICENSE).
