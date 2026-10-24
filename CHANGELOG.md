# Changelog

All notable changes to outlook-mcp are listed here. The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and the project uses [Semantic Versioning](https://semver.org/).

## [Unreleased]

### Added
- LICENSE file (MIT).
- Documentation: tools table, client setup for Claude Code, Codex and Claude Desktop, privacy section, CONTRIBUTING, SECURITY, Code of Conduct, AGENTS.md and GitHub templates.
- CI workflow that runs typecheck, build and tests.

### Fixed
- README listed the Graph permissions without `MailboxSettings.ReadWrite`, which the server requests and the blocked-sender tools use.

## [1.2.0] - 2026-05-18

### Added
- 18 tools for feature parity with the ghub (Gmail) server: `search_emails`, `list_drafts`, `search_drafts`, `send_draft`, `delete_drafts`, `archive_emails`, `trash_emails`, `mute_thread`, `unsubscribe_from_email`, `block_sender`, `unblock_sender`, `list_blocked_senders`, `get_event`, `list_categories`, `create_category`, `delete_category`, `add_categories`, `remove_categories`.

## [1.1.0] - 2026-04-16

### Added
- Initial server (package version 1.0.0 in the first commit, never tagged): account onboarding through OAuth, and mail, calendar and contact tools on Microsoft Graph.
- Attachment extraction (`get_attachment`, `get_all_attachments`): PDF, DOCX, XLSX, PPTX and text, plus OCR for images on macOS.
