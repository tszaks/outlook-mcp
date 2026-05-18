#!/usr/bin/env node

import { promises as fs } from 'node:fs';
import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
  type CallToolResult,
} from '@modelcontextprotocol/sdk/types.js';
import {
  ensureConfigLayout,
  getAccountPaths,
  getConfigRoot,
  getDefaultAccountPaths,
  loadAccountsConfig,
  saveAccountsConfig,
  upsertAccount,
  validateAccountId,
  type AccountConfig,
} from './config.js';
import {
  getAccountHealth,
  resolveReadAccounts,
  resolveWriteAccount,
} from './accounts.js';
import {
  beginAuthFromCredentials,
  finishAuthFromCredentials,
  OutlookAccountClient,
  readCredentialsFile,
  type AttachmentMetadata,
  type OutlookCredentials,
} from './graph-client.js';
import { saveAndExtract, type AttachmentContent } from './attachments.js';

interface BeginAuthArgs {
  account_id: string;
  email: string;
  display_name?: string;
  credentials_json?: unknown;
  credentials_path?: string;
}

interface FinishAuthArgs {
  account_id: string;
  authorization_code: string;
}

interface GetAttachmentArgs {
  account: string;
  email_id: string;
  attachment_id: string;
}

interface GetAllAttachmentsArgs {
  account: string;
  email_id: string;
}

function textResult(text: string): CallToolResult {
  return {
    content: [
      {
        type: 'text',
        text,
      },
    ],
  };
}

function requireString(value: unknown, field: string): string {
  if (typeof value !== 'string' || !value.trim()) {
    throw new Error(`Missing or invalid '${field}'`);
  }
  return value.trim();
}

function optionalString(value: unknown): string | undefined {
  if (typeof value === 'string' && value.trim()) {
    return value.trim();
  }
  return undefined;
}

function optionalBoolean(value: unknown, field: string): boolean | undefined {
  if (value === undefined || value === null) return undefined;
  if (typeof value === 'boolean') return value;
  if (typeof value === 'string') {
    const normalized = value.trim().toLowerCase();
    if (['true', '1', 'yes'].includes(normalized)) return true;
    if (['false', '0', 'no'].includes(normalized)) return false;
  }
  throw new Error(`Missing or invalid '${field}'`);
}

function optionalNumber(value: unknown, field: string, fallback: number): number {
  if (value === undefined || value === null) return fallback;
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  if (typeof value === 'string' && value.trim()) {
    const parsed = Number(value);
    if (Number.isFinite(parsed)) return parsed;
  }
  throw new Error(`Missing or invalid '${field}'`);
}

function requireStringArray(value: unknown, field: string): string[] {
  if (!Array.isArray(value)) {
    throw new Error(`Missing or invalid '${field}'`);
  }

  const normalized = value.map((item) => String(item).trim()).filter(Boolean);
  if (normalized.length === 0) {
    throw new Error(`Missing or invalid '${field}'`);
  }

  return normalized;
}

function requireConfirm(confirm: unknown, reason: unknown): void {
  if (confirm !== true) {
    throw new Error('This destructive action requires confirm=true.');
  }
  if (typeof reason !== 'string' || !reason.trim()) {
    throw new Error('This destructive action requires a non-empty reason.');
  }
}

function formatJson(value: unknown): string {
  return JSON.stringify(value, null, 2);
}

function defaultRangeStart(): string {
  return new Date().toISOString();
}

function defaultRangeEnd(): string {
  return new Date(Date.now() + 30 * 24 * 60 * 60 * 1000).toISOString();
}

async function readCredentialsInput(
  configRoot: string,
  accountId: string,
  args: BeginAuthArgs,
): Promise<{ credentials: OutlookCredentials; credentialsPath: string }> {
  const paths = getDefaultAccountPaths(configRoot, accountId);

  if (args.credentials_json) {
    await fs.mkdir(paths.accountDir, { recursive: true });
    await fs.writeFile(paths.credentialsPath, `${JSON.stringify(args.credentials_json, null, 2)}\n`, 'utf8');
    return {
      credentials: await readCredentialsFile(paths.credentialsPath),
      credentialsPath: paths.credentialsPath,
    };
  }

  if (args.credentials_path) {
    return {
      credentials: await readCredentialsFile(args.credentials_path),
      credentialsPath: args.credentials_path,
    };
  }

  throw new Error('Provide either credentials_json or credentials_path.');
}

class OutlookMcpServer {
  private readonly server: Server;
  private readonly configRoot: string;

  constructor() {
    this.server = new Server(
      {
        name: 'outlook-mcp',
        version: '1.0.0',
      },
      {
        capabilities: {
          tools: {},
        },
      },
    );

    this.configRoot = getConfigRoot();
    this.setupHandlers();

    this.server.onerror = (error) => {
      console.error('[outlook-mcp] MCP error:', error);
    };

    process.on('SIGINT', async () => {
      await this.server.close();
      process.exit(0);
    });
  }

  private setupHandlers(): void {
    this.server.setRequestHandler(ListToolsRequestSchema, async () => ({
      tools: [
        {
          name: 'list_accounts',
          description: 'List all configured Outlook accounts and their authentication/health status.',
          inputSchema: {
            type: 'object',
            properties: {},
            additionalProperties: false,
          },
        },
        {
          name: 'begin_account_auth',
          description: 'Start OAuth onboarding for an Outlook account using Microsoft app credentials.',
          inputSchema: {
            type: 'object',
            properties: {
              account_id: { type: 'string' },
              email: { type: 'string' },
              display_name: { type: 'string' },
              credentials_json: {
                description: 'Credentials object containing at least clientId and optionally tenantId/redirectUri/scopes.',
              },
              credentials_path: { type: 'string' },
            },
            required: ['account_id', 'email'],
            additionalProperties: false,
          },
        },
        {
          name: 'finish_account_auth',
          description: 'Complete OAuth onboarding using either the auth code or the full redirected URL.',
          inputSchema: {
            type: 'object',
            properties: {
              account_id: { type: 'string' },
              authorization_code: { type: 'string' },
            },
            required: ['account_id', 'authorization_code'],
            additionalProperties: false,
          },
        },
        {
          name: 'list_mail_folders',
          description: 'List Outlook mail folders for an account.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string' },
            },
            required: ['account'],
            additionalProperties: false,
          },
        },
        {
          name: 'read_emails',
          description: 'Read recent Outlook emails from one account or aggregate across all enabled accounts when account is omitted.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string' },
              folder: { type: 'string' },
              query: { type: 'string' },
              max_results: { type: 'number', default: 20 },
              include_body: { type: 'boolean', default: false },
            },
            additionalProperties: false,
          },
        },
        {
          name: 'get_email_thread',
          description: 'Get an Outlook conversation thread by conversation_id or by starting from a message_id.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string' },
              message_id: { type: 'string' },
              conversation_id: { type: 'string' },
            },
            required: ['account'],
            additionalProperties: false,
          },
        },
        {
          name: 'send_email',
          description: 'Send an email from an Outlook account.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string' },
              to: { type: 'string' },
              subject: { type: 'string' },
              body: { type: 'string' },
              cc: { type: 'string' },
              bcc: { type: 'string' },
              html: { type: 'boolean' },
            },
            required: ['account', 'to', 'subject', 'body'],
            additionalProperties: false,
          },
        },
        {
          name: 'create_draft',
          description: 'Create an Outlook draft message.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string' },
              to: { type: 'string' },
              subject: { type: 'string' },
              body: { type: 'string' },
              cc: { type: 'string' },
              bcc: { type: 'string' },
              html: { type: 'boolean' },
            },
            required: ['account', 'to', 'subject', 'body'],
            additionalProperties: false,
          },
        },
        {
          name: 'update_draft',
          description: 'Update an Outlook draft message by message_id.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string' },
              message_id: { type: 'string' },
              to: { type: 'string' },
              subject: { type: 'string' },
              body: { type: 'string' },
              cc: { type: 'string' },
              bcc: { type: 'string' },
              html: { type: 'boolean' },
            },
            required: ['account', 'message_id'],
            additionalProperties: false,
          },
        },
        {
          name: 'move_email',
          description: 'Move an Outlook message to another folder.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string' },
              message_id: { type: 'string' },
              destination_folder_id: { type: 'string' },
            },
            required: ['account', 'message_id', 'destination_folder_id'],
            additionalProperties: false,
          },
        },
        {
          name: 'mark_as_read',
          description: 'Mark Outlook messages as read or unread.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string' },
              message_ids: {
                type: 'array',
                items: { type: 'string' },
              },
              is_read: { type: 'boolean', default: true },
            },
            required: ['account', 'message_ids'],
            additionalProperties: false,
          },
        },
        {
          name: 'delete_email',
          description: 'Delete an Outlook email (destructive). Requires confirm=true and reason.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string' },
              message_id: { type: 'string' },
              confirm: { type: 'boolean' },
              reason: { type: 'string' },
            },
            required: ['account', 'message_id', 'confirm', 'reason'],
            additionalProperties: false,
          },
        },
        {
          name: 'list_calendars',
          description: 'List Outlook calendars for an account.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string' },
            },
            required: ['account'],
            additionalProperties: false,
          },
        },
        {
          name: 'list_events',
          description: 'List Outlook events. When calendar_id is omitted, this uses the default calendar.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string' },
              calendar_id: { type: 'string' },
              start: { type: 'string' },
              end: { type: 'string' },
              max_results: { type: 'number', default: 50 },
            },
            required: ['account'],
            additionalProperties: false,
          },
        },
        {
          name: 'create_event',
          description: 'Create an Outlook calendar event.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string' },
              calendar_id: { type: 'string' },
              subject: { type: 'string' },
              start: { type: 'string' },
              end: { type: 'string' },
              is_all_day: { type: 'boolean' },
              location: { type: 'string' },
              body: { type: 'string' },
              attendees: { type: 'string' },
            },
            required: ['account', 'subject', 'start', 'end'],
            additionalProperties: false,
          },
        },
        {
          name: 'update_event',
          description: 'Update an Outlook event.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string' },
              event_id: { type: 'string' },
              subject: { type: 'string' },
              start: { type: 'string' },
              end: { type: 'string' },
              is_all_day: { type: 'boolean' },
              location: { type: 'string' },
              body: { type: 'string' },
              attendees: { type: 'string' },
            },
            required: ['account', 'event_id'],
            additionalProperties: false,
          },
        },
        {
          name: 'delete_event',
          description: 'Delete an Outlook event (destructive). Requires confirm=true and reason.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string' },
              event_id: { type: 'string' },
              confirm: { type: 'boolean' },
              reason: { type: 'string' },
            },
            required: ['account', 'event_id', 'confirm', 'reason'],
            additionalProperties: false,
          },
        },
        {
          name: 'list_contacts',
          description: 'List or search Outlook contacts.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string' },
              query: { type: 'string' },
              max_results: { type: 'number', default: 100 },
            },
            required: ['account'],
            additionalProperties: false,
          },
        },
        {
          name: 'create_contact',
          description: 'Create an Outlook contact.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string' },
              given_name: { type: 'string' },
              surname: { type: 'string' },
              display_name: { type: 'string' },
              email_addresses: { type: 'string' },
              mobile_phone: { type: 'string' },
              business_phones: { type: 'string' },
              company_name: { type: 'string' },
              job_title: { type: 'string' },
              notes: { type: 'string' },
            },
            required: ['account'],
            additionalProperties: false,
          },
        },
        {
          name: 'update_contact',
          description: 'Update an Outlook contact.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string' },
              contact_id: { type: 'string' },
              given_name: { type: 'string' },
              surname: { type: 'string' },
              display_name: { type: 'string' },
              email_addresses: { type: 'string' },
              mobile_phone: { type: 'string' },
              business_phones: { type: 'string' },
              company_name: { type: 'string' },
              job_title: { type: 'string' },
              notes: { type: 'string' },
            },
            required: ['account', 'contact_id'],
            additionalProperties: false,
          },
        },
        {
          name: 'delete_contact',
          description: 'Delete an Outlook contact (destructive). Requires confirm=true and reason.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string' },
              contact_id: { type: 'string' },
              confirm: { type: 'boolean' },
              reason: { type: 'string' },
            },
            required: ['account', 'contact_id', 'confirm', 'reason'],
            additionalProperties: false,
          },
        },
        {
          name: 'get_attachment',
          description:
            'Fetch a single Outlook email attachment by id. Saves to ~/Downloads/mcp-attachments/ and returns extracted text for supported formats (PDF, DOCX, XLSX, PPTX, text, images via OCR). Handles file, item, and reference attachment types.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Account id.' },
              email_id: { type: 'string', description: 'Graph message id.' },
              attachment_id: {
                type: 'string',
                description: 'Attachment id from read_emails / get_email_thread metadata.',
              },
            },
            required: ['account', 'email_id', 'attachment_id'],
            additionalProperties: false,
          },
        },
        {
          name: 'get_all_attachments',
          description:
            'Fetch every attachment on an Outlook email in one call. Saves each to disk and returns extracted text for each.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string', description: 'Account id.' },
              email_id: { type: 'string', description: 'Graph message id.' },
            },
            required: ['account', 'email_id'],
            additionalProperties: false,
          },
        },
        {
          name: 'search_emails',
          description: 'Full-text search across Outlook emails. Returns matching messages sorted by relevance.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string' },
              query: { type: 'string' },
              folder: { type: 'string' },
              max_results: { type: 'number', default: 20 },
              include_body: { type: 'boolean', default: false },
            },
            required: ['account', 'query'],
            additionalProperties: false,
          },
        },
        {
          name: 'list_drafts',
          description: 'List Outlook draft messages, sorted by last modified date.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string' },
              max_results: { type: 'number', default: 20 },
              skip: { type: 'number' },
            },
            required: ['account'],
            additionalProperties: false,
          },
        },
        {
          name: 'search_drafts',
          description: 'Search within Outlook drafts by keyword.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string' },
              query: { type: 'string' },
              max_results: { type: 'number', default: 20 },
            },
            required: ['account', 'query'],
            additionalProperties: false,
          },
        },
        {
          name: 'send_draft',
          description: 'Send an existing Outlook draft by message_id.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string' },
              message_id: { type: 'string' },
            },
            required: ['account', 'message_id'],
            additionalProperties: false,
          },
        },
        {
          name: 'delete_drafts',
          description: 'Delete one or more Outlook drafts by message_id. Requires confirm=true and reason.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string' },
              message_ids: { type: 'array', items: { type: 'string' } },
              confirm: { type: 'boolean' },
              reason: { type: 'string' },
            },
            required: ['account', 'message_ids', 'confirm', 'reason'],
            additionalProperties: false,
          },
        },
        {
          name: 'archive_emails',
          description: 'Move one or more Outlook emails to the Archive folder.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string' },
              message_ids: { type: 'array', items: { type: 'string' } },
            },
            required: ['account', 'message_ids'],
            additionalProperties: false,
          },
        },
        {
          name: 'trash_emails',
          description: 'Move one or more Outlook emails to the Deleted Items folder (soft delete).',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string' },
              message_ids: { type: 'array', items: { type: 'string' } },
            },
            required: ['account', 'message_ids'],
            additionalProperties: false,
          },
        },
        {
          name: 'mute_thread',
          description: 'Mute an Outlook conversation thread by moving all messages to Deleted Items. Provide message_id or conversation_id.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string' },
              message_id: { type: 'string' },
              conversation_id: { type: 'string' },
            },
            required: ['account'],
            additionalProperties: false,
          },
        },
        {
          name: 'unsubscribe_from_email',
          description: 'Unsubscribe from a mailing list using the List-Unsubscribe header in the email. Handles both HTTP and mailto unsubscribe methods.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string' },
              message_id: { type: 'string' },
            },
            required: ['account', 'message_id'],
            additionalProperties: false,
          },
        },
        {
          name: 'block_sender',
          description: "Add a sender to Outlook's blocked/junk list so their future emails are filtered to Other.",
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string' },
              sender_email: { type: 'string' },
            },
            required: ['account', 'sender_email'],
            additionalProperties: false,
          },
        },
        {
          name: 'unblock_sender',
          description: 'Remove a sender from the Outlook blocked list by override_id (from list_blocked_senders).',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string' },
              override_id: { type: 'string' },
            },
            required: ['account', 'override_id'],
            additionalProperties: false,
          },
        },
        {
          name: 'list_blocked_senders',
          description: "List all senders blocked via Outlook's Focused Inbox override list.",
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string' },
            },
            required: ['account'],
            additionalProperties: false,
          },
        },
        {
          name: 'get_event',
          description: 'Get a single Outlook calendar event by event_id, including full body and attendee details.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string' },
              event_id: { type: 'string' },
            },
            required: ['account', 'event_id'],
            additionalProperties: false,
          },
        },
        {
          name: 'list_categories',
          description: 'List all Outlook categories (labels) for an account.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string' },
            },
            required: ['account'],
            additionalProperties: false,
          },
        },
        {
          name: 'create_category',
          description: 'Create a new Outlook category (label). Color options: none, red, orange, yellow, green, teal, olive, blue, purple, cranberry, steel, darkSteel, darkBlue, darkPurple, darkTeal, darkGreen, darkYellow, darkOrange, darkRed, darkCranberry.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string' },
              display_name: { type: 'string' },
              color: { type: 'string' },
            },
            required: ['account', 'display_name'],
            additionalProperties: false,
          },
        },
        {
          name: 'delete_category',
          description: 'Delete an Outlook category by category_id (from list_categories). Requires confirm=true and reason.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string' },
              category_id: { type: 'string' },
              confirm: { type: 'boolean' },
              reason: { type: 'string' },
            },
            required: ['account', 'category_id', 'confirm', 'reason'],
            additionalProperties: false,
          },
        },
        {
          name: 'add_categories',
          description: 'Add one or more Outlook categories to a message (merges with existing categories).',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string' },
              message_id: { type: 'string' },
              categories: { type: 'array', items: { type: 'string' } },
            },
            required: ['account', 'message_id', 'categories'],
            additionalProperties: false,
          },
        },
        {
          name: 'remove_categories',
          description: 'Remove categories from an Outlook message. Omit categories to remove all.',
          inputSchema: {
            type: 'object',
            properties: {
              account: { type: 'string' },
              message_id: { type: 'string' },
              categories: { type: 'array', items: { type: 'string' } },
            },
            required: ['account', 'message_id'],
            additionalProperties: false,
          },
        },
      ],
    }));

    this.server.setRequestHandler(CallToolRequestSchema, async (request) => {
      const args = (request.params.arguments ?? {}) as Record<string, unknown>;

      try {
        switch (request.params.name) {
          case 'list_accounts': {
            const config = await loadAccountsConfig(this.configRoot);
            const health = await Promise.all(
              config.accounts.map((account) => getAccountHealth(this.configRoot, account)),
            );
            return textResult(formatJson({ defaultAccount: config.defaultAccount, accounts: health }));
          }

          case 'begin_account_auth': {
            await ensureConfigLayout(this.configRoot);
            const accountId = requireString(args.account_id, 'account_id');
            const email = requireString(args.email, 'email');
            const displayName = optionalString(args.display_name);
            validateAccountId(accountId);

            const credentialsInput = await readCredentialsInput(this.configRoot, accountId, {
              account_id: accountId,
              email,
              display_name: displayName,
              credentials_json: args.credentials_json,
              credentials_path: optionalString(args.credentials_path),
            });

            const config = await loadAccountsConfig(this.configRoot);
            const nextAccount: AccountConfig = {
              id: accountId,
              email,
              displayName,
              enabled: true,
              credentialPath: credentialsInput.credentialsPath,
              tokenCachePath: getDefaultAccountPaths(this.configRoot, accountId).tokenCachePath,
            };
            const nextConfig = upsertAccount(config, nextAccount);
            await saveAccountsConfig(this.configRoot, {
              ...nextConfig,
              defaultAccount: nextConfig.defaultAccount ?? accountId,
            });

            const paths = getAccountPaths(this.configRoot, nextAccount);
            await fs.mkdir(paths.accountDir, { recursive: true });
            const auth = await beginAuthFromCredentials(credentialsInput.credentials, paths.pendingAuthPath);

            return textResult(
              formatJson({
                account_id: accountId,
                email,
                auth_url: auth.authUrl,
                redirect_uri: auth.redirectUri,
                scopes: auth.scopes,
                next_step:
                  'Open auth_url, sign in to Microsoft, then pass the redirected full URL or the code value to finish_account_auth.',
              }),
            );
          }

          case 'finish_account_auth': {
            const accountId = requireString(args.account_id, 'account_id');
            const authorizationCode = requireString(args.authorization_code, 'authorization_code');
            const config = await loadAccountsConfig(this.configRoot);
            const account = resolveWriteAccount(config, accountId);
            const paths = getAccountPaths(this.configRoot, account);
            const credentials = await readCredentialsFile(paths.credentialsPath);
            const profile = await finishAuthFromCredentials(
              credentials,
              paths.pendingAuthPath,
              paths.tokenCachePath,
              authorizationCode,
            );

            const updated = upsertAccount(config, {
              ...account,
              email: profile.email || account.email,
              displayName: profile.displayName ?? account.displayName,
              enabled: true,
            });
            await saveAccountsConfig(this.configRoot, updated);

            return textResult(
              formatJson({
                account_id: accountId,
                email: profile.email,
                display_name: profile.displayName,
                ready: true,
              }),
            );
          }

          case 'list_mail_folders': {
            const config = await loadAccountsConfig(this.configRoot);
            const account = resolveWriteAccount(config, requireString(args.account, 'account'));
            const client = await OutlookAccountClient.create(this.configRoot, account);
            return textResult(formatJson(await client.listMailFolders()));
          }

          case 'read_emails': {
            const config = await loadAccountsConfig(this.configRoot);
            const accounts = resolveReadAccounts(config, optionalString(args.account));
            const folder = optionalString(args.folder);
            const query = optionalString(args.query);
            const maxResults = optionalNumber(args.max_results, 'max_results', 20);
            const includeBody = optionalBoolean(args.include_body, 'include_body') ?? false;

            const settled = await Promise.allSettled(
              accounts.map(async (account) => {
                const client = await OutlookAccountClient.create(this.configRoot, account);
                const result = await client.readEmails({
                  folder,
                  query,
                  maxResults,
                  includeBody,
                });
                return {
                  account: account.id,
                  email: account.email,
                  result,
                };
              }),
            );

            const results = settled
              .filter(
                (item): item is PromiseFulfilledResult<{ account: string; email: string; result: unknown }> =>
                  item.status === 'fulfilled',
              )
              .map((item) => item.value);

            const errors = settled
              .filter((item): item is PromiseRejectedResult => item.status === 'rejected')
              .map((item) => ({ error: item.reason instanceof Error ? item.reason.message : String(item.reason) }));

            return textResult(formatJson({ results, errors }));
          }

          case 'get_email_thread': {
            const config = await loadAccountsConfig(this.configRoot);
            const account = resolveWriteAccount(config, requireString(args.account, 'account'));
            const client = await OutlookAccountClient.create(this.configRoot, account);
            return textResult(
              formatJson(
                await client.getEmailThread({
                  messageId: optionalString(args.message_id),
                  conversationId: optionalString(args.conversation_id),
                }),
              ),
            );
          }

          case 'send_email': {
            const config = await loadAccountsConfig(this.configRoot);
            const account = resolveWriteAccount(config, requireString(args.account, 'account'));
            const client = await OutlookAccountClient.create(this.configRoot, account);
            return textResult(
              formatJson(
                await client.sendEmail({
                  to: requireString(args.to, 'to'),
                  subject: requireString(args.subject, 'subject'),
                  body: requireString(args.body, 'body'),
                  cc: optionalString(args.cc),
                  bcc: optionalString(args.bcc),
                  html: optionalBoolean(args.html, 'html') ?? false,
                }),
              ),
            );
          }

          case 'create_draft': {
            const config = await loadAccountsConfig(this.configRoot);
            const account = resolveWriteAccount(config, requireString(args.account, 'account'));
            const client = await OutlookAccountClient.create(this.configRoot, account);
            return textResult(
              formatJson(
                await client.createDraft({
                  to: requireString(args.to, 'to'),
                  subject: requireString(args.subject, 'subject'),
                  body: requireString(args.body, 'body'),
                  cc: optionalString(args.cc),
                  bcc: optionalString(args.bcc),
                  html: optionalBoolean(args.html, 'html') ?? false,
                }),
              ),
            );
          }

          case 'update_draft': {
            const config = await loadAccountsConfig(this.configRoot);
            const account = resolveWriteAccount(config, requireString(args.account, 'account'));
            const client = await OutlookAccountClient.create(this.configRoot, account);
            return textResult(
              formatJson(
                await client.updateDraft(requireString(args.message_id, 'message_id'), {
                  to: optionalString(args.to),
                  subject: optionalString(args.subject),
                  body: optionalString(args.body),
                  cc: optionalString(args.cc),
                  bcc: optionalString(args.bcc),
                  html: optionalBoolean(args.html, 'html'),
                }),
              ),
            );
          }

          case 'move_email': {
            const config = await loadAccountsConfig(this.configRoot);
            const account = resolveWriteAccount(config, requireString(args.account, 'account'));
            const client = await OutlookAccountClient.create(this.configRoot, account);
            return textResult(
              formatJson(
                await client.moveEmail(
                  requireString(args.message_id, 'message_id'),
                  requireString(args.destination_folder_id, 'destination_folder_id'),
                ),
              ),
            );
          }

          case 'mark_as_read': {
            const config = await loadAccountsConfig(this.configRoot);
            const account = resolveWriteAccount(config, requireString(args.account, 'account'));
            const client = await OutlookAccountClient.create(this.configRoot, account);
            return textResult(
              formatJson(
                await client.markAsRead(
                  requireStringArray(args.message_ids, 'message_ids'),
                  optionalBoolean(args.is_read, 'is_read') ?? true,
                ),
              ),
            );
          }

          case 'delete_email': {
            requireConfirm(args.confirm, args.reason);
            const config = await loadAccountsConfig(this.configRoot);
            const account = resolveWriteAccount(config, requireString(args.account, 'account'));
            const client = await OutlookAccountClient.create(this.configRoot, account);
            return textResult(
              formatJson(await client.deleteEmail(requireString(args.message_id, 'message_id'))),
            );
          }

          case 'list_calendars': {
            const config = await loadAccountsConfig(this.configRoot);
            const account = resolveWriteAccount(config, requireString(args.account, 'account'));
            const client = await OutlookAccountClient.create(this.configRoot, account);
            return textResult(formatJson(await client.listCalendars()));
          }

          case 'list_events': {
            const config = await loadAccountsConfig(this.configRoot);
            const account = resolveWriteAccount(config, requireString(args.account, 'account'));
            const client = await OutlookAccountClient.create(this.configRoot, account);
            return textResult(
              formatJson(
                await client.listEvents({
                  calendarId: optionalString(args.calendar_id),
                  start: optionalString(args.start) ?? defaultRangeStart(),
                  end: optionalString(args.end) ?? defaultRangeEnd(),
                  maxResults: optionalNumber(args.max_results, 'max_results', 50),
                }),
              ),
            );
          }

          case 'create_event': {
            const config = await loadAccountsConfig(this.configRoot);
            const account = resolveWriteAccount(config, requireString(args.account, 'account'));
            const client = await OutlookAccountClient.create(this.configRoot, account);
            return textResult(
              formatJson(
                await client.createEvent({
                  calendarId: optionalString(args.calendar_id),
                  subject: requireString(args.subject, 'subject'),
                  start: requireString(args.start, 'start'),
                  end: requireString(args.end, 'end'),
                  is_all_day: optionalBoolean(args.is_all_day, 'is_all_day'),
                  location: optionalString(args.location),
                  body: optionalString(args.body),
                  attendees: optionalString(args.attendees),
                }),
              ),
            );
          }

          case 'update_event': {
            const config = await loadAccountsConfig(this.configRoot);
            const account = resolveWriteAccount(config, requireString(args.account, 'account'));
            const client = await OutlookAccountClient.create(this.configRoot, account);
            return textResult(
              formatJson(
                await client.updateEvent(requireString(args.event_id, 'event_id'), {
                  subject: optionalString(args.subject),
                  start: optionalString(args.start),
                  end: optionalString(args.end),
                  is_all_day: optionalBoolean(args.is_all_day, 'is_all_day'),
                  location: optionalString(args.location),
                  body: optionalString(args.body),
                  attendees: optionalString(args.attendees),
                }),
              ),
            );
          }

          case 'delete_event': {
            requireConfirm(args.confirm, args.reason);
            const config = await loadAccountsConfig(this.configRoot);
            const account = resolveWriteAccount(config, requireString(args.account, 'account'));
            const client = await OutlookAccountClient.create(this.configRoot, account);
            return textResult(
              formatJson(await client.deleteEvent(requireString(args.event_id, 'event_id'))),
            );
          }

          case 'list_contacts': {
            const config = await loadAccountsConfig(this.configRoot);
            const account = resolveWriteAccount(config, requireString(args.account, 'account'));
            const client = await OutlookAccountClient.create(this.configRoot, account);
            return textResult(
              formatJson(
                await client.listContacts({
                  query: optionalString(args.query),
                  maxResults: optionalNumber(args.max_results, 'max_results', 100),
                }),
              ),
            );
          }

          case 'create_contact': {
            const config = await loadAccountsConfig(this.configRoot);
            const account = resolveWriteAccount(config, requireString(args.account, 'account'));
            const client = await OutlookAccountClient.create(this.configRoot, account);
            return textResult(
              formatJson(
                await client.createContact({
                  given_name: optionalString(args.given_name),
                  surname: optionalString(args.surname),
                  display_name: optionalString(args.display_name),
                  email_addresses: optionalString(args.email_addresses),
                  mobile_phone: optionalString(args.mobile_phone),
                  business_phones: optionalString(args.business_phones),
                  company_name: optionalString(args.company_name),
                  job_title: optionalString(args.job_title),
                  notes: optionalString(args.notes),
                }),
              ),
            );
          }

          case 'update_contact': {
            const config = await loadAccountsConfig(this.configRoot);
            const account = resolveWriteAccount(config, requireString(args.account, 'account'));
            const client = await OutlookAccountClient.create(this.configRoot, account);
            return textResult(
              formatJson(
                await client.updateContact(requireString(args.contact_id, 'contact_id'), {
                  given_name: optionalString(args.given_name),
                  surname: optionalString(args.surname),
                  display_name: optionalString(args.display_name),
                  email_addresses: optionalString(args.email_addresses),
                  mobile_phone: optionalString(args.mobile_phone),
                  business_phones: optionalString(args.business_phones),
                  company_name: optionalString(args.company_name),
                  job_title: optionalString(args.job_title),
                  notes: optionalString(args.notes),
                }),
              ),
            );
          }

          case 'delete_contact': {
            requireConfirm(args.confirm, args.reason);
            const config = await loadAccountsConfig(this.configRoot);
            const account = resolveWriteAccount(config, requireString(args.account, 'account'));
            const client = await OutlookAccountClient.create(this.configRoot, account);
            return textResult(
              formatJson(await client.deleteContact(requireString(args.contact_id, 'contact_id'))),
            );
          }

          case 'get_attachment': {
            const parsed: GetAttachmentArgs = {
              account: requireString(args.account, 'account'),
              email_id: requireString(args.email_id, 'email_id'),
              attachment_id: requireString(args.attachment_id, 'attachment_id'),
            };

            const config = await loadAccountsConfig(this.configRoot);
            const account = resolveWriteAccount(config, parsed.account);
            const client = await OutlookAccountClient.create(this.configRoot, account);

            const fetched = await client.getAttachment(parsed.email_id, parsed.attachment_id);

            if (fetched.bytes === null) {
              return textResult(
                formatJson({
                  account: account.id,
                  email: account.email,
                  email_id: parsed.email_id,
                  attachment_id: parsed.attachment_id,
                  attachment_type: fetched.attachmentType,
                  reference_url: fetched.referenceUrl,
                  filename: fetched.metadata.filename,
                  content_type: fetched.metadata.contentType,
                  size_bytes: fetched.metadata.sizeBytes,
                  saved_path: null,
                  text: null,
                  extraction_method: 'none',
                  extraction_error: fetched.note,
                }),
              );
            }

            const content: AttachmentContent = await saveAndExtract(
              fetched.bytes,
              fetched.metadata,
            );
            content.attachmentType = fetched.attachmentType;
            content.referenceUrl = fetched.referenceUrl;

            return textResult(
              formatJson({
                account: account.id,
                email: account.email,
                email_id: parsed.email_id,
                attachment_id: parsed.attachment_id,
                attachment: content,
              }),
            );
          }

          case 'get_all_attachments': {
            const parsed: GetAllAttachmentsArgs = {
              account: requireString(args.account, 'account'),
              email_id: requireString(args.email_id, 'email_id'),
            };

            const config = await loadAccountsConfig(this.configRoot);
            const account = resolveWriteAccount(config, parsed.account);
            const client = await OutlookAccountClient.create(this.configRoot, account);

            const list: AttachmentMetadata[] = await client.listAttachments(parsed.email_id);

            const settled = await Promise.allSettled(
              list.map(async (meta) => {
                const fetched = await client.getAttachment(parsed.email_id, meta.id);
                if (fetched.bytes === null) {
                  const placeholder: AttachmentContent = {
                    ...fetched.metadata,
                    savedPath: '',
                    text: null,
                    extractionMethod: 'none',
                    extractionError: fetched.note,
                    attachmentType: fetched.attachmentType,
                    referenceUrl: fetched.referenceUrl,
                  };
                  return placeholder;
                }
                const content = await saveAndExtract(fetched.bytes, fetched.metadata);
                content.attachmentType = fetched.attachmentType;
                content.referenceUrl = fetched.referenceUrl;
                return content;
              }),
            );

            const attachments: AttachmentContent[] = [];
            const errors: Array<{ attachment_id: string; error: string }> = [];

            settled.forEach((result, index) => {
              const originalId = list[index]?.id ?? '(unknown)';
              if (result.status === 'fulfilled') {
                attachments.push(result.value);
              } else {
                errors.push({
                  attachment_id: originalId,
                  error:
                    result.reason instanceof Error
                      ? result.reason.message
                      : String(result.reason),
                });
              }
            });

            return textResult(
              formatJson({
                account: account.id,
                email: account.email,
                email_id: parsed.email_id,
                count: attachments.length,
                attachments,
                errors,
              }),
            );
          }

          case 'search_emails': {
            const config = await loadAccountsConfig(this.configRoot);
            const account = resolveWriteAccount(config, requireString(args.account, 'account'));
            const client = await OutlookAccountClient.create(this.configRoot, account);
            return textResult(
              formatJson(
                await client.searchEmails({
                  query: requireString(args.query, 'query'),
                  folder: optionalString(args.folder),
                  maxResults: optionalNumber(args.max_results, 'max_results', 20),
                  includeBody: optionalBoolean(args.include_body, 'include_body') ?? false,
                }),
              ),
            );
          }

          case 'list_drafts': {
            const config = await loadAccountsConfig(this.configRoot);
            const account = resolveWriteAccount(config, requireString(args.account, 'account'));
            const client = await OutlookAccountClient.create(this.configRoot, account);
            return textResult(
              formatJson(
                await client.listDrafts({
                  maxResults: optionalNumber(args.max_results, 'max_results', 20),
                  skip: args.skip !== undefined ? optionalNumber(args.skip, 'skip', 0) : undefined,
                }),
              ),
            );
          }

          case 'search_drafts': {
            const config = await loadAccountsConfig(this.configRoot);
            const account = resolveWriteAccount(config, requireString(args.account, 'account'));
            const client = await OutlookAccountClient.create(this.configRoot, account);
            return textResult(
              formatJson(
                await client.searchEmails({
                  query: requireString(args.query, 'query'),
                  folder: 'drafts',
                  maxResults: optionalNumber(args.max_results, 'max_results', 20),
                  includeBody: false,
                }),
              ),
            );
          }

          case 'send_draft': {
            const config = await loadAccountsConfig(this.configRoot);
            const account = resolveWriteAccount(config, requireString(args.account, 'account'));
            const client = await OutlookAccountClient.create(this.configRoot, account);
            return textResult(
              formatJson(await client.sendDraft(requireString(args.message_id, 'message_id'))),
            );
          }

          case 'delete_drafts': {
            requireConfirm(args.confirm, args.reason);
            const config = await loadAccountsConfig(this.configRoot);
            const account = resolveWriteAccount(config, requireString(args.account, 'account'));
            const client = await OutlookAccountClient.create(this.configRoot, account);
            return textResult(
              formatJson(await client.deleteDrafts(requireStringArray(args.message_ids, 'message_ids'))),
            );
          }

          case 'archive_emails': {
            const config = await loadAccountsConfig(this.configRoot);
            const account = resolveWriteAccount(config, requireString(args.account, 'account'));
            const client = await OutlookAccountClient.create(this.configRoot, account);
            return textResult(
              formatJson(await client.archiveEmails(requireStringArray(args.message_ids, 'message_ids'))),
            );
          }

          case 'trash_emails': {
            const config = await loadAccountsConfig(this.configRoot);
            const account = resolveWriteAccount(config, requireString(args.account, 'account'));
            const client = await OutlookAccountClient.create(this.configRoot, account);
            return textResult(
              formatJson(await client.trashEmails(requireStringArray(args.message_ids, 'message_ids'))),
            );
          }

          case 'mute_thread': {
            const config = await loadAccountsConfig(this.configRoot);
            const account = resolveWriteAccount(config, requireString(args.account, 'account'));
            const client = await OutlookAccountClient.create(this.configRoot, account);
            return textResult(
              formatJson(
                await client.muteThread({
                  messageId: optionalString(args.message_id),
                  conversationId: optionalString(args.conversation_id),
                }),
              ),
            );
          }

          case 'unsubscribe_from_email': {
            const config = await loadAccountsConfig(this.configRoot);
            const account = resolveWriteAccount(config, requireString(args.account, 'account'));
            const client = await OutlookAccountClient.create(this.configRoot, account);
            return textResult(
              formatJson(
                await client.unsubscribeFromEmail(requireString(args.message_id, 'message_id')),
              ),
            );
          }

          case 'block_sender': {
            const config = await loadAccountsConfig(this.configRoot);
            const account = resolveWriteAccount(config, requireString(args.account, 'account'));
            const client = await OutlookAccountClient.create(this.configRoot, account);
            return textResult(
              formatJson(await client.blockSender(requireString(args.sender_email, 'sender_email'))),
            );
          }

          case 'unblock_sender': {
            const config = await loadAccountsConfig(this.configRoot);
            const account = resolveWriteAccount(config, requireString(args.account, 'account'));
            const client = await OutlookAccountClient.create(this.configRoot, account);
            return textResult(
              formatJson(await client.unblockSender(requireString(args.override_id, 'override_id'))),
            );
          }

          case 'list_blocked_senders': {
            const config = await loadAccountsConfig(this.configRoot);
            const account = resolveWriteAccount(config, requireString(args.account, 'account'));
            const client = await OutlookAccountClient.create(this.configRoot, account);
            return textResult(formatJson(await client.listBlockedSenders()));
          }

          case 'get_event': {
            const config = await loadAccountsConfig(this.configRoot);
            const account = resolveWriteAccount(config, requireString(args.account, 'account'));
            const client = await OutlookAccountClient.create(this.configRoot, account);
            return textResult(
              formatJson(await client.getEvent(requireString(args.event_id, 'event_id'))),
            );
          }

          case 'list_categories': {
            const config = await loadAccountsConfig(this.configRoot);
            const account = resolveWriteAccount(config, requireString(args.account, 'account'));
            const client = await OutlookAccountClient.create(this.configRoot, account);
            return textResult(formatJson(await client.listCategories()));
          }

          case 'create_category': {
            const config = await loadAccountsConfig(this.configRoot);
            const account = resolveWriteAccount(config, requireString(args.account, 'account'));
            const client = await OutlookAccountClient.create(this.configRoot, account);
            return textResult(
              formatJson(
                await client.createCategory(
                  requireString(args.display_name, 'display_name'),
                  optionalString(args.color),
                ),
              ),
            );
          }

          case 'delete_category': {
            requireConfirm(args.confirm, args.reason);
            const config = await loadAccountsConfig(this.configRoot);
            const account = resolveWriteAccount(config, requireString(args.account, 'account'));
            const client = await OutlookAccountClient.create(this.configRoot, account);
            return textResult(
              formatJson(await client.deleteCategory(requireString(args.category_id, 'category_id'))),
            );
          }

          case 'add_categories': {
            const config = await loadAccountsConfig(this.configRoot);
            const account = resolveWriteAccount(config, requireString(args.account, 'account'));
            const client = await OutlookAccountClient.create(this.configRoot, account);
            return textResult(
              formatJson(
                await client.addCategories(
                  requireString(args.message_id, 'message_id'),
                  requireStringArray(args.categories, 'categories'),
                ),
              ),
            );
          }

          case 'remove_categories': {
            const config = await loadAccountsConfig(this.configRoot);
            const account = resolveWriteAccount(config, requireString(args.account, 'account'));
            const client = await OutlookAccountClient.create(this.configRoot, account);
            const categories = Array.isArray(args.categories)
              ? (args.categories as string[]).map(String).filter(Boolean)
              : undefined;
            return textResult(
              formatJson(
                await client.removeCategories(
                  requireString(args.message_id, 'message_id'),
                  categories,
                ),
              ),
            );
          }

          default:
            throw new Error(`Unknown tool: ${request.params.name}`);
        }
      } catch (error) {
        return textResult(
          formatJson({
            error: (error as Error).message,
          }),
        );
      }
    });
  }

  async start(): Promise<void> {
    await ensureConfigLayout(this.configRoot);
    const transport = new StdioServerTransport();
    await this.server.connect(transport);
  }
}

const app = new OutlookMcpServer();
await app.start();
