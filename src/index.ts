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
  type OutlookCredentials,
} from './graph-client.js';

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
