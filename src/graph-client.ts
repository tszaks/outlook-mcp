import { promises as fs } from 'node:fs';
import crypto from 'node:crypto';
import {
  PublicClientApplication,
  type AccountInfo,
  type AuthorizationUrlRequest,
  type SilentFlowRequest,
} from '@azure/msal-node';
import { AccountConfig, getAccountPaths, type AccountPaths } from './config.js';

export const OUTLOOK_SCOPES = [
  'openid',
  'profile',
  'offline_access',
  'User.Read',
  'Mail.ReadWrite',
  'Mail.Send',
  'Calendars.ReadWrite',
  'Contacts.ReadWrite',
] as const;

export interface OutlookCredentials {
  clientId: string;
  tenantId?: string;
  redirectUri?: string;
  scopes?: string[];
}

interface PendingAuthState {
  state: string;
  codeVerifier: string;
  redirectUri: string;
  scopes: string[];
  createdAt: string;
}

interface Recipient {
  emailAddress: {
    address: string;
    name?: string;
  };
}

interface BodyContent {
  contentType: 'Text' | 'HTML';
  content: string;
}

function getAuthority(tenantId?: string): string {
  return `https://login.microsoftonline.com/${tenantId?.trim() || 'common'}`;
}

function normalizeScopes(credentials: OutlookCredentials): string[] {
  const customScopes = Array.isArray(credentials.scopes)
    ? credentials.scopes.map((scope) => String(scope).trim()).filter(Boolean)
    : [];
  return customScopes.length > 0 ? customScopes : [...OUTLOOK_SCOPES];
}

function buildCredentialsObject(input: unknown): OutlookCredentials {
  if (!input || typeof input !== 'object') {
    throw new Error('Invalid credentials content.');
  }

  const candidate = input as Partial<OutlookCredentials>;
  if (typeof candidate.clientId !== 'string' || !candidate.clientId.trim()) {
    throw new Error('Credentials must include "clientId".');
  }

  return {
    clientId: candidate.clientId.trim(),
    tenantId:
      typeof candidate.tenantId === 'string' && candidate.tenantId.trim()
        ? candidate.tenantId.trim()
        : undefined,
    redirectUri:
      typeof candidate.redirectUri === 'string' && candidate.redirectUri.trim()
        ? candidate.redirectUri.trim()
        : 'http://localhost',
    scopes: normalizeScopes(candidate as OutlookCredentials),
  };
}

function createMsalClient(credentials: OutlookCredentials): PublicClientApplication {
  return new PublicClientApplication({
    auth: {
      clientId: credentials.clientId,
      authority: getAuthority(credentials.tenantId),
    },
  });
}

async function loadJsonFile<T>(filePath: string): Promise<T> {
  const raw = await fs.readFile(filePath, 'utf8');
  return JSON.parse(raw) as T;
}

async function saveJsonFile(filePath: string, value: unknown): Promise<void> {
  await fs.writeFile(filePath, `${JSON.stringify(value, null, 2)}\n`, 'utf8');
}

function randomBase64Url(bytes: number): string {
  return crypto
    .randomBytes(bytes)
    .toString('base64')
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/g, '');
}

function sha256Base64Url(input: string): string {
  return crypto
    .createHash('sha256')
    .update(input)
    .digest('base64')
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/g, '');
}

export function parseAuthorizationInput(input: string): string {
  const trimmed = input.trim();
  if (!trimmed) {
    throw new Error('authorization_code is required.');
  }

  if (trimmed.startsWith('http://') || trimmed.startsWith('https://')) {
    const url = new URL(trimmed);
    const code = url.searchParams.get('code');
    if (!code) {
      throw new Error('The redirect URL does not contain a code query parameter.');
    }
    return code;
  }

  return trimmed;
}

function parseAddresses(value?: string): Recipient[] {
  if (!value || value.trim() === '') return [];

  return value
    .split(',')
    .map((item) => item.trim())
    .filter(Boolean)
    .map((address) => ({
      emailAddress: {
        address,
      },
    }));
}

function buildMessagePayload(input: {
  subject?: string;
  body?: string;
  to?: string;
  cc?: string;
  bcc?: string;
  html?: boolean;
}): Record<string, unknown> {
  const payload: Record<string, unknown> = {};

  if (typeof input.subject === 'string') {
    payload.subject = input.subject;
  }

  if (typeof input.body === 'string') {
    payload.body = {
      contentType: input.html ? 'HTML' : 'Text',
      content: input.body,
    } satisfies BodyContent;
  }

  if (input.to !== undefined) {
    payload.toRecipients = parseAddresses(input.to);
  }

  if (input.cc !== undefined) {
    payload.ccRecipients = parseAddresses(input.cc);
  }

  if (input.bcc !== undefined) {
    payload.bccRecipients = parseAddresses(input.bcc);
  }

  return payload;
}

function toGraphDateTimeParts(input: string): { dateTime: string; timeZone: string } {
  const parsed = new Date(input);
  if (Number.isNaN(parsed.getTime())) {
    throw new Error(`Invalid ISO date string: ${input}`);
  }

  return {
    dateTime: parsed.toISOString().replace('Z', ''),
    timeZone: 'UTC',
  };
}

function toGraphAllDayDateTimeParts(input: string): { dateTime: string; timeZone: string } {
  const datePortion = extractDatePortion(input);

  return {
    dateTime: `${datePortion}T00:00:00.0000000`,
    timeZone: 'UTC',
  };
}

function extractDatePortion(input: string): string {
  const trimmed = input.trim();
  const match = trimmed.match(/^(\d{4}-\d{2}-\d{2})/);
  if (match) {
    return match[1];
  }

  const parsed = new Date(trimmed);
  if (Number.isNaN(parsed.getTime())) {
    throw new Error(`Invalid ISO date string: ${input}`);
  }

  return parsed.toISOString().slice(0, 10);
}

function addDays(datePortion: string, days: number): string {
  const parsed = new Date(`${datePortion}T00:00:00.000Z`);
  if (Number.isNaN(parsed.getTime())) {
    throw new Error(`Invalid date value: ${datePortion}`);
  }
  parsed.setUTCDate(parsed.getUTCDate() + days);
  return parsed.toISOString().slice(0, 10);
}

function buildEventPayload(input: {
  subject?: string;
  body?: string;
  start?: string;
  end?: string;
  is_all_day?: boolean;
  location?: string;
  attendees?: string;
}): Record<string, unknown> {
  const payload: Record<string, unknown> = {};

  if (typeof input.subject === 'string') {
    payload.subject = input.subject;
  }

  if (typeof input.body === 'string') {
    payload.body = {
      contentType: 'HTML',
      content: input.body,
    } satisfies BodyContent;
  }

  if (typeof input.location === 'string') {
    payload.location = {
      displayName: input.location,
    };
  }

  if (typeof input.attendees === 'string') {
    payload.attendees = parseAddresses(input.attendees).map((attendee) => ({
      emailAddress: attendee.emailAddress,
      type: 'required',
    }));
  }

  if (input.is_all_day !== undefined) {
    payload.isAllDay = input.is_all_day;
  }

  if (input.is_all_day) {
    const startDatePortion = typeof input.start === 'string' ? extractDatePortion(input.start) : undefined;
    const endDatePortion = typeof input.end === 'string' ? extractDatePortion(input.end) : undefined;

    if (startDatePortion) {
      payload.start = toGraphAllDayDateTimeParts(startDatePortion);
    }

    if (endDatePortion) {
      const normalizedEnd = endDatePortion <= (startDatePortion ?? endDatePortion)
        ? addDays(startDatePortion ?? endDatePortion, 1)
        : addDays(endDatePortion, 1);
      payload.end = toGraphAllDayDateTimeParts(normalizedEnd);
    }
  } else {
    if (typeof input.start === 'string') {
      payload.start = toGraphDateTimeParts(input.start);
    }

    if (typeof input.end === 'string') {
      payload.end = toGraphDateTimeParts(input.end);
    }
  }

  return payload;
}

function buildContactPayload(input: {
  given_name?: string;
  surname?: string;
  display_name?: string;
  email_addresses?: string;
  mobile_phone?: string;
  business_phones?: string;
  company_name?: string;
  job_title?: string;
  notes?: string;
}): Record<string, unknown> {
  const payload: Record<string, unknown> = {};

  if (typeof input.given_name === 'string') payload.givenName = input.given_name;
  if (typeof input.surname === 'string') payload.surname = input.surname;
  if (typeof input.display_name === 'string') payload.displayName = input.display_name;
  if (typeof input.mobile_phone === 'string') payload.mobilePhone = input.mobile_phone;
  if (typeof input.company_name === 'string') payload.companyName = input.company_name;
  if (typeof input.job_title === 'string') payload.jobTitle = input.job_title;
  if (typeof input.notes === 'string') payload.personalNotes = input.notes;

  if (typeof input.email_addresses === 'string') {
    payload.emailAddresses = input.email_addresses
      .split(',')
      .map((item) => item.trim())
      .filter(Boolean)
      .map((address) => ({ address, name: address }));
  }

  if (typeof input.business_phones === 'string') {
    payload.businessPhones = input.business_phones
      .split(',')
      .map((item) => item.trim())
      .filter(Boolean);
  }

  return payload;
}

export async function readCredentialsFile(credentialsPath: string): Promise<OutlookCredentials> {
  const raw = await fs.readFile(credentialsPath, 'utf8');
  return buildCredentialsObject(JSON.parse(raw));
}

export async function beginAuthFromCredentials(
  credentials: OutlookCredentials,
  pendingAuthPath: string,
): Promise<{ authUrl: string; redirectUri: string; scopes: string[] }> {
  const client = createMsalClient(credentials);
  const codeVerifier = randomBase64Url(32);
  const state = randomBase64Url(16);
  const redirectUri = credentials.redirectUri ?? 'http://localhost';
  const scopes = normalizeScopes(credentials);

  const authUrlRequest: AuthorizationUrlRequest = {
    scopes,
    redirectUri,
    codeChallenge: sha256Base64Url(codeVerifier),
    codeChallengeMethod: 'S256',
    prompt: 'select_account',
    state,
  };

  const authUrl = await client.getAuthCodeUrl(authUrlRequest);

  await saveJsonFile(pendingAuthPath, {
    state,
    codeVerifier,
    redirectUri,
    scopes,
    createdAt: new Date().toISOString(),
  } satisfies PendingAuthState);

  return { authUrl, redirectUri, scopes };
}

async function loadPendingAuth(pendingAuthPath: string): Promise<PendingAuthState> {
  try {
    return await loadJsonFile<PendingAuthState>(pendingAuthPath);
  } catch (error) {
    throw new Error(
      `Pending auth state not found at ${pendingAuthPath}. Run begin_account_auth first: ${(error as Error).message}`,
    );
  }
}

export async function finishAuthFromCredentials(
  credentials: OutlookCredentials,
  pendingAuthPath: string,
  tokenCachePath: string,
  authorizationInput: string,
): Promise<{ email: string; displayName?: string }> {
  const pending = await loadPendingAuth(pendingAuthPath);
  const client = createMsalClient(credentials);

  try {
    const cacheRaw = await fs.readFile(tokenCachePath, 'utf8');
    await client.getTokenCache().deserialize(cacheRaw);
  } catch {
    // Fresh cache is fine during first auth.
  }

  const parsed = authorizationInput.trim().startsWith('http')
    ? new URL(authorizationInput.trim())
    : null;

  if (parsed) {
    const returnedState = parsed.searchParams.get('state');
    if (returnedState && returnedState !== pending.state) {
      throw new Error('OAuth state mismatch. Start the auth flow again.');
    }
  }

  const code = parseAuthorizationInput(authorizationInput);
  const result = await client.acquireTokenByCode({
    code,
    scopes: pending.scopes,
    redirectUri: pending.redirectUri,
    codeVerifier: pending.codeVerifier,
  });

  if (!result?.account) {
    throw new Error('Microsoft login succeeded, but no account information was returned.');
  }

  await fs.writeFile(tokenCachePath, await client.getTokenCache().serialize(), 'utf8');
  await fs.rm(pendingAuthPath, { force: true });

  return {
    email: result.account.username,
    displayName: result.account.name ?? undefined,
  };
}

function escapeFilterValue(value: string): string {
  return value.replace(/'/g, "''");
}

async function ensureCacheLoaded(
  client: PublicClientApplication,
  tokenCachePath: string,
): Promise<void> {
  try {
    const raw = await fs.readFile(tokenCachePath, 'utf8');
    await client.getTokenCache().deserialize(raw);
  } catch (error) {
    throw new Error(
      `Token cache missing or invalid at ${tokenCachePath}. Reconnect the account: ${(error as Error).message}`,
    );
  }
}

async function getCachedAccount(
  client: PublicClientApplication,
  account: AccountConfig,
): Promise<AccountInfo> {
  const accounts = await client.getTokenCache().getAllAccounts();
  const match = accounts.find((item) => item.username.toLowerCase() === account.email.toLowerCase());
  if (!match) {
    throw new Error(
      `No cached Microsoft account matched "${account.email}" for "${account.id}". Reconnect the correct account.`,
    );
  }
  return match;
}

export class OutlookAccountClient {
  readonly account: AccountConfig;
  readonly paths: AccountPaths;
  private readonly credentials: OutlookCredentials;
  private readonly client: PublicClientApplication;

  private constructor(
    account: AccountConfig,
    paths: AccountPaths,
    credentials: OutlookCredentials,
    client: PublicClientApplication,
  ) {
    this.account = account;
    this.paths = paths;
    this.credentials = credentials;
    this.client = client;
  }

  static async create(configRoot: string, account: AccountConfig): Promise<OutlookAccountClient> {
    const paths = getAccountPaths(configRoot, account);
    const credentials = await readCredentialsFile(paths.credentialsPath);
    const client = createMsalClient(credentials);
    await ensureCacheLoaded(client, paths.tokenCachePath);
    return new OutlookAccountClient(account, paths, credentials, client);
  }

  private async acquireAccessToken(): Promise<string> {
    const cachedAccount = await getCachedAccount(this.client, this.account);
    const silentRequest: SilentFlowRequest = {
      account: cachedAccount,
      scopes: normalizeScopes(this.credentials),
    };

    const token = await this.client.acquireTokenSilent(silentRequest);
    if (!token?.accessToken) {
      throw new Error(`Failed to acquire an Outlook access token for "${this.account.id}".`);
    }

    await fs.writeFile(this.paths.tokenCachePath, await this.client.getTokenCache().serialize(), 'utf8');
    return token.accessToken;
  }

  private async graphRequest<T>(
    path: string,
    options?: {
      method?: string;
      headers?: Record<string, string>;
      body?: unknown;
    },
  ): Promise<T> {
    const token = await this.acquireAccessToken();
    const response = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
      method: options?.method ?? 'GET',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
        ...options?.headers,
      },
      body: options?.body === undefined ? undefined : JSON.stringify(options.body),
    });

    if (!response.ok) {
      const text = await response.text();
      throw new Error(`Graph request failed (${response.status} ${response.statusText}): ${text}`);
    }

    if (response.status === 204) {
      return null as T;
    }

    const contentLength = response.headers.get('content-length');
    if (contentLength === '0') {
      return null as T;
    }

    const contentType = response.headers.get('content-type') ?? '';
    if (!contentType.toLowerCase().includes('application/json')) {
      const text = await response.text();
      if (!text.trim()) {
        return null as T;
      }
      throw new Error(`Graph response was not JSON: ${text}`);
    }

    return (await response.json()) as T;
  }

  async getProfile(): Promise<unknown> {
    return this.graphRequest('/me?$select=id,displayName,userPrincipalName,mail');
  }

  async listMailFolders(): Promise<unknown> {
    return this.graphRequest('/me/mailFolders?$top=100&$select=id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount');
  }

  async readEmails(input: {
    folder?: string;
    query?: string;
    maxResults: number;
    includeBody: boolean;
  }): Promise<unknown> {
    const top = Math.max(1, Math.min(input.maxResults, 100));
    const basePath = input.folder
      ? `/me/mailFolders/${encodeURIComponent(input.folder)}/messages`
      : '/me/messages';
    const selectFields = [
      'id',
      'conversationId',
      'internetMessageId',
      'subject',
      'receivedDateTime',
      'sentDateTime',
      'from',
      'toRecipients',
      'ccRecipients',
      'bccRecipients',
      'isRead',
      'bodyPreview',
      'parentFolderId',
    ];

    if (input.includeBody) {
      selectFields.push('body');
    }

    const params = new URLSearchParams({
      $top: String(top),
      $select: selectFields.join(','),
    });

    if (input.query && input.query.trim()) {
      params.set('$search', `"${input.query.trim().replace(/"/g, '\\"')}"`);
    } else {
      params.set('$orderby', 'receivedDateTime DESC');
    }

    return this.graphRequest(`${basePath}?${params.toString()}`, input.query?.trim()
      ? { headers: { ConsistencyLevel: 'eventual' } }
      : undefined);
  }

  async getEmailThread(input: { messageId?: string; conversationId?: string }): Promise<unknown> {
    let conversationId = input.conversationId?.trim();

    if (!conversationId && input.messageId?.trim()) {
      const message = await this.graphRequest<{ conversationId?: string }>(
        `/me/messages/${encodeURIComponent(input.messageId.trim())}?$select=conversationId`,
      );
      conversationId = message.conversationId?.trim();
    }

    if (!conversationId) {
      throw new Error('Provide either message_id or conversation_id.');
    }

    const params = new URLSearchParams({
      $filter: `conversationId eq '${escapeFilterValue(conversationId)}'`,
      $orderby: 'receivedDateTime ASC',
      $top: '100',
      $select:
        'id,conversationId,subject,receivedDateTime,from,toRecipients,ccRecipients,bccRecipients,isRead,bodyPreview,body,parentFolderId',
    });

    return this.graphRequest(`/me/messages?${params.toString()}`);
  }

  async sendEmail(input: {
    to: string;
    subject: string;
    body: string;
    cc?: string;
    bcc?: string;
    html?: boolean;
  }): Promise<{ success: boolean }> {
    await this.graphRequest('/me/sendMail', {
      method: 'POST',
      body: {
        message: buildMessagePayload(input),
        saveToSentItems: true,
      },
    });
    return { success: true };
  }

  async createDraft(input: {
    to: string;
    subject: string;
    body: string;
    cc?: string;
    bcc?: string;
    html?: boolean;
  }): Promise<unknown> {
    return this.graphRequest('/me/messages', {
      method: 'POST',
      body: buildMessagePayload(input),
    });
  }

  async updateDraft(
    messageId: string,
    input: {
      to?: string;
      subject?: string;
      body?: string;
      cc?: string;
      bcc?: string;
      html?: boolean;
    },
  ): Promise<unknown> {
    return this.graphRequest(`/me/messages/${encodeURIComponent(messageId)}`, {
      method: 'PATCH',
      body: buildMessagePayload(input),
    });
  }

  async deleteEmail(messageId: string): Promise<{ success: boolean }> {
    await this.graphRequest(`/me/messages/${encodeURIComponent(messageId)}`, {
      method: 'DELETE',
    });
    return { success: true };
  }

  async moveEmail(messageId: string, destinationFolderId: string): Promise<unknown> {
    return this.graphRequest(`/me/messages/${encodeURIComponent(messageId)}/move`, {
      method: 'POST',
      body: {
        destinationId: destinationFolderId,
      },
    });
  }

  async markAsRead(messageIds: string[], isRead: boolean): Promise<unknown[]> {
    const uniqueIds = Array.from(
      new Set(messageIds.map((messageId) => messageId.trim()).filter(Boolean)),
    );
    return Promise.all(
      uniqueIds.map((messageId) =>
        this.graphRequest(`/me/messages/${encodeURIComponent(messageId)}`, {
          method: 'PATCH',
          body: { isRead },
        }),
      ),
    );
  }

  async listCalendars(): Promise<unknown> {
    return this.graphRequest('/me/calendars?$top=100&$select=id,name,color,canEdit,canShare,canViewPrivateItems,owner');
  }

  async listEvents(input: {
    calendarId?: string;
    start: string;
    end: string;
    maxResults: number;
  }): Promise<unknown> {
    const top = Math.max(1, Math.min(input.maxResults, 100));
    const basePath = input.calendarId
      ? `/me/calendars/${encodeURIComponent(input.calendarId)}/calendarView`
      : '/me/calendarView';
    const params = new URLSearchParams({
      startDateTime: new Date(input.start).toISOString(),
      endDateTime: new Date(input.end).toISOString(),
      $top: String(top),
      $orderby: 'start/dateTime',
      $select:
        'id,subject,start,end,isAllDay,location,organizer,attendees,bodyPreview,webLink,lastModifiedDateTime',
    });
    return this.graphRequest(`${basePath}?${params.toString()}`);
  }

  async createEvent(input: {
    calendarId?: string;
    subject: string;
    start: string;
    end: string;
    is_all_day?: boolean;
    location?: string;
    body?: string;
    attendees?: string;
  }): Promise<unknown> {
    const basePath = input.calendarId
      ? `/me/calendars/${encodeURIComponent(input.calendarId)}/events`
      : '/me/events';
    return this.graphRequest(basePath, {
      method: 'POST',
      body: buildEventPayload(input),
    });
  }

  async updateEvent(
    eventId: string,
    input: {
      subject?: string;
      start?: string;
      end?: string;
      is_all_day?: boolean;
      location?: string;
      body?: string;
      attendees?: string;
    },
  ): Promise<unknown> {
    return this.graphRequest(`/me/events/${encodeURIComponent(eventId)}`, {
      method: 'PATCH',
      body: buildEventPayload(input),
    });
  }

  async deleteEvent(eventId: string): Promise<{ success: boolean }> {
    await this.graphRequest(`/me/events/${encodeURIComponent(eventId)}`, {
      method: 'DELETE',
    });
    return { success: true };
  }

  async listContacts(input: { query?: string; maxResults: number }): Promise<unknown> {
    const top = Math.max(1, Math.min(input.maxResults, 200));
    const params = new URLSearchParams({
      $top: String(top),
      $orderby: 'displayName',
      $select:
        'id,displayName,givenName,surname,emailAddresses,mobilePhone,businessPhones,companyName,jobTitle',
    });

    const result = await this.graphRequest<{ value?: Array<Record<string, unknown>> }>(
      `/me/contacts?${params.toString()}`,
    );

    if (!input.query?.trim()) {
      return result;
    }

    const query = input.query.trim().toLowerCase();
    return {
      value: (result.value ?? []).filter((contact) => {
        const fields = [
          contact.displayName,
          contact.givenName,
          contact.surname,
          contact.companyName,
          contact.jobTitle,
          ...(Array.isArray(contact.emailAddresses)
            ? contact.emailAddresses.flatMap((item) =>
                typeof item === 'object' && item !== null
                  ? [Reflect.get(item, 'address'), Reflect.get(item, 'name')]
                  : [],
              )
            : []),
        ];

        return fields.some(
          (field) => typeof field === 'string' && field.toLowerCase().includes(query),
        );
      }),
    };
  }

  async createContact(input: {
    given_name?: string;
    surname?: string;
    display_name?: string;
    email_addresses?: string;
    mobile_phone?: string;
    business_phones?: string;
    company_name?: string;
    job_title?: string;
    notes?: string;
  }): Promise<unknown> {
    return this.graphRequest('/me/contacts', {
      method: 'POST',
      body: buildContactPayload(input),
    });
  }

  async updateContact(
    contactId: string,
    input: {
      given_name?: string;
      surname?: string;
      display_name?: string;
      email_addresses?: string;
      mobile_phone?: string;
      business_phones?: string;
      company_name?: string;
      job_title?: string;
      notes?: string;
    },
  ): Promise<unknown> {
    return this.graphRequest(`/me/contacts/${encodeURIComponent(contactId)}`, {
      method: 'PATCH',
      body: buildContactPayload(input),
    });
  }

  async deleteContact(contactId: string): Promise<{ success: boolean }> {
    await this.graphRequest(`/me/contacts/${encodeURIComponent(contactId)}`, {
      method: 'DELETE',
    });
    return { success: true };
  }
}
