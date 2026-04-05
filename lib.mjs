import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import process from 'node:process';
import graphPkg from '@microsoft/microsoft-graph-client';

import {readJsonFile, writeJsonFile, withLockedFile} from './auth-store.mjs';

const {Client} = graphPkg;

const DEFAULT_STATE_DIR = path.join(
  process.env.XDG_STATE_HOME || path.join(os.homedir(), '.local', 'state'),
  'outlook-mcp'
);
const DEFAULT_TOKEN_FILE = path.join(DEFAULT_STATE_DIR, 'tokens.json');
const DEFAULT_NOTIFICATION_LOG = path.join(DEFAULT_STATE_DIR, 'notifications.ndjson');
const DEFAULT_SCOPES =
  'offline_access openid profile User.Read Mail.Read Mail.ReadWrite Mail.Send Calendars.Read Calendars.ReadWrite';

export function buildMessageSelect({includeBody = false} = {}) {
  const fields = [
    'id',
    'conversationId',
    'subject',
    'from',
    'sender',
    'toRecipients',
    'ccRecipients',
    'receivedDateTime',
    'sentDateTime',
    'isRead',
    'hasAttachments',
    'importance',
    'categories',
    'bodyPreview',
    'webLink',
    'parentFolderId',
  ];
  if (includeBody) fields.push('body', 'internetMessageHeaders', 'replyTo');
  return fields.join(',');
}

export function parseEnvFile(file) {
  if (!file || !fs.existsSync(file)) return {};
  const text = fs.readFileSync(file, 'utf8');
  const env = {};
  for (const raw of text.split(/\r?\n/)) {
    const line = raw.trim();
    if (!line || line.startsWith('#')) continue;
    const idx = line.indexOf('=');
    if (idx < 0) continue;
    const key = line.slice(0, idx).trim();
    let value = line.slice(idx + 1).trim();
    if (
      (value.startsWith('"') && value.endsWith('"')) ||
      (value.startsWith("'") && value.endsWith("'"))
    ) {
      value = value.slice(1, -1);
    }
    env[key] = value;
  }
  return env;
}

export function loadConfig(envFile) {
  const fileEnv = parseEnvFile(envFile);
  const inheritedEnv = {...process.env};
  if (envFile) {
    for (const key of Object.keys(inheritedEnv)) {
      if (key.startsWith('OUTLOOK_') || key.startsWith('MS_')) {
        delete inheritedEnv[key];
      }
    }
  }
  const env = envFile ? {...inheritedEnv, ...fileEnv} : {...fileEnv, ...inheritedEnv};
  return {
    envFile,
    envDir: envFile ? path.dirname(envFile) : process.cwd(),
    env,
  };
}

export function resolveMaybeRelative(baseDir, value, fallback = '') {
  if (!value) return fallback;
  if (path.isAbsolute(value)) return value;
  return path.join(baseDir, value);
}

export function getTenant(config) {
  return config.env.OUTLOOK_TENANT_ID || 'common';
}

export function getScopeString(config) {
  return `${config.env.OUTLOOK_OAUTH_SCOPES || DEFAULT_SCOPES}`.trim();
}

export function getTokenFilePath(config) {
  const configured = config.env.OUTLOOK_OAUTH_TOKEN_FILE || DEFAULT_TOKEN_FILE;
  return resolveMaybeRelative(config.envDir, configured);
}

export function getNotificationLogPath(config) {
  return resolveMaybeRelative(
    config.envDir,
    config.env.OUTLOOK_NOTIFICATION_LOG_FILE,
    DEFAULT_NOTIFICATION_LOG,
  );
}

export function getReceiverConfig(config) {
  return {
    host: config.env.OUTLOOK_RECEIVER_HOST || '127.0.0.1',
    port: Number(config.env.OUTLOOK_RECEIVER_PORT || 8777),
    path: config.env.OUTLOOK_RECEIVER_PATH || '/graph/notifications',
    notificationUrl:
      config.env.OUTLOOK_NOTIFICATION_URL ||
      `http://${config.env.OUTLOOK_RECEIVER_HOST || '127.0.0.1'}:${Number(config.env.OUTLOOK_RECEIVER_PORT || 8777)}${config.env.OUTLOOK_RECEIVER_PATH || '/graph/notifications'}`,
    lifecycleNotificationUrl:
      config.env.OUTLOOK_LIFECYCLE_NOTIFICATION_URL ||
      `http://${config.env.OUTLOOK_RECEIVER_HOST || '127.0.0.1'}:${Number(config.env.OUTLOOK_RECEIVER_PORT || 8777)}${config.env.OUTLOOK_RECEIVER_PATH || '/graph/notifications'}`,
    clientState: config.env.OUTLOOK_SUBSCRIPTION_CLIENT_STATE || '',
    maxBodyBytes: Number(config.env.OUTLOOK_NOTIFICATION_MAX_BODY_BYTES || 262144),
    maxLogBytes: Number(config.env.OUTLOOK_NOTIFICATION_MAX_LOG_BYTES || 5242880),
    logFile: getNotificationLogPath(config),
  };
}

export function buildPath(pathname, query = {}) {
  const params = new URLSearchParams();
  for (const [key, value] of Object.entries(query)) {
    if (value === undefined || value === null || value === '') continue;
    params.set(key, `${value}`);
  }
  const encoded = params.toString();
  return encoded ? `${pathname}?${encoded}` : pathname;
}

async function fetchJson(url, options = {}) {
  const response = await fetch(url, options);
  const text = await response.text();
  let body;
  try {
    body = text ? JSON.parse(text) : {};
  } catch {
    body = text;
  }
  return {response, body};
}

export async function refreshUserToken(config) {
  const tokenFile = getTokenFilePath(config);
  return withLockedFile(tokenFile, async () => {
    const tokens = readJsonFile(tokenFile, null, {strict: true});
    const clientId = config.env.OUTLOOK_CLIENT_ID;
    const clientSecret = config.env.OUTLOOK_CLIENT_SECRET;
    const redirectUri = config.env.OUTLOOK_REDIRECT_URI;
    if (!clientId || !clientSecret || !redirectUri || !tokens?.refresh_token) {
      throw new Error('Missing Outlook OAuth refresh configuration');
    }

    if (Number(tokens.expires_at || 0) > Date.now() + 60_000) {
      return tokens;
    }

    const form = new URLSearchParams();
    form.set('client_id', clientId);
    form.set('client_secret', clientSecret);
    form.set('grant_type', 'refresh_token');
    form.set('refresh_token', tokens.refresh_token);
    form.set('redirect_uri', redirectUri);
    form.set('scope', getScopeString(config));

    const {response, body} = await fetchJson(
      `https://login.microsoftonline.com/${encodeURIComponent(getTenant(config))}/oauth2/v2.0/token`,
      {
        method: 'POST',
        headers: {'content-type': 'application/x-www-form-urlencoded'},
        body: form.toString(),
      },
    );

    if (!response.ok) {
      const error = new Error(`Outlook OAuth refresh failed with ${response.status}`);
      error.status = response.status;
      error.body = body;
      throw error;
    }

    const now = Date.now();
    const refreshed = {
      ...tokens,
      ...body,
      obtained_at: now,
      expires_at: now + Number(body.expires_in || 0) * 1000,
    };
    writeJsonFile(tokenFile, refreshed);
    return refreshed;
  });
}

export async function getUserToken(config, {refreshIfNeeded = true} = {}) {
  if (config.env.OUTLOOK_USER_TOKEN) return config.env.OUTLOOK_USER_TOKEN;
  const tokens = readJsonFile(getTokenFilePath(config), null, {strict: refreshIfNeeded});
  if (!tokens?.access_token) return null;
  if (!refreshIfNeeded) return tokens.access_token;
  if (Number(tokens.expires_at || 0) > Date.now() + 60_000) {
    return tokens.access_token;
  }
  const refreshed = await refreshUserToken(config);
  return refreshed.access_token;
}

export function createGraphApi(token) {
  const client = Client.init({
    authProvider: (done) => done(null, token),
  });

  async function fetchAbsolute(url, {method = 'GET', headers = {}, body} = {}) {
    const response = await fetch(url, {
      method,
      headers: {
        Authorization: `Bearer ${token}`,
        ...headers,
      },
      body,
    });
    const text = await response.text();
    let payload;
    try {
      payload = text ? JSON.parse(text) : {};
    } catch {
      payload = text;
    }
    if (!response.ok) {
      const error = new Error(`Microsoft Graph ${method} ${url} failed with ${response.status}`);
      error.status = response.status;
      error.body = payload;
      throw error;
    }
    return payload;
  }

  async function request(method, apiPath, {headers = {}, body = undefined} = {}) {
    if (String(apiPath).startsWith('http://') || String(apiPath).startsWith('https://')) {
      const requestBody =
        body && typeof body === 'object' && !(body instanceof URLSearchParams)
          ? JSON.stringify(body)
          : body;
      return fetchAbsolute(apiPath, {
        method,
        headers:
          body && typeof body === 'object' && !(body instanceof URLSearchParams)
            ? {'content-type': 'application/json', ...headers}
            : headers,
        body: requestBody,
      });
    }

    let req = client.api(apiPath).version('v1.0');
    for (const [key, value] of Object.entries(headers)) {
      req = req.header(key, value);
    }

    if (method === 'GET') return req.get();
    if (method === 'POST') return req.post(body);
    if (method === 'PATCH') return req.patch(body);
    if (method === 'PUT') return req.put(body);
    if (method === 'DELETE') return req.delete();
    throw new Error(`Unsupported Graph method ${method}`);
  }

  async function listCollection(apiPath, {headers = {}, max = 50} = {}) {
    const items = [];
    let nextPath = apiPath;
    while (nextPath && items.length < max) {
      const page = await request('GET', nextPath, {headers});
      items.push(...(page.value || []));
      nextPath = page['@odata.nextLink'] || null;
    }
    return items.slice(0, max);
  }

  return {request, listCollection, fetchAbsolute};
}

export function formatEmailAddress(item) {
  const address = item?.emailAddress || item || {};
  return {
    name: address.name || '',
    address: address.address || '',
  };
}

export function normalizeHtmlToText(value) {
  return `${value || ''}`
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n\n')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/\r\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .replace(/[ \t]{2,}/g, ' ')
    .trim();
}

export function firstMeaningfulParagraph(input) {
  return `${input || ''}`
    .split(/\n{2,}/)
    .map((part) => part.trim())
    .find((part) => part.length > 0 && !/^Hello[, ]/i.test(part) && !/^To view the latest comment/i.test(part)) ||
    '';
}

export function extractTicketUrl(input) {
  const match = /https:\/\/drfirsthelp\.zendesk\.com\/agent\/tickets\/(\d+)/i.exec(`${input || ''}`);
  return match ? match[0] : null;
}

export function extractTicketId(input) {
  const url = extractTicketUrl(input);
  if (url) {
    const match = /\/agent\/tickets\/(\d+)/.exec(url);
    return match?.[1] ?? null;
  }
  const follower = /You are a follower on this request\s*\((\d+)\)/i.exec(`${input || ''}`);
  if (follower) {
    return follower[1] ?? null;
  }
  return null;
}

export function normalizeTicketSubject(subject) {
  const trimmed = `${subject || ''}`.trim();
  const withoutExternal = trimmed.replace(/^\[EXTERNAL\]\s*/i, '');
  const ticketPrefixed = /^\d+\s*-\s*(.+)$/.exec(withoutExternal);
  return (ticketPrefixed?.[1] ?? withoutExternal).trim();
}

export function summarizeSupportMessage(message) {
  const preview = typeof message?.bodyPreview === 'string' ? message.bodyPreview.trim() : '';
  const bodyObject =
    message?.body && typeof message.body === 'object' && !Array.isArray(message.body)
      ? message.body
      : null;
  const bodyContent = typeof bodyObject?.content === 'string' ? bodyObject.content : '';
  const normalized = normalizeHtmlToText(bodyContent);
  return firstMeaningfulParagraph(normalized) || preview;
}

export function isZendeskFollowerEmail(message) {
  const subject = typeof message?.subject === 'string' ? message.subject : '';
  const preview = typeof message?.bodyPreview === 'string' ? message.bodyPreview : '';
  const bodyObject =
    message?.body && typeof message.body === 'object' && !Array.isArray(message.body)
      ? message.body
      : null;
  const bodyContent = typeof bodyObject?.content === 'string' ? bodyObject.content : '';
  const haystack = `${subject}\n${preview}\n${normalizeHtmlToText(bodyContent)}`;
  return (
    /You are a follower on this request/i.test(haystack) ||
    /Reply to this email to add an internal note/i.test(haystack) ||
    /Ticket link:\s*https:\/\/drfirsthelp\.zendesk\.com\/agent\/tickets\//i.test(haystack)
  );
}

export function classifyMessage(message) {
  const summary = summarizeSupportMessage(message);
  const subject = typeof message?.subject === 'string' ? message.subject : '';
  const from = formatEmailAddress(message?.from);
  const senderAddress = (from.address || '').toLowerCase();
  const ticketInput = `${subject}\n${message?.bodyPreview || ''}\n${summary}`;
  const ticketId = extractTicketId(ticketInput);
  const ticketUrl = extractTicketUrl(ticketInput);
  const zendeskFollower = isZendeskFollowerEmail(message);
  const webexNotice = senderAddress === 'messenger@webex.com';
  let label = 'generic';
  if (zendeskFollower) {
    label = 'zendesk_follower';
  } else if (webexNotice) {
    label = 'webex_notification';
  } else if (senderAddress === 'donotreply@drfirst.com' && ticketId) {
    label = 'support_mail';
  }
  return {
    label,
    ticketId,
    ticketUrl,
    normalizedSubject: normalizeTicketSubject(subject),
    summary,
    sender: from,
    isZendeskFollower: zendeskFollower,
    isWebexNotification: webexNotice,
  };
}

export function extractNotificationMessageIds(payload, clientState = '') {
  const root = payload && typeof payload === 'object' ? payload : null;
  const values = Array.isArray(root?.value) ? root.value : [];
  const messageIds = new Set();
  for (const rawValue of values) {
    const value = rawValue && typeof rawValue === 'object' ? rawValue : null;
    if (!value) continue;
    const incomingClientState =
      typeof value.clientState === 'string' ? value.clientState : undefined;
    if (clientState && incomingClientState && incomingClientState !== clientState) {
      throw new Error('Outlook notification clientState mismatch.');
    }
    const resourceData =
      value.resourceData && typeof value.resourceData === 'object' ? value.resourceData : null;
    const directId = typeof resourceData?.id === 'string' ? resourceData.id : null;
    const resource =
      typeof value.resource === 'string'
        ? value.resource
        : typeof value.resourceUrl === 'string'
          ? value.resourceUrl
          : null;
    const resourceId = resource ? /messages\/([^/?]+)/i.exec(resource)?.[1] ?? null : null;
    if (directId) messageIds.add(directId);
    if (resourceId) messageIds.add(resourceId);
  }
  return [...messageIds];
}

export async function fetchMessagesByIds(api, ids, {select = buildMessageSelect({includeBody: true})} = {}) {
  const messages = [];
  for (const id of ids) {
    try {
      const message = await api.request(
        'GET',
        `/me/messages/${encodeURIComponent(id)}?$select=${select}`,
      );
      messages.push(message);
    } catch {
      // Ignore stale or inaccessible notifications and continue.
    }
  }
  return messages;
}

export async function fetchMessageAttachments(api, messageId) {
  return api.listCollection(
    buildPath(`/me/messages/${encodeURIComponent(messageId)}/attachments`, {
      '$top': 200,
      '$select': 'id,name,contentType,size,isInline,lastModifiedDateTime',
    }),
    {max: 200},
  );
}

export async function ensureGraphSubscription(
  config,
  api,
  {
    resource = '/me/messages',
    changeType = 'created,updated',
    expirationWindowMs = 2 * 24 * 60 * 60 * 1000,
    renewIfExpiringWithinMs = 6 * 60 * 60 * 1000,
    listMax = 100,
  } = {},
) {
  const receiver = getReceiverConfig(config);
  if (!receiver.notificationUrl.startsWith('https://')) {
    return {status: 'fallback', receiver, subscription: null};
  }
  const desiredExpiry = new Date(Date.now() + expirationWindowMs).toISOString();
  const subscriptions = await api.listCollection('/subscriptions', {max: listMax});
  const existing = subscriptions.find(
    (subscription) =>
      subscription?.notificationUrl === receiver.notificationUrl &&
      subscription?.resource === resource,
  );
  if (!existing) {
    const subscription = await api.request('POST', '/subscriptions', {
      body: {
        changeType,
        notificationUrl: receiver.notificationUrl,
        lifecycleNotificationUrl: receiver.lifecycleNotificationUrl,
        clientState: receiver.clientState || undefined,
        resource,
        expirationDateTime: desiredExpiry,
      },
    });
    return {status: 'created', receiver, subscription};
  }
  const expiresAt =
    typeof existing.expirationDateTime === 'string'
      ? Date.parse(existing.expirationDateTime)
      : Number.NaN;
  if (!Number.isFinite(expiresAt) || expiresAt < Date.now() + renewIfExpiringWithinMs) {
    const subscription = await api.request('PATCH', `/subscriptions/${encodeURIComponent(existing.id)}`, {
      body: {
        expirationDateTime: desiredExpiry,
      },
    });
    return {status: 'renewed', receiver, subscription: subscription || existing};
  }
  return {status: 'active', receiver, subscription: existing};
}

export function readNotificationLogEntries(config, max = 50) {
  const file = getNotificationLogPath(config);
  if (!fs.existsSync(file)) {
    return [];
  }

  const stat = fs.statSync(file);
  const bytesToRead = Math.min(stat.size, 512 * 1024);
  const fd = fs.openSync(file, 'r');
  const buffer = Buffer.alloc(bytesToRead);
  fs.readSync(fd, buffer, 0, bytesToRead, stat.size - bytesToRead);
  fs.closeSync(fd);

  return buffer
    .toString('utf8')
    .split(/\r?\n/)
    .filter(Boolean)
    .slice(-Math.min(max, 1000))
    .map((line) => {
      try {
        return JSON.parse(line);
      } catch {
        return {raw: line};
      }
    });
}
