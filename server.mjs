#!/usr/bin/env bun

import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import process from 'node:process';
import readline from 'node:readline';
import {Database} from 'bun:sqlite';
import graphPkg from '@microsoft/microsoft-graph-client';
import {ensurePrivateDir, readJsonFile, writeJsonFile, withLockedFile} from './auth-store.mjs';

const {Client} = graphPkg;

const SERVER_NAME = 'outlook';
const SERVER_VERSION = '0.1.0';
const PROTOCOL_VERSION = '2024-11-05';
const DEFAULT_STATE_DIR = path.join(
  process.env.XDG_STATE_HOME || path.join(os.homedir(), '.local', 'state'),
  'outlook-mcp'
);
const DEFAULT_TOKEN_FILE = path.join(DEFAULT_STATE_DIR, 'tokens.json');
const DEFAULT_INDEX_DB = path.join(DEFAULT_STATE_DIR, 'index.sqlite');
const DEFAULT_NOTIFICATION_LOG = path.join(DEFAULT_STATE_DIR, 'notifications.ndjson');

function parseArgs(argv) {
  const out = {envFile: process.env.ENV_FILE || ''};
  for (let index = 2; index < argv.length; index += 1) {
    const arg = argv[index];
    const next = argv[index + 1];
    if (arg === '--env-file' && next) {
      out.envFile = next;
      index += 1;
    }
  }
  return out;
}

function parseEnvFile(file) {
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

function loadConfig(envFile) {
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

function resolveMaybeRelative(baseDir, value, fallback = '') {
  if (!value) return fallback;
  if (path.isAbsolute(value)) return value;
  return path.join(baseDir, value);
}

function getTenant(config) {
  return config.env.OUTLOOK_TENANT_ID || 'common';
}

function getTokenFilePath(config) {
  const configured = config.env.OUTLOOK_OAUTH_TOKEN_FILE || DEFAULT_TOKEN_FILE;
  return resolveMaybeRelative(config.envDir, configured);
}

function getScopeString(config) {
  return `${config.env.OUTLOOK_OAUTH_SCOPES || ''}`.trim();
}

function getIndexDbPath(config) {
  return resolveMaybeRelative(config.envDir, config.env.OUTLOOK_MCP_INDEX_DB, DEFAULT_INDEX_DB);
}

function getNotificationLogPath(config) {
  return resolveMaybeRelative(config.envDir, config.env.OUTLOOK_NOTIFICATION_LOG_FILE, DEFAULT_NOTIFICATION_LOG);
}

function getReceiverConfig(config) {
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
    logFile: getNotificationLogPath(config),
  };
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

async function refreshUserToken(config) {
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
      }
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

async function getUserToken(config, {refreshIfNeeded = true} = {}) {
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

function contentText(text) {
  return {content: [{type: 'text', text}]};
}

function jsonText(value) {
  return contentText(JSON.stringify(value, null, 2));
}

function errorText(message, extra = null) {
  return {content: [{type: 'text', text: extra ? `${message}\n\n${extra}` : message}], isError: true};
}

function buildPath(pathname, query = {}) {
  const params = new URLSearchParams();
  for (const [key, value] of Object.entries(query)) {
    if (value === undefined || value === null || value === '') continue;
    params.set(key, `${value}`);
  }
  const encoded = params.toString();
  return encoded ? `${pathname}?${encoded}` : pathname;
}

function createGraphApi(token) {
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
        headers: body && typeof body === 'object' && !(body instanceof URLSearchParams)
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

const runtime = {
  user: null,
  index: null,
};

async function getUserContext(config) {
  const token = await getUserToken(config, {refreshIfNeeded: true});
  if (!token) {
    throw new Error('Outlook user auth is unavailable. Configure OUTLOOK_USER_TOKEN or Outlook OAuth credentials.');
  }
  if (runtime.user?.token === token) return runtime.user;
  const api = createGraphApi(token);
  const me = await api.request('GET', '/me');
  runtime.user = {api, me, token};
  return runtime.user;
}

function formatEmailAddress(item) {
  const address = item?.emailAddress || item || {};
  return {
    name: address.name || '',
    address: address.address || '',
  };
}

function summarizeMessage(message) {
  return {
    id: message.id,
    conversationId: message.conversationId || '',
    subject: message.subject || '',
    from: formatEmailAddress(message.from),
    sender: formatEmailAddress(message.sender),
    toRecipients: (message.toRecipients || []).map(formatEmailAddress),
    ccRecipients: (message.ccRecipients || []).map(formatEmailAddress),
    receivedDateTime: message.receivedDateTime || null,
    sentDateTime: message.sentDateTime || null,
    isRead: Boolean(message.isRead),
    hasAttachments: Boolean(message.hasAttachments),
    importance: message.importance || '',
    categories: message.categories || [],
    bodyPreview: message.bodyPreview || '',
    webLink: message.webLink || '',
    parentFolderId: message.parentFolderId || '',
  };
}

function summarizeFolder(folder) {
  return {
    id: folder.id,
    displayName: folder.displayName || '',
    parentFolderId: folder.parentFolderId || null,
    childFolderCount: folder.childFolderCount || 0,
    unreadItemCount: folder.unreadItemCount || 0,
    totalItemCount: folder.totalItemCount || 0,
    wellKnownName: folder.wellKnownName || '',
    isHidden: Boolean(folder.isHidden),
  };
}

function summarizeSubscription(sub) {
  return {
    id: sub.id,
    resource: sub.resource || '',
    applicationId: sub.applicationId || '',
    changeType: sub.changeType || '',
    notificationUrl: sub.notificationUrl || '',
    expirationDateTime: sub.expirationDateTime || null,
    clientState: sub.clientState || '',
    latestSupportedTlsVersion: sub.latestSupportedTlsVersion || '',
  };
}

function summarizeCalendar(calendar) {
  return {
    id: calendar.id,
    name: calendar.name || '',
    color: calendar.color || '',
    canEdit: Boolean(calendar.canEdit),
    canShare: Boolean(calendar.canShare),
    canViewPrivateItems: Boolean(calendar.canViewPrivateItems),
    owner: formatEmailAddress(calendar.owner),
    isDefaultCalendar: Boolean(calendar.isDefaultCalendar),
  };
}

function summarizeEvent(event) {
  return {
    id: event.id,
    subject: event.subject || '',
    bodyPreview: event.bodyPreview || '',
    start: event.start || null,
    end: event.end || null,
    location: event.location || null,
    locations: event.locations || [],
    organizer: formatEmailAddress(event.organizer),
    attendees: (event.attendees || []).map((item) => ({
      type: item.type || '',
      status: item.status || null,
      emailAddress: formatEmailAddress(item.emailAddress),
    })),
    isAllDay: Boolean(event.isAllDay),
    isCancelled: Boolean(event.isCancelled),
    showAs: event.showAs || '',
    responseRequested: Boolean(event.responseRequested),
    webLink: event.webLink || '',
    onlineMeetingUrl: event.onlineMeeting?.joinUrl || event.onlineMeetingUrl || '',
    seriesMasterId: event.seriesMasterId || null,
    iCalUId: event.iCalUId || '',
    createdDateTime: event.createdDateTime || null,
    lastModifiedDateTime: event.lastModifiedDateTime || null,
  };
}

function getIndexStore(config) {
  const dbPath = getIndexDbPath(config);
  if (runtime.index?.path === dbPath) return runtime.index;

  ensurePrivateDir(path.dirname(dbPath));
  const db = new Database(dbPath);
  try {
    fs.chmodSync(dbPath, 0o600);
  } catch {}

  db.exec(`
    PRAGMA journal_mode = WAL;
    PRAGMA synchronous = NORMAL;

    CREATE TABLE IF NOT EXISTS mail_messages (
      id TEXT PRIMARY KEY,
      conversation_id TEXT NOT NULL DEFAULT '',
      folder_id TEXT NOT NULL DEFAULT '',
      subject TEXT NOT NULL DEFAULT '',
      from_name TEXT NOT NULL DEFAULT '',
      from_address TEXT NOT NULL DEFAULT '',
      to_json TEXT NOT NULL DEFAULT '[]',
      cc_json TEXT NOT NULL DEFAULT '[]',
      received_at TEXT,
      sent_at TEXT,
      is_read INTEGER NOT NULL DEFAULT 0,
      has_attachments INTEGER NOT NULL DEFAULT 0,
      importance TEXT NOT NULL DEFAULT '',
      categories_json TEXT NOT NULL DEFAULT '[]',
      body_preview TEXT NOT NULL DEFAULT '',
      body_text TEXT NOT NULL DEFAULT '',
      web_link TEXT NOT NULL DEFAULT '',
      indexed_at INTEGER NOT NULL
    );

    CREATE TABLE IF NOT EXISTS sync_state (
      scope TEXT NOT NULL,
      scope_id TEXT NOT NULL,
      cursor TEXT,
      last_synced_at INTEGER NOT NULL,
      PRIMARY KEY (scope, scope_id)
    );

    CREATE INDEX IF NOT EXISTS idx_mail_messages_conversation_id
      ON mail_messages(conversation_id, received_at DESC);

    CREATE VIRTUAL TABLE IF NOT EXISTS mail_search USING fts5(
      message_id UNINDEXED,
      conversation_id UNINDEXED,
      folder_id UNINDEXED,
      subject,
      participants,
      body_preview,
      body_text,
      tokenize = 'unicode61'
    );
  `);

  runtime.index = {
    path: dbPath,
    backend: 'bun:sqlite',
    db,
    statements: {
      upsertMessage: db.query(`
        INSERT INTO mail_messages (
          id, conversation_id, folder_id, subject, from_name, from_address, to_json, cc_json,
          received_at, sent_at, is_read, has_attachments, importance, categories_json, body_preview, body_text, web_link, indexed_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(id) DO UPDATE SET
          conversation_id = excluded.conversation_id,
          folder_id = excluded.folder_id,
          subject = excluded.subject,
          from_name = excluded.from_name,
          from_address = excluded.from_address,
          to_json = excluded.to_json,
          cc_json = excluded.cc_json,
          received_at = excluded.received_at,
          sent_at = excluded.sent_at,
          is_read = excluded.is_read,
          has_attachments = excluded.has_attachments,
          importance = excluded.importance,
          categories_json = excluded.categories_json,
          body_preview = excluded.body_preview,
          body_text = excluded.body_text,
          web_link = excluded.web_link,
          indexed_at = excluded.indexed_at
      `),
      deleteMessage: db.query(`DELETE FROM mail_messages WHERE id = ?`),
      deleteSearchMessage: db.query(`DELETE FROM mail_search WHERE message_id = ?`),
      insertSearchMessage: db.query(`
        INSERT INTO mail_search (
          message_id, conversation_id, folder_id, subject, participants, body_preview, body_text
        ) VALUES (?, ?, ?, ?, ?, ?, ?)
      `),
      upsertSyncState: db.query(`
        INSERT INTO sync_state (scope, scope_id, cursor, last_synced_at)
        VALUES (?, ?, ?, ?)
        ON CONFLICT(scope, scope_id) DO UPDATE SET
          cursor = excluded.cursor,
          last_synced_at = excluded.last_synced_at
      `),
      getSyncState: db.query(`SELECT * FROM sync_state WHERE scope = ? AND scope_id = ?`),
      getConversationMessages: db.query(`
        SELECT * FROM mail_messages
        WHERE conversation_id = ?
        ORDER BY COALESCE(received_at, sent_at) ASC
      `),
      countMessages: db.query(`SELECT COUNT(*) AS count FROM mail_messages`),
    },
  };

  return runtime.index;
}

function summarizeCachedMessage(row) {
  return {
    id: row.id,
    conversationId: row.conversation_id,
    folderId: row.folder_id,
    subject: row.subject,
    from: {name: row.from_name, address: row.from_address},
    toRecipients: JSON.parse(row.to_json || '[]'),
    ccRecipients: JSON.parse(row.cc_json || '[]'),
    receivedDateTime: row.received_at || null,
    sentDateTime: row.sent_at || null,
    isRead: Boolean(row.is_read),
    hasAttachments: Boolean(row.has_attachments),
    importance: row.importance || '',
    categories: JSON.parse(row.categories_json || '[]'),
    bodyPreview: row.body_preview || '',
    webLink: row.web_link || '',
  };
}

function buildFtsQuery(query) {
  const normalized = `${query || ''}`.trim().toLowerCase();
  if (!normalized) return '';
  const tokens = normalized.split(/\s+/).filter(Boolean);
  const parts = [];
  if (normalized.includes(' ')) {
    parts.push(`"${normalized.replaceAll('"', '""')}"`);
  }
  for (const token of tokens) {
    parts.push(`${token.replaceAll('"', '""')}*`);
  }
  return [...new Set(parts)].join(' OR ');
}

function indexMailMessage(store, message) {
  const participants = [
    message.from?.emailAddress?.address || '',
    message.from?.emailAddress?.name || '',
    ...(message.toRecipients || []).flatMap((item) => [item.emailAddress?.address || '', item.emailAddress?.name || '']),
    ...(message.ccRecipients || []).flatMap((item) => [item.emailAddress?.address || '', item.emailAddress?.name || '']),
  ]
    .filter(Boolean)
    .join('\n');

  store.statements.upsertMessage.run(
    message.id,
    message.conversationId || '',
    message.parentFolderId || '',
    message.subject || '',
    message.from?.emailAddress?.name || '',
    message.from?.emailAddress?.address || '',
    JSON.stringify((message.toRecipients || []).map(formatEmailAddress)),
    JSON.stringify((message.ccRecipients || []).map(formatEmailAddress)),
    message.receivedDateTime || null,
    message.sentDateTime || null,
    message.isRead ? 1 : 0,
    message.hasAttachments ? 1 : 0,
    message.importance || '',
    JSON.stringify(message.categories || []),
    message.bodyPreview || '',
    message.body?.content || '',
    message.webLink || '',
    Date.now()
  );
  store.statements.deleteSearchMessage.run(message.id);
  store.statements.insertSearchMessage.run(
    message.id,
    message.conversationId || '',
    message.parentFolderId || '',
    message.subject || '',
    participants,
    message.bodyPreview || '',
    message.body?.content || ''
  );
}

function deleteIndexedMailMessage(store, messageId) {
  store.statements.deleteSearchMessage.run(messageId);
  store.statements.deleteMessage.run(messageId);
}

async function syncMailFolder(config, args = {}) {
  const context = await getUserContext(config);
  const store = getIndexStore(config);
  const folderScope = args.folderId || 'me';
  const scope = args.folderId ? 'mail_folder_delta' : 'mail_root_delta';
  const existing = args.fullResync
    ? null
    : store.statements.getSyncState.get(scope, folderScope) || null;

  let nextLink = existing?.cursor || buildPath(
    args.folderId ? `/me/mailFolders/${encodeURIComponent(args.folderId)}/messages/delta` : '/me/messages/delta',
    {
      '$top': Math.min(Number(args.pageSize || 50), 100),
      '$select': 'id,conversationId,parentFolderId,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,importance,categories,bodyPreview,body,webLink',
    }
  );

  let syncedCount = 0;
  let deletedCount = 0;
  let pageCount = 0;
  let deltaLink = existing?.cursor || null;

  while (nextLink && pageCount < Number(args.maxPages || 10)) {
    const page = await context.api.request('GET', nextLink);
    for (const item of page.value || []) {
      if (item['@removed']) {
        deleteIndexedMailMessage(store, item.id);
        deletedCount += 1;
      } else {
        indexMailMessage(store, item);
        syncedCount += 1;
      }
    }
    pageCount += 1;
    deltaLink = page['@odata.deltaLink'] || deltaLink;
    nextLink = page['@odata.nextLink'] || null;
  }

  if (deltaLink) {
    store.statements.upsertSyncState.run(scope, folderScope, deltaLink, Date.now());
  }

  return {
    folderId: args.folderId || null,
    pageCount,
    syncedCount,
    deletedCount,
    cursorStored: Boolean(deltaLink),
    index: {
      dbPath: store.path,
      backend: store.backend,
      messageCount: Number(store.statements.countMessages.get()?.count || 0),
    },
  };
}

function searchCachedMessages(config, args = {}) {
  const store = getIndexStore(config);
  const ftsQuery = buildFtsQuery(args.query);
  if (!ftsQuery) {
    return {
      count: 0,
      index: {dbPath: store.path, backend: store.backend},
      conversations: [],
    };
  }

  const rows = store.db.query(`
    SELECT
      m.*,
      bm25(mail_search) AS search_rank
    FROM mail_search
    JOIN mail_messages m
      ON m.id = mail_search.message_id
    WHERE mail_search MATCH ?
    ${args.folderId ? 'AND m.folder_id = ?' : ''}
    ORDER BY search_rank ASC, COALESCE(m.received_at, m.sent_at) DESC
    LIMIT ?
  `).all(...(args.folderId ? [ftsQuery, args.folderId, Math.max((args.maxResults || 20) * 5, 50)] : [ftsQuery, Math.max((args.maxResults || 20) * 5, 50)]));

  const grouped = new Map();
  for (const row of rows) {
    const key = row.conversation_id || row.id;
    const bucket = grouped.get(key) || {
      conversationId: key,
      subject: row.subject || '',
      messageCount: 0,
      latestReceivedDateTime: row.received_at || row.sent_at || null,
      messages: [],
    };
    bucket.messageCount += 1;
    if (!bucket.latestReceivedDateTime || new Date(row.received_at || row.sent_at || 0) > new Date(bucket.latestReceivedDateTime || 0)) {
      bucket.latestReceivedDateTime = row.received_at || row.sent_at || null;
    }
    bucket.messages.push(summarizeCachedMessage(row));
    grouped.set(key, bucket);
  }

  const conversations = [...grouped.values()]
    .sort((a, b) => new Date(b.latestReceivedDateTime || 0) - new Date(a.latestReceivedDateTime || 0))
    .slice(0, args.maxResults || 20);

  return {
    count: conversations.length,
    index: {dbPath: store.path, backend: store.backend},
    conversations,
  };
}

function listCachedConversationMessages(config, conversationId) {
  const store = getIndexStore(config);
  const rows = store.statements.getConversationMessages.all(conversationId);
  return {
    conversationId,
    count: rows.length,
    index: {dbPath: store.path, backend: store.backend},
    messages: rows.map(summarizeCachedMessage),
  };
}

function readNotificationLog(config, max = 50) {
  const file = getNotificationLogPath(config);
  if (!fs.existsSync(file)) {
    return {
      logFile: file,
      count: 0,
      notifications: [],
    };
  }

  const lines = fs.readFileSync(file, 'utf8').split(/\r?\n/).filter(Boolean);
  const notifications = lines
    .slice(-Math.min(max, lines.length))
    .map((line) => {
      try {
        return JSON.parse(line);
      } catch {
        return {raw: line};
      }
    });

  return {
    logFile: file,
    count: notifications.length,
    notifications,
  };
}

function normalizeRecipients(values = []) {
  return values.map((item) => {
    if (typeof item === 'string') {
      return {emailAddress: {address: item}};
    }
    if (item?.address) {
      return {emailAddress: {address: item.address, name: item.name || ''}};
    }
    if (item?.emailAddress?.address) {
      return {emailAddress: {address: item.emailAddress.address, name: item.emailAddress.name || ''}};
    }
    throw new Error('Recipients must be strings or {address, name} objects.');
  });
}

function buildMessageSelect({includeBody = false} = {}) {
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

const TOOLS = [
  {
    name: 'whoami',
    description: 'Show the authenticated Microsoft Graph user identity.',
    inputSchema: {type: 'object', properties: {}},
  },
  {
    name: 'list_mail_folders',
    description: 'List mail folders for the authenticated user.',
    inputSchema: {
      type: 'object',
      properties: {
        max: {type: 'integer', minimum: 1, maximum: 500},
        includeHidden: {type: 'boolean'},
      },
    },
  },
  {
    name: 'list_calendars',
    description: 'List calendars visible to the authenticated user.',
    inputSchema: {
      type: 'object',
      properties: {
        max: {type: 'integer', minimum: 1, maximum: 200},
      },
    },
  },
  {
    name: 'list_events',
    description: 'List events from the default calendar or one calendarId. If startDateTime and endDateTime are provided, use calendar view.',
    inputSchema: {
      type: 'object',
      properties: {
        calendarId: {type: 'string'},
        startDateTime: {type: 'string'},
        endDateTime: {type: 'string'},
        max: {type: 'integer', minimum: 1, maximum: 200},
      },
    },
  },
  {
    name: 'get_event',
    description: 'Fetch one event by eventId.',
    inputSchema: {
      type: 'object',
      required: ['eventId'],
      properties: {
        eventId: {type: 'string'},
      },
    },
  },
  {
    name: 'create_event',
    description: 'Create an event in the default calendar or one calendarId.',
    inputSchema: {
      type: 'object',
      required: ['event'],
      properties: {
        calendarId: {type: 'string'},
        event: {type: 'object', additionalProperties: true},
      },
    },
  },
  {
    name: 'update_event',
    description: 'Update an event by eventId.',
    inputSchema: {
      type: 'object',
      required: ['eventId', 'changes'],
      properties: {
        eventId: {type: 'string'},
        changes: {type: 'object', additionalProperties: true},
      },
    },
  },
  {
    name: 'delete_event',
    description: 'Delete an event by eventId.',
    inputSchema: {
      type: 'object',
      required: ['eventId'],
      properties: {
        eventId: {type: 'string'},
      },
    },
  },
  {
    name: 'list_messages',
    description: 'List recent messages, optionally constrained to one mail folder or one Outlook conversation.',
    inputSchema: {
      type: 'object',
      properties: {
        folderId: {type: 'string'},
        conversationId: {type: 'string'},
        unreadOnly: {type: 'boolean'},
        includeBody: {type: 'boolean'},
        max: {type: 'integer', minimum: 1, maximum: 100},
      },
    },
  },
  {
    name: 'search_messages',
    description: 'Search Outlook mail with Microsoft Graph $search.',
    inputSchema: {
      type: 'object',
      required: ['query'],
      properties: {
        query: {type: 'string'},
        folderId: {type: 'string'},
        max: {type: 'integer', minimum: 1, maximum: 100},
      },
    },
  },
  {
    name: 'get_message',
    description: 'Fetch one Outlook message by messageId.',
    inputSchema: {
      type: 'object',
      required: ['messageId'],
      properties: {
        messageId: {type: 'string'},
      },
    },
  },
  {
    name: 'list_conversation_messages',
    description: 'List messages for one Outlook conversationId.',
    inputSchema: {
      type: 'object',
      required: ['conversationId'],
      properties: {
        conversationId: {type: 'string'},
        folderId: {type: 'string'},
        max: {type: 'integer', minimum: 1, maximum: 100},
      },
    },
  },
  {
    name: 'sync_mail_folder',
    description: 'Sync one folder or the root mailbox into the local SQLite mail cache using Microsoft Graph delta.',
    inputSchema: {
      type: 'object',
      properties: {
        folderId: {type: 'string'},
        fullResync: {type: 'boolean'},
        maxPages: {type: 'integer', minimum: 1, maximum: 200},
        pageSize: {type: 'integer', minimum: 1, maximum: 100},
      },
    },
  },
  {
    name: 'search_cached_messages',
    description: 'Search the local SQLite mail cache and group hits by conversationId.',
    inputSchema: {
      type: 'object',
      required: ['query'],
      properties: {
        query: {type: 'string'},
        folderId: {type: 'string'},
        maxResults: {type: 'integer', minimum: 1, maximum: 100},
      },
    },
  },
  {
    name: 'list_cached_conversation_messages',
    description: 'List cached messages for one Outlook conversationId.',
    inputSchema: {
      type: 'object',
      required: ['conversationId'],
      properties: {
        conversationId: {type: 'string'},
      },
    },
  },
  {
    name: 'send_mail',
    description: 'Send a new Outlook email.',
    inputSchema: {
      type: 'object',
      required: ['to', 'subject'],
      properties: {
        to: {type: 'array', items: {type: ['string', 'object']}},
        cc: {type: 'array', items: {type: ['string', 'object']}},
        bcc: {type: 'array', items: {type: ['string', 'object']}},
        subject: {type: 'string'},
        body: {type: 'string'},
        bodyType: {type: 'string', enum: ['text', 'html']},
        saveToSentItems: {type: 'boolean'},
      },
    },
  },
  {
    name: 'reply_to_message',
    description: 'Reply or reply-all to one Outlook message.',
    inputSchema: {
      type: 'object',
      required: ['messageId', 'body'],
      properties: {
        messageId: {type: 'string'},
        body: {type: 'string'},
        replyAll: {type: 'boolean'},
      },
    },
  },
  {
    name: 'list_subscriptions',
    description: 'List active Microsoft Graph subscriptions visible to the current user/app context.',
    inputSchema: {
      type: 'object',
      properties: {
        max: {type: 'integer', minimum: 1, maximum: 200},
      },
    },
  },
  {
    name: 'create_subscription',
    description: 'Create a Microsoft Graph subscription, for example on /me/messages.',
    inputSchema: {
      type: 'object',
      required: ['resource', 'changeType', 'notificationUrl', 'expirationDateTime'],
      properties: {
        resource: {type: 'string'},
        changeType: {type: 'string'},
        notificationUrl: {type: 'string'},
        expirationDateTime: {type: 'string'},
        clientState: {type: 'string'},
        lifecycleNotificationUrl: {type: 'string'},
        latestSupportedTlsVersion: {type: 'string'},
      },
    },
  },
  {
    name: 'renew_subscription',
    description: 'Renew a Microsoft Graph subscription expiration.',
    inputSchema: {
      type: 'object',
      required: ['subscriptionId', 'expirationDateTime'],
      properties: {
        subscriptionId: {type: 'string'},
        expirationDateTime: {type: 'string'},
      },
    },
  },
  {
    name: 'delete_subscription',
    description: 'Delete a Microsoft Graph subscription.',
    inputSchema: {
      type: 'object',
      required: ['subscriptionId'],
      properties: {
        subscriptionId: {type: 'string'},
      },
    },
  },
  {
    name: 'receiver_status',
    description: 'Show the local Graph notification receiver scaffold configuration.',
    inputSchema: {type: 'object', properties: {}},
  },
  {
    name: 'list_received_notifications',
    description: 'List the most recently received Graph notification payloads captured by the local receiver scaffold.',
    inputSchema: {
      type: 'object',
      properties: {
        max: {type: 'integer', minimum: 1, maximum: 500},
      },
    },
  },
];

async function callTool(config, name, args = {}) {
  if (name === 'receiver_status') {
    return jsonText(getReceiverConfig(config));
  }
  if (name === 'list_received_notifications') {
    return jsonText(readNotificationLog(config, Number(args.max || 50)));
  }
  if (name === 'search_cached_messages') {
    return jsonText(searchCachedMessages(config, args));
  }
  if (name === 'list_cached_conversation_messages') {
    return jsonText(listCachedConversationMessages(config, args.conversationId));
  }

  const context = await getUserContext(config);
  const api = context.api;

  switch (name) {
    case 'whoami':
      return jsonText({
        id: context.me.id,
        displayName: context.me.displayName || '',
        userPrincipalName: context.me.userPrincipalName || '',
        mail: context.me.mail || '',
      });
    case 'list_mail_folders': {
      const folders = await api.listCollection(
        buildPath('/me/mailFolders', {
          '$top': Math.min(Number(args.max || 50), 200),
          '$select': 'id,displayName,parentFolderId,childFolderCount,unreadItemCount,totalItemCount,isHidden,wellKnownName',
          ...(args.includeHidden ? {} : {'$filter': 'isHidden eq false'}),
        }),
        {max: args.max || 50}
      );
      return jsonText({count: folders.length, folders: folders.map(summarizeFolder)});
    }
    case 'list_calendars': {
      const calendars = await api.listCollection(
        buildPath('/me/calendars', {
          '$top': Math.min(Number(args.max || 50), 200),
          '$select': 'id,name,color,canEdit,canShare,canViewPrivateItems,owner,isDefaultCalendar',
        }),
        {max: args.max || 50}
      );
      return jsonText({count: calendars.length, calendars: calendars.map(summarizeCalendar)});
    }
    case 'list_events': {
      const useCalendarView = Boolean(args.startDateTime && args.endDateTime);
      const base = useCalendarView
        ? args.calendarId
          ? `/me/calendars/${encodeURIComponent(args.calendarId)}/calendarView`
          : '/me/calendar/calendarView'
        : args.calendarId
          ? `/me/calendars/${encodeURIComponent(args.calendarId)}/events`
          : '/me/events';
      const events = await api.listCollection(
        buildPath(base, {
          '$top': Math.min(Number(args.max || 50), 200),
          ...(useCalendarView ? {} : {'$orderby': 'start/dateTime DESC'}),
          '$select': 'id,subject,bodyPreview,start,end,location,locations,organizer,attendees,isAllDay,isCancelled,showAs,responseRequested,webLink,onlineMeeting,onlineMeetingUrl,seriesMasterId,iCalUId,createdDateTime,lastModifiedDateTime',
          ...(useCalendarView
            ? {
                startDateTime: args.startDateTime,
                endDateTime: args.endDateTime,
              }
            : {}),
        }),
        {max: args.max || 50}
      );
      return jsonText({count: events.length, events: events.map(summarizeEvent)});
    }
    case 'get_event': {
      const event = await api.request(
        'GET',
        buildPath(`/me/events/${encodeURIComponent(args.eventId)}`, {
          '$select': 'id,subject,bodyPreview,start,end,location,locations,organizer,attendees,isAllDay,isCancelled,showAs,responseRequested,webLink,onlineMeeting,onlineMeetingUrl,seriesMasterId,iCalUId,createdDateTime,lastModifiedDateTime,body',
        })
      );
      return jsonText({
        ...summarizeEvent(event),
        body: event.body || null,
      });
    }
    case 'create_event': {
      const endpoint = args.calendarId
        ? `/me/calendars/${encodeURIComponent(args.calendarId)}/events`
        : '/me/events';
      const event = await api.request('POST', endpoint, {body: args.event});
      return jsonText(summarizeEvent(event));
    }
    case 'update_event': {
      await api.request('PATCH', `/me/events/${encodeURIComponent(args.eventId)}`, {
        body: args.changes,
      });
      const event = await api.request(
        'GET',
        buildPath(`/me/events/${encodeURIComponent(args.eventId)}`, {
          '$select': 'id,subject,bodyPreview,start,end,location,locations,organizer,attendees,isAllDay,isCancelled,showAs,responseRequested,webLink,onlineMeeting,onlineMeetingUrl,seriesMasterId,iCalUId,createdDateTime,lastModifiedDateTime',
        })
      );
      return jsonText(summarizeEvent(event));
    }
    case 'delete_event': {
      await api.request('DELETE', `/me/events/${encodeURIComponent(args.eventId)}`);
      return jsonText({deleted: true, eventId: args.eventId});
    }
    case 'list_messages': {
      const base = args.folderId ? `/me/mailFolders/${encodeURIComponent(args.folderId)}/messages` : '/me/messages';
      const filters = [];
      if (args.unreadOnly) filters.push('isRead eq false');
      if (args.conversationId) filters.push(`conversationId eq '${String(args.conversationId).replaceAll("'", "''")}'`);
      const messages = await api.listCollection(
        buildPath(base, {
          '$top': Math.min(Number(args.max || 25), 100),
          '$orderby': 'receivedDateTime DESC',
          '$select': buildMessageSelect({includeBody: Boolean(args.includeBody)}),
          ...(filters.length ? {'$filter': filters.join(' and ')} : {}),
        }),
        {max: args.max || 25}
      );
      return jsonText({count: messages.length, messages: messages.map(summarizeMessage)});
    }
    case 'search_messages': {
      const base = args.folderId ? `/me/mailFolders/${encodeURIComponent(args.folderId)}/messages` : '/me/messages';
      const messages = await api.listCollection(
        buildPath(base, {
          '$top': Math.min(Number(args.max || 25), 100),
          '$search': `"${String(args.query).replaceAll('"', '\\"')}"`,
          '$select': buildMessageSelect({includeBody: false}),
        }),
        {
          max: args.max || 25,
          headers: {'ConsistencyLevel': 'eventual'},
        }
      );
      return jsonText({count: messages.length, messages: messages.map(summarizeMessage)});
    }
    case 'get_message': {
      const message = await api.request(
        'GET',
        buildPath(`/me/messages/${encodeURIComponent(args.messageId)}`, {
          '$select': buildMessageSelect({includeBody: true}),
        })
      );
      return jsonText({
        ...summarizeMessage(message),
        body: message.body || null,
        replyTo: (message.replyTo || []).map(formatEmailAddress),
        internetMessageHeaders: message.internetMessageHeaders || [],
      });
    }
    case 'list_conversation_messages': {
      const base = args.folderId ? `/me/mailFolders/${encodeURIComponent(args.folderId)}/messages` : '/me/messages';
      const messages = await api.listCollection(
        buildPath(base, {
          '$top': Math.min(Number(args.max || 50), 100),
          '$orderby': 'receivedDateTime ASC',
          '$select': buildMessageSelect({includeBody: false}),
          '$filter': `conversationId eq '${String(args.conversationId).replaceAll("'", "''")}'`,
        }),
        {max: args.max || 50}
      );
      return jsonText({
        conversationId: args.conversationId,
        count: messages.length,
        messages: messages.map(summarizeMessage),
      });
    }
    case 'sync_mail_folder': {
      return jsonText(await syncMailFolder(config, args));
    }
    case 'search_cached_messages':
    case 'list_cached_conversation_messages':
    case 'receiver_status':
    case 'list_received_notifications':
      throw new Error(`Local-only tool dispatch bug for ${name}.`);
    case 'send_mail': {
      await api.request('POST', '/me/sendMail', {
        body: {
          message: {
            subject: args.subject,
            body: {
              contentType: (args.bodyType || 'text').toUpperCase(),
              content: args.body || '',
            },
            toRecipients: normalizeRecipients(args.to || []),
            ccRecipients: normalizeRecipients(args.cc || []),
            bccRecipients: normalizeRecipients(args.bcc || []),
          },
          saveToSentItems: args.saveToSentItems !== false,
        },
      });
      return jsonText({
        sent: true,
        subject: args.subject,
        toCount: (args.to || []).length,
      });
    }
    case 'reply_to_message': {
      const endpoint = args.replyAll ? 'replyAll' : 'reply';
      await api.request('POST', `/me/messages/${encodeURIComponent(args.messageId)}/${endpoint}`, {
        body: {
          comment: args.body,
        },
      });
      return jsonText({
        sent: true,
        messageId: args.messageId,
        replyAll: Boolean(args.replyAll),
      });
    }
    case 'list_subscriptions': {
      const subscriptions = await api.listCollection(
        buildPath('/subscriptions', {
          '$top': Math.min(Number(args.max || 50), 200),
        }),
        {max: args.max || 50}
      );
      return jsonText({
        count: subscriptions.length,
        subscriptions: subscriptions.map(summarizeSubscription),
      });
    }
    case 'create_subscription': {
      const subscription = await api.request('POST', '/subscriptions', {
        body: {
          resource: args.resource,
          changeType: args.changeType,
          notificationUrl: args.notificationUrl,
          expirationDateTime: args.expirationDateTime,
          clientState: args.clientState,
          lifecycleNotificationUrl: args.lifecycleNotificationUrl,
          latestSupportedTlsVersion: args.latestSupportedTlsVersion,
        },
      });
      return jsonText(summarizeSubscription(subscription));
    }
    case 'renew_subscription': {
      const subscription = await api.request('PATCH', `/subscriptions/${encodeURIComponent(args.subscriptionId)}`, {
        body: {
          expirationDateTime: args.expirationDateTime,
        },
      });
      return jsonText(summarizeSubscription(subscription));
    }
    case 'delete_subscription': {
      await api.request('DELETE', `/subscriptions/${encodeURIComponent(args.subscriptionId)}`);
      return jsonText({deleted: true, subscriptionId: args.subscriptionId});
    }
    default:
      throw new Error(`Unknown tool: ${name}`);
  }
}

function sendResponse(message) {
  process.stdout.write(`${JSON.stringify(message)}\n`);
}

function sendResult(id, result) {
  sendResponse({jsonrpc: '2.0', id, result});
}

function sendError(id, code, message, data = null) {
  sendResponse({
    jsonrpc: '2.0',
    id,
    error: {code, message, ...(data ? {data} : {})},
  });
}

async function handleRequest(config, message) {
  const {id, method, params = {}} = message;
  try {
    switch (method) {
      case 'initialize':
        sendResult(id, {
          protocolVersion: PROTOCOL_VERSION,
          serverInfo: {name: SERVER_NAME, version: SERVER_VERSION},
          capabilities: {tools: {}},
        });
        return;
      case 'tools/list':
        sendResult(id, {tools: TOOLS});
        return;
      case 'tools/call': {
        const result = await callTool(config, params.name, params.arguments || {});
        sendResult(id, result);
        return;
      }
      default:
        sendError(id, -32601, `Method not found: ${method}`);
    }
  } catch (error) {
    const detail = error?.stack || error?.message || String(error);
    sendResult(id, errorText(error.message || 'Tool execution failed', detail.slice(0, 8000)));
  }
}

async function main() {
  const args = parseArgs(process.argv);
  const config = loadConfig(args.envFile);
  const rl = readline.createInterface({
    input: process.stdin,
    crlfDelay: Infinity,
  });

  rl.on('line', (line) => {
    if (!line.trim()) return;
    let message;
    try {
      message = JSON.parse(line);
    } catch (error) {
      sendError(null, -32700, 'Parse error', error.message);
      return;
    }
    if (!Object.prototype.hasOwnProperty.call(message, 'id')) return;
    void handleRequest(config, message);
  });
}

main().catch((error) => {
  console.error(error.stack || error.message || String(error));
  process.exit(1);
});
