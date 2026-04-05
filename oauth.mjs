#!/usr/bin/env bun

import fs from 'node:fs';
import http from 'node:http';
import os from 'node:os';
import path from 'node:path';
import process from 'node:process';
import crypto from 'node:crypto';
import {spawn} from 'node:child_process';
import {ensurePrivateDir, readJsonFile, writeJsonFile, withLockedFile} from './auth-store.mjs';

const DEFAULT_STATE_DIR = path.join(
  process.env.XDG_STATE_HOME || path.join(os.homedir(), '.local', 'state'),
  'outlook-mcp'
);
const DEFAULT_TOKEN_FILE = path.join(DEFAULT_STATE_DIR, 'tokens.json');
const DEFAULT_SCOPES = 'offline_access openid profile User.Read Mail.Read Mail.Send Calendars.Read Calendars.ReadWrite';

function parseArgs(argv) {
  const out = {
    command: argv[2] || 'status',
    envFile: process.env.ENV_FILE || '',
    openBrowser: true,
    timeoutSeconds: 180,
  };

  for (let index = 3; index < argv.length; index += 1) {
    const arg = argv[index];
    const next = argv[index + 1];

    if (arg === '--env-file' && next) {
      out.envFile = next;
      index += 1;
      continue;
    }

    if (arg === '--no-open') {
      out.openBrowser = false;
      continue;
    }

    if (arg === '--timeout' && next) {
      out.timeoutSeconds = Number(next) || out.timeoutSeconds;
      index += 1;
      continue;
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
  return `${config.env.OUTLOOK_OAUTH_SCOPES || DEFAULT_SCOPES}`.trim();
}

function assertOauthConfig(config) {
  const missing = [];
  for (const key of ['OUTLOOK_CLIENT_ID', 'OUTLOOK_CLIENT_SECRET', 'OUTLOOK_REDIRECT_URI']) {
    if (!config.env[key]) missing.push(key);
  }
  if (!getScopeString(config)) missing.push('OUTLOOK_OAUTH_SCOPES');
  if (missing.length) {
    throw new Error(`Missing required Outlook OAuth env: ${missing.join(', ')}`);
  }
}

function oauthBase(config) {
  return `https://login.microsoftonline.com/${encodeURIComponent(getTenant(config))}/oauth2/v2.0`;
}

function buildAuthorizeUrl(config, state) {
  const query = new URLSearchParams();
  query.set('client_id', config.env.OUTLOOK_CLIENT_ID);
  query.set('response_type', 'code');
  query.set('redirect_uri', config.env.OUTLOOK_REDIRECT_URI);
  query.set('response_mode', 'query');
  query.set('scope', getScopeString(config));
  query.set('state', state);
  return `${oauthBase(config)}/authorize?${query.toString()}`;
}

async function exchangeToken(config, params) {
  const response = await fetch(`${oauthBase(config)}/token`, {
    method: 'POST',
    headers: {'content-type': 'application/x-www-form-urlencoded'},
    body: params.toString(),
  });
  const text = await response.text();
  let body;
  try {
    body = text ? JSON.parse(text) : {};
  } catch {
    body = text;
  }

  if (!response.ok) {
    throw new Error(`OAuth token exchange failed with ${response.status}: ${JSON.stringify(body)}`);
  }

  return body;
}

function persistTokenUnlocked(tokenFile, body) {
  const now = Date.now();
  const token = {
    ...body,
    obtained_at: now,
    expires_at: now + Number(body.expires_in || 0) * 1000,
  };
  writeJsonFile(tokenFile, token);
  return {tokenFile, token};
}

async function persistToken(config, body) {
  const tokenFile = getTokenFilePath(config);
  return withLockedFile(tokenFile, async () => persistTokenUnlocked(tokenFile, body));
}

async function exchangeAuthorizationCode(config, code) {
  const form = new URLSearchParams();
  form.set('client_id', config.env.OUTLOOK_CLIENT_ID);
  form.set('client_secret', config.env.OUTLOOK_CLIENT_SECRET);
  form.set('grant_type', 'authorization_code');
  form.set('code', code);
  form.set('redirect_uri', config.env.OUTLOOK_REDIRECT_URI);
  form.set('scope', getScopeString(config));
  const body = await exchangeToken(config, form);
  return persistToken(config, body);
}

async function refreshToken(config) {
  const tokenFile = getTokenFilePath(config);
  return withLockedFile(tokenFile, async () => {
    const existing = readJsonFile(tokenFile, null, {strict: true});
    if (!existing?.refresh_token) {
      throw new Error('No refresh_token is stored yet. Run auth:login first.');
    }
    const form = new URLSearchParams();
    form.set('client_id', config.env.OUTLOOK_CLIENT_ID);
    form.set('client_secret', config.env.OUTLOOK_CLIENT_SECRET);
    form.set('grant_type', 'refresh_token');
    form.set('refresh_token', existing.refresh_token);
    form.set('redirect_uri', config.env.OUTLOOK_REDIRECT_URI);
    form.set('scope', getScopeString(config));
    const body = await exchangeToken(config, form);
    return persistTokenUnlocked(tokenFile, {...existing, ...body});
  });
}

function openBrowser(url) {
  let command;
  let args;
  if (process.platform === 'darwin') {
    command = 'open';
    args = [url];
  } else if (process.platform === 'win32') {
    command = 'cmd';
    args = ['/c', 'start', '', url];
  } else {
    command = 'xdg-open';
    args = [url];
  }

  const child = spawn(command, args, {
    detached: true,
    stdio: 'ignore',
  });
  child.unref();
}

function respondHtml(res, statusCode, body) {
  res.writeHead(statusCode, {'content-type': 'text/html; charset=utf-8'});
  res.end(`<!doctype html><html><body style="font-family:sans-serif;padding:24px;"><p>${body}</p></body></html>`);
}

async function login(config, options = {}) {
  assertOauthConfig(config);
  const redirect = new URL(config.env.OUTLOOK_REDIRECT_URI);
  if (!/^https?:$/.test(redirect.protocol)) {
    throw new Error('OUTLOOK_REDIRECT_URI must be http:// or https://');
  }

  const state = crypto.randomBytes(24).toString('hex');
  const authorizeUrl = buildAuthorizeUrl(config, state);

  return new Promise((resolve, reject) => {
    let timeoutHandle = null;
    function finish(fn, value) {
      if (timeoutHandle) {
        clearTimeout(timeoutHandle);
        timeoutHandle = null;
      }
      fn(value);
    }

    const server = http.createServer(async (req, res) => {
      try {
        const incoming = new URL(req.url || '/', config.env.OUTLOOK_REDIRECT_URI);
        if (incoming.pathname !== redirect.pathname) {
          respondHtml(res, 404, 'Not found.');
          return;
        }

        const returnedState = incoming.searchParams.get('state') || '';
        const code = incoming.searchParams.get('code') || '';
        const error = incoming.searchParams.get('error') || '';
        const errorDescription = incoming.searchParams.get('error_description') || '';

        if (error) {
          respondHtml(res, 400, `Microsoft OAuth returned an error: ${error}${errorDescription ? ` (${errorDescription})` : ''}`);
          server.close();
          finish(reject, new Error(`Microsoft OAuth returned an error: ${error}${errorDescription ? ` (${errorDescription})` : ''}`));
          return;
        }

        if (returnedState !== state) {
          respondHtml(res, 400, 'OAuth state mismatch.');
          server.close();
          finish(reject, new Error('OAuth state mismatch.'));
          return;
        }

        if (!code) {
          respondHtml(res, 400, 'No OAuth code was returned.');
          server.close();
          finish(reject, new Error('No OAuth code was returned.'));
          return;
        }

        const persisted = await exchangeAuthorizationCode(config, code);
        respondHtml(res, 200, 'Microsoft OAuth complete. You can close this tab.');
        server.close();
        finish(resolve, {authorizeUrl, ...persisted});
      } catch (error) {
        try {
          respondHtml(res, 500, 'OAuth callback handling failed.');
        } catch {}
        server.close();
        finish(reject, error);
      }
    });

    server.on('error', (error) => finish(reject, error));
    server.listen(Number(redirect.port || (redirect.protocol === 'https:' ? 443 : 80)), redirect.hostname, () => {
      if (options.openBrowser) {
        try {
          openBrowser(authorizeUrl);
        } catch {}
      }
      console.log(`Listening for OAuth callback on ${config.env.OUTLOOK_REDIRECT_URI}`);
      console.log(`Authorize URL:\n${authorizeUrl}\n`);
      console.log('Complete the login in your browser. This process will exit after the callback is received.');
    });

    timeoutHandle = setTimeout(() => {
      server.close();
      finish(reject, new Error(`Timed out waiting for OAuth callback after ${options.timeoutSeconds} seconds.`));
    }, Math.max(30, Number(options.timeoutSeconds || 180)) * 1000);
    timeoutHandle.unref?.();
  });
}

function describeStatus(config) {
  if (config.env.OUTLOOK_USER_TOKEN) {
    return {
      hasTokens: true,
      authSource: 'OUTLOOK_USER_TOKEN',
      tokenFile: getTokenFilePath(config),
      redirectUri: config.env.OUTLOOK_REDIRECT_URI || '',
      scopes: getScopeString(config).split(/\s+/).filter(Boolean),
      hasRefreshToken: false,
    };
  }

  const tokenFile = getTokenFilePath(config);
  const token = readJsonFile(tokenFile, null, {strict: true});
  if (!token?.access_token) {
    return {
      hasTokens: false,
      authSource: 'token_file',
      tokenFile,
      redirectUri: config.env.OUTLOOK_REDIRECT_URI || '',
      scopes: getScopeString(config).split(/\s+/).filter(Boolean),
    };
  }

  const now = Date.now();
  return {
    hasTokens: true,
    authSource: 'token_file',
    tokenFile,
    redirectUri: config.env.OUTLOOK_REDIRECT_URI || '',
    scopes: `${token.scope || getScopeString(config)}`.split(/\s+/).filter(Boolean),
    expiresAt: token.expires_at || null,
    expiresInSeconds: token.expires_at ? Math.max(0, Math.floor((token.expires_at - now) / 1000)) : null,
    hasRefreshToken: Boolean(token.refresh_token),
  };
}

function clearToken(config) {
  const tokenFile = getTokenFilePath(config);
  ensurePrivateDir(path.dirname(tokenFile));
  if (fs.existsSync(tokenFile)) {
    fs.unlinkSync(tokenFile);
  }
  return {cleared: true, tokenFile};
}

async function main() {
  const args = parseArgs(process.argv);
  const config = loadConfig(args.envFile);

  switch (args.command) {
    case 'status':
      console.log(JSON.stringify(describeStatus(config), null, 2));
      return;
    case 'login': {
      const result = await login(config, args);
      console.log(JSON.stringify({
        success: true,
        tokenFile: result.tokenFile,
        expiresAt: result.token.expires_at || null,
        scope: result.token.scope || '',
      }, null, 2));
      return;
    }
    case 'refresh': {
      assertOauthConfig(config);
      const result = await refreshToken(config);
      console.log(JSON.stringify({
        success: true,
        tokenFile: result.tokenFile,
        expiresAt: result.token.expires_at || null,
        scope: result.token.scope || '',
      }, null, 2));
      return;
    }
    case 'clear':
      console.log(JSON.stringify(clearToken(config), null, 2));
      return;
    default:
      console.error(`Unknown command "${args.command}". Use: status | login | refresh | clear`);
      process.exitCode = 1;
  }
}

await main();
