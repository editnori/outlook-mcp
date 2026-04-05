#!/usr/bin/env bun

import fs from 'node:fs';
import http from 'node:http';
import os from 'node:os';
import path from 'node:path';
import process from 'node:process';
import {ensurePrivateDir} from './auth-store.mjs';

const DEFAULT_STATE_DIR = path.join(
  process.env.XDG_STATE_HOME || path.join(os.homedir(), '.local', 'state'),
  'outlook-mcp'
);
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
    maxBodyBytes: Number(config.env.OUTLOOK_NOTIFICATION_MAX_BODY_BYTES || 262144),
    maxLogBytes: Number(config.env.OUTLOOK_NOTIFICATION_MAX_LOG_BYTES || 5242880),
    logFile: getNotificationLogPath(config),
  };
}

function writeNotification(config, record) {
  const logFile = getNotificationLogPath(config);
  ensurePrivateDir(path.dirname(logFile));
  const receiver = getReceiverConfig(config);
  if (fs.existsSync(logFile)) {
    const stat = fs.statSync(logFile);
    if (stat.size > receiver.maxLogBytes) {
      const bytesToKeep = Math.min(Math.floor(receiver.maxLogBytes / 2), stat.size);
      const fd = fs.openSync(logFile, 'r');
      const buffer = Buffer.alloc(bytesToKeep);
      fs.readSync(fd, buffer, 0, bytesToKeep, stat.size - bytesToKeep);
      fs.closeSync(fd);
      const lines = buffer.toString('utf8').split(/\r?\n/).filter(Boolean);
      const trimmed = lines.slice(-Math.max(1, Math.floor(lines.length / 2))).join('\n');
      fs.writeFileSync(logFile, `${trimmed}${trimmed ? '\n' : ''}`, {mode: 0o600});
    }
  }
  fs.appendFileSync(logFile, `${JSON.stringify(record)}\n`, {mode: 0o600});
  try {
    fs.chmodSync(logFile, 0o600);
  } catch {}
}

async function readBody(req, maxBytes) {
  const chunks = [];
  let total = 0;
  for await (const chunk of req) {
    total += chunk.length;
    if (total > maxBytes) {
      throw new Error(`Notification body exceeded ${maxBytes} bytes.`);
    }
    chunks.push(chunk);
  }
  return Buffer.concat(chunks).toString('utf8');
}

async function main() {
  const args = parseArgs(process.argv);
  const config = loadConfig(args.envFile);
  const receiver = getReceiverConfig(config);

  const server = http.createServer(async (req, res) => {
    try {
      const url = new URL(req.url || '/', `http://${req.headers.host || `${receiver.host}:${receiver.port}`}`);

      if (url.pathname === '/health') {
        res.writeHead(200, {'content-type': 'application/json'});
        res.end(JSON.stringify({ok: true, receiver}));
        return;
      }

      if (url.pathname !== receiver.path) {
        res.writeHead(404, {'content-type': 'text/plain; charset=utf-8'});
        res.end('not found');
        return;
      }

      const validationToken = url.searchParams.get('validationToken');
      if (validationToken) {
        res.writeHead(200, {'content-type': 'text/plain; charset=utf-8'});
        res.end(validationToken);
        return;
      }

      if ((req.method || '').toUpperCase() !== 'POST') {
        res.writeHead(405, {'content-type': 'text/plain; charset=utf-8'});
        res.end('method not allowed');
        return;
      }

      const bodyText = await readBody(req, receiver.maxBodyBytes);
      let payload = null;
      try {
        payload = bodyText ? JSON.parse(bodyText) : {};
      } catch {
        payload = {raw: bodyText};
      }

      if (
        receiver.clientState &&
        Array.isArray(payload?.value) &&
        payload.value.some((item) => item.clientState !== receiver.clientState)
      ) {
        res.writeHead(403, {'content-type': 'application/json'});
        res.end(JSON.stringify({accepted: false, error: 'clientState mismatch'}));
        return;
      }

      writeNotification(config, {
        receivedAt: new Date().toISOString(),
        method: req.method || '',
        path: url.pathname,
        query: Object.fromEntries(url.searchParams.entries()),
        headers: {
          'content-type': req.headers['content-type'] || '',
          'content-length': req.headers['content-length'] || '',
          'user-agent': req.headers['user-agent'] || '',
        },
        payload,
      });

      res.writeHead(202, {'content-type': 'application/json'});
      res.end(JSON.stringify({accepted: true}));
    } catch (error) {
      const status = /exceeded/.test(error.message || '') ? 413 : 500;
      res.writeHead(status, {'content-type': 'application/json'});
      res.end(JSON.stringify({accepted: false, error: error.message || String(error)}));
    }
  });

  server.listen(receiver.port, receiver.host, () => {
    console.log(JSON.stringify({
      ok: true,
      receiver,
      healthUrl: `http://${receiver.host}:${receiver.port}/health`,
    }, null, 2));
  });
}

main().catch((error) => {
  console.error(error.stack || error.message || String(error));
  process.exit(1);
});
