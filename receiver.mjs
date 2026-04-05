#!/usr/bin/env bun

import fs from 'node:fs';
import http from 'node:http';
import process from 'node:process';
import {ensurePrivateDir} from './auth-store.mjs';
import {getNotificationLogPath, getReceiverConfig, loadConfig} from './lib.mjs';

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
