'use strict';

const http = require('http');
const fs = require('fs');
const path = require('path');

const { loadEnv } = require('./lib/env');
loadEnv();

const { PartCache } = require('./lib/cache');
const { DigiKeyClient } = require('./lib/digikey');
const { MouserClient } = require('./lib/mouser');
const { LookupService, summarizeBom } = require('./lib/lookup');
const { parseWorkbook, extractBom, lineFromRow } = require('./lib/spreadsheet');

const PUBLIC_DIR = path.join(__dirname, 'public');
const PORT = Number(process.env.PORT) || 8787;
const HOST = process.env.HOST || '0.0.0.0';
const MAX_UPLOAD_BYTES = Number(process.env.MAX_UPLOAD_BYTES) || 12 * 1024 * 1024;
const MAX_JSON_BYTES = 4 * 1024 * 1024;
const MAX_PARTS_PER_REQUEST = Number(process.env.MAX_PARTS_PER_REQUEST) || 500;
const ALLOWED_ORIGINS = String(process.env.ALLOWED_ORIGINS || '*')
  .split(',')
  .map((o) => o.trim())
  .filter(Boolean);

const digikey = new DigiKeyClient({
  clientId: process.env.DIGIKEY_CLIENT_ID,
  clientSecret: process.env.DIGIKEY_CLIENT_SECRET,
  sandbox: /^(1|true|yes)$/i.test(String(process.env.DIGIKEY_SANDBOX || '')),
  site: process.env.DIGIKEY_SITE,
  language: process.env.DIGIKEY_LANGUAGE,
  currency: process.env.DIGIKEY_CURRENCY,
});

const mouser = new MouserClient({
  apiKey: process.env.MOUSER_API_KEY,
  currency: process.env.MOUSER_CURRENCY,
});

const cache = new PartCache({
  ttlMs: (Number(process.env.CACHE_TTL_HOURS) || 6) * 60 * 60 * 1000,
  file: process.env.CACHE_FILE === 'none'
    ? null
    : process.env.CACHE_FILE || path.join(__dirname, '.cache', 'parts.json'),
});

const lookupService = new LookupService({
  clients: [digikey, mouser],
  cache,
  concurrency: Number(process.env.LOOKUP_CONCURRENCY) || 3,
});

const MIME_TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.ico': 'image/x-icon',
  '.csv': 'text/csv; charset=utf-8',
  '.webmanifest': 'application/manifest+json',
};

const server = http.createServer((req, res) => {
  handle(req, res).catch((err) => {
    console.error('Unhandled error:', err);
    if (!res.headersSent) sendJson(req, res, 500, { error: 'Internal server error' });
    else res.end();
  });
});

async function handle(req, res) {
  const url = new URL(req.url, 'http://' + (req.headers.host || 'localhost'));

  if (req.method === 'OPTIONS') {
    applyCors(req, res);
    res.writeHead(204);
    res.end();
    return;
  }

  if (url.pathname.startsWith('/api/')) {
    applyCors(req, res);
    return handleApi(req, res, url);
  }

  if (req.method !== 'GET' && req.method !== 'HEAD') {
    return sendJson(req, res, 405, { error: 'Method not allowed' });
  }
  return serveStatic(req, res, url.pathname);
}

async function handleApi(req, res, url) {
  const route = url.pathname;

  if (route === '/api/health' && req.method === 'GET') {
    return sendJson(req, res, 200, {
      ok: true,
      suppliers: [
        { id: 'digikey', name: 'DigiKey', configured: digikey.configured, sandbox: digikey.sandbox },
        { id: 'mouser', name: 'Mouser', configured: mouser.configured },
      ],
      maxPartsPerRequest: MAX_PARTS_PER_REQUEST,
      cacheEntries: cache.size,
      currency: process.env.DIGIKEY_CURRENCY || process.env.MOUSER_CURRENCY || 'USD',
    });
  }

  if (route === '/api/parse' && req.method === 'POST') {
    const body = await readBody(req, MAX_UPLOAD_BYTES);
    if (body === null) return sendJson(req, res, 413, { error: 'File is too large' });
    const filename = String(req.headers['x-file-name'] || 'bom.csv');
    try {
      const grid = parseWorkbook(body, filename);
      const parsed = extractBom(grid);
      return sendJson(req, res, 200, {
        filename,
        headerRow: parsed.headerRow,
        headers: parsed.headers,
        mapping: parsed.mapping,
        lines: parsed.lines,
        skipped: parsed.skipped,
        totalRows: parsed.totalRows,
        // The raw grid lets the UI re-derive lines when the user corrects a
        // column mapping, without a second upload. rowOffset keeps the row
        // numbers on screen pointing at the original spreadsheet rows.
        rows: grid.slice(parsed.headerRow + 1, parsed.headerRow + 1 + 5000),
        rowOffset: parsed.headerRow + 1,
      });
    } catch (err) {
      return sendJson(req, res, 400, { error: 'Could not read that file: ' + err.message });
    }
  }

  if (route === '/api/remap' && req.method === 'POST') {
    const payload = await readJson(req, res);
    if (!payload) return;
    const rows = Array.isArray(payload.rows) ? payload.rows : [];
    const mapping = payload.mapping || {};
    const offset = Number.isFinite(payload.rowOffset) ? payload.rowOffset : 0;
    const lines = [];
    let skipped = 0;
    rows.forEach((row, index) => {
      if (!Array.isArray(row) || row.every((cell) => !String(cell || '').trim())) return;
      const line = lineFromRow(row, mapping, offset + index);
      if (!line.mpn) {
        skipped++;
        return;
      }
      lines.push(line);
    });
    return sendJson(req, res, 200, { lines, skipped });
  }

  if (route === '/api/lookup' && req.method === 'POST') {
    const payload = await readJson(req, res);
    if (!payload) return;
    const parts = Array.isArray(payload.parts) ? payload.parts : [];
    if (parts.length === 0) return sendJson(req, res, 400, { error: 'No parts supplied' });
    if (parts.length > MAX_PARTS_PER_REQUEST) {
      return sendJson(req, res, 400, {
        error: 'Send at most ' + MAX_PARTS_PER_REQUEST + ' parts per request',
      });
    }

    const clean = parts
      .map((p, i) => ({
        row: Number.isFinite(p.row) ? p.row : i + 1,
        mpn: String(p.mpn || '').trim(),
        quantity: Number.isFinite(p.quantity) && p.quantity > 0 ? Math.trunc(p.quantity) : 1,
        manufacturer: p.manufacturer ? String(p.manufacturer).trim() : null,
        reference: p.reference ? String(p.reference).trim() : null,
        description: p.description ? String(p.description).trim() : null,
      }))
      .filter((p) => p.mpn);

    if (clean.length === 0) return sendJson(req, res, 400, { error: 'No usable part numbers supplied' });

    // A large BOM takes a while, so the client can ask for server-sent events
    // and watch progress instead of staring at a spinner.
    if (payload.stream) return streamLookup(req, res, clean);

    try {
      const result = await lookupService.lookupParts(clean);
      return sendJson(req, res, 200, {
        rows: result.rows,
        suppliers: result.suppliers,
        stats: result.stats,
        summary: summarizeBom(result.rows, result.suppliers),
      });
    } catch (err) {
      console.error('Lookup failed:', err);
      return sendJson(req, res, 502, { error: err.message || 'Supplier lookup failed' });
    }
  }

  if (route === '/api/cache' && req.method === 'DELETE') {
    cache.clear();
    return sendJson(req, res, 200, { ok: true, cacheEntries: cache.size });
  }

  return sendJson(req, res, 404, { error: 'Unknown endpoint ' + route });
}

async function streamLookup(req, res, parts) {
  res.writeHead(200, {
    'Content-Type': 'text/event-stream; charset=utf-8',
    'Cache-Control': 'no-cache, no-transform',
    Connection: 'keep-alive',
    'X-Accel-Buffering': 'no',
  });

  const send = (event, data) => {
    if (res.writableEnded) return;
    res.write('event: ' + event + '\ndata: ' + JSON.stringify(data) + '\n\n');
  };

  let cancelled = false;
  req.on('close', () => {
    cancelled = true;
  });

  send('start', { parts: parts.length, suppliers: lookupService.suppliers });

  // Progress fires once per supplier lookup, which is faster than the client
  // can usefully repaint; throttle to a readable rate.
  let lastSent = 0;
  const onProgress = (progress) => {
    if (cancelled) return;
    const now = Date.now();
    if (progress.completed === progress.total || now - lastSent > 120) {
      lastSent = now;
      send('progress', progress);
    }
  };

  try {
    const result = await lookupService.lookupParts(parts, { onProgress });
    send('done', {
      rows: result.rows,
      suppliers: result.suppliers,
      stats: result.stats,
      summary: summarizeBom(result.rows, result.suppliers),
    });
  } catch (err) {
    console.error('Lookup failed:', err);
    send('error', { error: err.message || 'Supplier lookup failed' });
  } finally {
    if (!res.writableEnded) res.end();
  }
}

function applyCors(req, res) {
  const origin = req.headers.origin;
  if (ALLOWED_ORIGINS.includes('*')) {
    res.setHeader('Access-Control-Allow-Origin', '*');
  } else if (origin && ALLOWED_ORIGINS.includes(origin)) {
    res.setHeader('Access-Control-Allow-Origin', origin);
    res.setHeader('Vary', 'Origin');
  }
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, DELETE, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, X-File-Name');
  res.setHeader('Access-Control-Max-Age', '86400');
}

function readBody(req, limit) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let size = 0;
    let aborted = false;
    req.on('data', (chunk) => {
      if (aborted) return;
      size += chunk.length;
      if (size > limit) {
        aborted = true;
        resolve(null);
        req.destroy();
        return;
      }
      chunks.push(chunk);
    });
    req.on('end', () => {
      if (!aborted) resolve(Buffer.concat(chunks));
    });
    req.on('error', (err) => {
      if (!aborted) reject(err);
    });
  });
}

async function readJson(req, res) {
  const body = await readBody(req, MAX_JSON_BYTES);
  if (body === null) {
    sendJson(req, res, 413, { error: 'Request body is too large' });
    return null;
  }
  try {
    return JSON.parse(body.toString('utf8') || '{}');
  } catch (err) {
    sendJson(req, res, 400, { error: 'Request body is not valid JSON' });
    return null;
  }
}

function sendJson(req, res, status, payload) {
  const body = JSON.stringify(payload);
  res.writeHead(status, {
    'Content-Type': 'application/json; charset=utf-8',
    'Content-Length': Buffer.byteLength(body),
    'Cache-Control': 'no-store',
  });
  if (req.method === 'HEAD') res.end();
  else res.end(body);
}

function serveStatic(req, res, pathname) {
  const relative = pathname === '/' ? 'index.html' : pathname.replace(/^\/+/, '');
  const target = path.join(PUBLIC_DIR, relative);
  // path.join collapses "..", so confirm the result is still inside PUBLIC_DIR
  // before reading anything off disk.
  if (!target.startsWith(PUBLIC_DIR + path.sep) && target !== PUBLIC_DIR) {
    return sendJson(req, res, 403, { error: 'Forbidden' });
  }

  fs.stat(target, (err, stats) => {
    if (err || !stats.isFile()) {
      res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
      res.end('Not found');
      return;
    }
    const ext = path.extname(target).toLowerCase();
    res.writeHead(200, {
      'Content-Type': MIME_TYPES[ext] || 'application/octet-stream',
      'Content-Length': stats.size,
      'Cache-Control': ext === '.html' ? 'no-cache' : 'public, max-age=3600',
    });
    if (req.method === 'HEAD') {
      res.end();
      return;
    }
    fs.createReadStream(target).pipe(res);
  });
}

function shutdown() {
  cache.flush();
  server.close(() => process.exit(0));
  setTimeout(() => process.exit(0), 3000).unref();
}

process.on('SIGINT', shutdown);
process.on('SIGTERM', shutdown);

if (require.main === module) {
  server.listen(PORT, HOST, () => {
    const configured = [digikey, mouser].filter((c) => c.configured).map((c) => c.name);
    console.log('BOM Supplier Analyzer listening on http://localhost:' + PORT);
    console.log(
      configured.length > 0
        ? 'Suppliers configured: ' + configured.join(', ') + (digikey.sandbox ? ' (DigiKey sandbox)' : '')
        : 'No supplier credentials found — copy .env.example to .env and add your API keys.'
    );
  });
}

module.exports = { server, lookupService, cache };
