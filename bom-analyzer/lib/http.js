'use strict';

class HttpError extends Error {
  constructor(message, status, body) {
    super(message);
    this.name = 'HttpError';
    this.status = status;
    this.body = body;
  }
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

// Supplier APIs throttle aggressively on free tiers, so 429 and 5xx get a
// backoff rather than being surfaced as a hard failure on the first try.
async function requestJson(url, options) {
  const opts = options || {};
  const timeoutMs = opts.timeoutMs || 20000;
  const retries = opts.retries === undefined ? 2 : opts.retries;
  const baseDelayMs = opts.baseDelayMs || 700;

  let lastError = null;
  for (let attempt = 0; attempt <= retries; attempt++) {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), timeoutMs);
    try {
      const res = await fetch(url, {
        method: opts.method || 'GET',
        headers: opts.headers,
        body: opts.body,
        signal: controller.signal,
      });
      const text = await res.text();
      let data = null;
      if (text) {
        try {
          data = JSON.parse(text);
        } catch (err) {
          data = null;
        }
      }

      if (res.ok) return { status: res.status, data, text };

      const retryable = res.status === 429 || res.status >= 500;
      const error = new HttpError(
        'HTTP ' + res.status + ' from ' + hostOf(url),
        res.status,
        data || truncate(text, 400)
      );
      if (!retryable || attempt === retries) throw error;
      lastError = error;
      const retryAfter = Number(res.headers.get('retry-after'));
      const delay = Number.isFinite(retryAfter) && retryAfter > 0
        ? Math.min(retryAfter * 1000, 10000)
        : baseDelayMs * Math.pow(2, attempt);
      await sleep(delay);
    } catch (err) {
      if (err instanceof HttpError) throw err;
      const message = err.name === 'AbortError'
        ? 'Request to ' + hostOf(url) + ' timed out after ' + timeoutMs + 'ms'
        : err.message || String(err);
      lastError = new HttpError(message, 0, null);
      if (attempt === retries) throw lastError;
      await sleep(baseDelayMs * Math.pow(2, attempt));
    } finally {
      clearTimeout(timer);
    }
  }
  throw lastError || new HttpError('Request failed', 0, null);
}

function hostOf(url) {
  try {
    return new URL(url).host;
  } catch (err) {
    return url;
  }
}

function truncate(text, max) {
  if (typeof text !== 'string') return text;
  return text.length > max ? text.slice(0, max) + '…' : text;
}

// Bounded parallelism keeps a 300-line BOM from opening 300 sockets at a
// supplier that would rate-limit us for it.
async function mapWithConcurrency(items, limit, worker) {
  const results = new Array(items.length);
  let cursor = 0;
  const size = Math.max(1, Math.min(limit, items.length || 1));

  async function run() {
    while (true) {
      const index = cursor++;
      if (index >= items.length) return;
      results[index] = await worker(items[index], index);
    }
  }

  await Promise.all(Array.from({ length: size }, run));
  return results;
}

module.exports = { requestJson, mapWithConcurrency, HttpError };
