'use strict';

const fs = require('fs');
const path = require('path');

// Both supplier APIs have modest free-tier quotas (DigiKey and Mouser each cap
// a free key in the low thousands of calls per day), so every lookup that can
// be served from cache is one that stays inside the quota. Entries survive a
// restart because a BOM is usually re-uploaded several times while a user is
// fixing column mappings.
class PartCache {
  constructor(options) {
    const opts = options || {};
    this.ttlMs = opts.ttlMs || 6 * 60 * 60 * 1000;
    this.maxEntries = opts.maxEntries || 5000;
    this.file = opts.file || null;
    this.map = new Map();
    this.dirty = false;
    this.flushTimer = null;
    if (this.file) this.load();
  }

  load() {
    let raw;
    try {
      raw = fs.readFileSync(this.file, 'utf8');
    } catch (err) {
      return;
    }
    let data;
    try {
      data = JSON.parse(raw);
    } catch (err) {
      return;
    }
    if (!data || typeof data !== 'object') return;
    const now = Date.now();
    for (const [key, entry] of Object.entries(data)) {
      if (!entry || typeof entry.storedAt !== 'number') continue;
      if (now - entry.storedAt > this.ttlMs) continue;
      this.map.set(key, entry);
    }
  }

  get(key) {
    const entry = this.map.get(key);
    if (!entry) return null;
    if (Date.now() - entry.storedAt > this.ttlMs) {
      this.map.delete(key);
      return null;
    }
    // Refresh recency so the size trim evicts genuinely cold entries.
    this.map.delete(key);
    this.map.set(key, entry);
    return entry.value;
  }

  set(key, value) {
    this.map.delete(key);
    this.map.set(key, { storedAt: Date.now(), value });
    while (this.map.size > this.maxEntries) {
      const oldest = this.map.keys().next();
      if (oldest.done) break;
      this.map.delete(oldest.value);
    }
    this.scheduleFlush();
  }

  clear() {
    this.map.clear();
    this.scheduleFlush();
  }

  get size() {
    return this.map.size;
  }

  scheduleFlush() {
    if (!this.file) return;
    this.dirty = true;
    if (this.flushTimer) return;
    this.flushTimer = setTimeout(() => {
      this.flushTimer = null;
      this.flush();
    }, 2000);
    if (this.flushTimer.unref) this.flushTimer.unref();
  }

  flush() {
    if (!this.file || !this.dirty) return;
    this.dirty = false;
    const obj = {};
    for (const [key, entry] of this.map) obj[key] = entry;
    const tmp = this.file + '.tmp';
    try {
      fs.mkdirSync(path.dirname(this.file), { recursive: true });
      fs.writeFileSync(tmp, JSON.stringify(obj));
      fs.renameSync(tmp, this.file);
    } catch (err) {
      // A cache that cannot persist is still a working in-memory cache.
    }
  }
}

module.exports = { PartCache };
