"""TTL cache with disk persistence.

Both supplier APIs have modest free-tier quotas, so every lookup served from
cache is one that stays inside the quota. Entries survive a restart because a
BOM is usually re-uploaded several times while a user fixes column mappings.
"""

import json
import os
import threading
import time
from collections import OrderedDict


class PartCache:
    def __init__(self, ttl_seconds=6 * 60 * 60, max_entries=5000, path=None):
        self.ttl_seconds = ttl_seconds
        self.max_entries = max_entries
        self.path = path
        self._lock = threading.Lock()
        self._entries = OrderedDict()
        if self.path:
            self._load()

    def _load(self):
        try:
            with open(self.path, 'r', encoding='utf-8') as handle:
                data = json.load(handle)
        except (OSError, ValueError):
            return
        if not isinstance(data, dict):
            return
        now = time.time()
        for key, entry in data.items():
            if not isinstance(entry, dict) or not isinstance(entry.get('storedAt'), (int, float)):
                continue
            if now - entry['storedAt'] > self.ttl_seconds:
                continue
            self._entries[key] = entry

    def get(self, key):
        with self._lock:
            entry = self._entries.get(key)
            if entry is None:
                return None
            if time.time() - entry['storedAt'] > self.ttl_seconds:
                del self._entries[key]
                return None
            # Refresh recency so the size trim evicts genuinely cold entries.
            self._entries.move_to_end(key)
            return entry['value']

    def set(self, key, value):
        with self._lock:
            self._entries[key] = {'storedAt': time.time(), 'value': value}
            self._entries.move_to_end(key)
            while len(self._entries) > self.max_entries:
                self._entries.popitem(last=False)
        self.flush()

    def clear(self):
        with self._lock:
            self._entries.clear()
        self.flush()

    def __len__(self):
        with self._lock:
            return len(self._entries)

    def flush(self):
        if not self.path:
            return
        with self._lock:
            snapshot = dict(self._entries)
        tmp = self.path + '.tmp'
        try:
            directory = os.path.dirname(self.path)
            if directory:
                os.makedirs(directory, exist_ok=True)
            with open(tmp, 'w', encoding='utf-8') as handle:
                json.dump(snapshot, handle)
            os.replace(tmp, self.path)
        except OSError:
            # A cache that cannot persist is still a working in-memory cache.
            pass
