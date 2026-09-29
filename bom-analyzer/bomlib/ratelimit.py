"""Stay inside a supplier's published request limits.

TrustedParts meters in *parts*, not requests, over several windows at once:
50 parts in any 10 seconds, 150 in a minute, 2,000 in an hour, 20,000 in a day.
A batched request for 50 parts therefore spends a whole 10-second allowance in
one call, and the limiter has to know that before the call goes out rather than
after a 429 comes back.

Sliding windows, not fixed buckets. A fixed bucket lets a burst at 0:09 and
another at 0:11 through as two separate 10-second windows, which is twice the
allowance the supplier actually granted.

The clock and the sleep are injectable so the tests can drive a year of traffic
without waiting for it.
"""

import json
import os
import threading
import time as _time
from collections import deque


class RateLimited(Exception):
    """The request would breach a limit, and the wait is longer than allowed.

    Carries how long it would actually take, so the caller can say so rather
    than only that something went wrong.
    """

    def __init__(self, message, retry_after=None, window=None):
        Exception.__init__(self, message)
        self.retry_after = retry_after
        self.window = window


class SlidingWindowLimiter:
    """Several sliding windows over one stream of metered units.

    `windows` is [(seconds, allowance), ...]. Units are whatever the supplier
    counts — parts, for TrustedParts.
    """

    def __init__(self, windows, name='API', clock=None, sleeper=None, store=None):
        self.name = name
        # Shortest window first, so the error names the one that is really
        # biting rather than whichever happened to be checked first.
        self.windows = sorted(
            [(float(span), int(limit)) for span, limit in windows if span > 0 and limit > 0],
            key=lambda w: w[0],
        )
        self.clock = clock or _time.monotonic
        self.sleeper = sleeper or _time.sleep
        self.store = store
        self._lock = threading.Lock()
        self._events = deque()   # (timestamp, units), oldest first
        self._spent = 0
        if self.store:
            self._load()

    @property
    def longest(self):
        return self.windows[-1][0] if self.windows else 0.0

    @property
    def ceiling(self):
        """The most units any single request may ask for."""
        return min(limit for _, limit in self.windows) if self.windows else 0

    def _trim(self, now):
        cutoff = now - self.longest
        while self._events and self._events[0][0] <= cutoff:
            self._spent -= self._events[0][1]
            self._events.popleft()
        if not self._events:
            self._spent = 0

    def _wait_for(self, units, now):
        """Seconds until `units` would fit every window. 0.0 when it fits now."""
        wait = 0.0
        blocking = None
        for span, limit in self.windows:
            edge = now - span
            used = 0
            for timestamp, spent in self._events:
                if timestamp > edge:
                    used += spent
            room = limit - used
            if units <= room:
                continue
            # Let the oldest events in this window age out until there is room.
            needed = units - room
            freed = 0
            for timestamp, spent in self._events:
                if timestamp <= edge:
                    continue
                freed += spent
                if freed >= needed:
                    candidate = (timestamp + span) - now
                    if candidate > wait:
                        wait, blocking = candidate, (span, limit)
                    break
        return max(0.0, wait), blocking

    def reserve(self, units, max_wait=None, on_wait=None):
        """Claim `units`, waiting if that is what it takes.

        Waits up to `max_wait` seconds (None for as long as needed). Beyond
        that it raises RateLimited rather than blocking a run for an hour:
        telling somebody their daily allowance is gone is more useful than a
        progress bar that does not move until tomorrow.
        """
        units = int(units)
        if units <= 0 or not self.windows:
            return 0.0

        if units > self.ceiling:
            raise RateLimited(
                '%s allows at most %d parts in its shortest window, and this request '
                'asks for %d' % (self.name, self.ceiling, units),
                retry_after=None,
            )

        waited = 0.0
        while True:
            with self._lock:
                now = self.clock()
                self._trim(now)
                wait, blocking = self._wait_for(units, now)
                if wait <= 0:
                    self._events.append((now, units))
                    self._spent += units
                    if self.store:
                        self._save()
                    return waited

            if max_wait is not None and wait > max_wait:
                span, limit = blocking if blocking else (0, 0)
                raise RateLimited(
                    '%s limit reached: %d parts per %s. %s before the next %d can go out.'
                    % (self.name, limit, describe_span(span), describe_wait(wait), units),
                    retry_after=wait,
                    window=(span, limit),
                )

            if on_wait:
                on_wait(wait, blocking)
            # A margin, so the recheck lands after the event has really aged
            # out rather than a microsecond before it.
            self.sleeper(wait + 0.01)
            waited += wait

    def snapshot(self):
        """What each window has used right now, for a status endpoint."""
        with self._lock:
            now = self.clock()
            self._trim(now)
            out = []
            for span, limit in self.windows:
                edge = now - span
                used = sum(spent for timestamp, spent in self._events if timestamp > edge)
                out.append({
                    'seconds': span,
                    'window': describe_span(span),
                    'limit': limit,
                    'used': used,
                    'remaining': max(0, limit - used),
                })
            return out

    # ── Persistence ─────────────────────────────────────────────────────────
    #
    # The hour and day windows outlive the process. Without this, restarting
    # the server hands back a full daily allowance that the supplier has not,
    # and the limits stop being limits.

    def _load(self):
        try:
            with open(self.store, 'r', encoding='utf-8') as handle:
                data = json.load(handle)
        except (OSError, ValueError):
            return
        if not isinstance(data, dict):
            return
        # Stored against the wall clock, since a monotonic clock restarts with
        # the process; converted back into this run's clock on the way in.
        now_wall = _time.time()
        now = self.clock()
        cutoff = now_wall - self.longest
        for entry in data.get('events') or []:
            try:
                stamp, units = float(entry[0]), int(entry[1])
            except (TypeError, ValueError, IndexError):
                continue
            if stamp <= cutoff or units <= 0:
                continue
            self._events.append((now - (now_wall - stamp), units))
            self._spent += units
        self._events = deque(sorted(self._events, key=lambda e: e[0]))

    def _save(self):
        now = self.clock()
        now_wall = _time.time()
        payload = {'events': [[now_wall - (now - stamp), units]
                              for stamp, units in self._events]}
        try:
            directory = os.path.dirname(self.store)
            if directory:
                os.makedirs(directory, exist_ok=True)
            temporary = self.store + '.tmp'
            with open(temporary, 'w', encoding='utf-8') as handle:
                json.dump(payload, handle)
            os.replace(temporary, self.store)
        except OSError:
            # A ledger that cannot be written is worth less than a run that
            # cannot proceed; the in-memory windows still hold for this run.
            pass


def describe_span(seconds):
    seconds = float(seconds or 0)
    if seconds >= 86400:
        return '%g hours' % (seconds / 3600) if seconds != 86400 else '24 hours'
    if seconds >= 3600:
        return '%g hour%s' % (seconds / 3600, '' if seconds == 3600 else 's')
    if seconds >= 60:
        return '%g minute%s' % (seconds / 60, '' if seconds == 60 else 's')
    return '%g seconds' % seconds


def describe_wait(seconds):
    seconds = float(seconds or 0)
    if seconds < 90:
        return '%d seconds' % max(1, round(seconds))
    if seconds < 5400:
        return '%d minutes' % round(seconds / 60)
    return '%.1f hours' % (seconds / 3600)


def parse_windows(text, fallback):
    """Read "10:50,60:150" into [(10.0, 50), ...]; fallback when unreadable."""
    windows = []
    for chunk in str(text or '').split(','):
        chunk = chunk.strip()
        if not chunk:
            continue
        span, _, limit = chunk.partition(':')
        try:
            windows.append((float(span.strip()), int(limit.strip())))
        except ValueError:
            return list(fallback)
    return windows or list(fallback)
