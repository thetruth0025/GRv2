"""Staying inside TrustedParts' published limits.

They meter in parts across four windows at once. The clock and the sleep are
injected here so a day of traffic runs in milliseconds.
"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from bomlib.ratelimit import (  # noqa: E402
    RateLimited,
    SlidingWindowLimiter,
    describe_span,
    describe_wait,
    parse_windows,
)

TRUSTEDPARTS = [(10, 50), (60, 150), (3600, 2000), (86400, 20000)]


class FakeClock:
    """A clock that only moves when something sleeps on it."""

    def __init__(self):
        self.now = 1000.0
        self.slept = []

    def __call__(self):
        return self.now

    def sleep(self, seconds):
        self.slept.append(seconds)
        self.now += seconds

    def advance(self, seconds):
        self.now += seconds


def limiter(windows=None, clock=None):
    clock = clock or FakeClock()
    return SlidingWindowLimiter(windows or TRUSTEDPARTS, name='TrustedParts',
                                clock=clock, sleeper=clock.sleep), clock


class WindowTests(unittest.TestCase):
    def test_a_request_inside_every_allowance_goes_straight_out(self):
        rate, clock = limiter()
        self.assertEqual(rate.reserve(50), 0.0)
        self.assertEqual(clock.slept, [])

    def test_the_ten_second_allowance_is_spent_by_one_full_batch(self):
        rate, clock = limiter()
        rate.reserve(50)
        rate.reserve(50)
        # The second had to wait out the first window rather than going now.
        self.assertGreater(sum(clock.slept), 9.9)
        self.assertLess(sum(clock.slept), 10.2)

    def test_waiting_is_only_as_long_as_it_has_to_be(self):
        rate, clock = limiter()
        rate.reserve(40)
        clock.advance(6)
        # 40 already spent, so 10 more fit now and 20 need 4 more seconds.
        self.assertEqual(rate.reserve(10), 0.0)
        rate.reserve(20)
        self.assertAlmostEqual(sum(clock.slept), 4.01, places=2)

    def test_units_age_out_of_a_window(self):
        rate, clock = limiter()
        rate.reserve(50)
        clock.advance(11)
        self.assertEqual(rate.reserve(50), 0.0)
        self.assertEqual(clock.slept, [])

    def test_the_minute_allowance_binds_after_three_batches(self):
        rate, clock = limiter()
        for _ in range(3):
            rate.reserve(50)
            clock.advance(10)
        # 150 parts inside the minute: the fourth waits for the minute, not
        # the ten seconds.
        rate.reserve(50)
        self.assertGreater(sum(clock.slept), 25)

    def test_no_trailing_window_ever_holds_more_than_its_allowance(self):
        """The invariant, checked continuously rather than at the end.

        Totalling what went out across a fixed hour is the wrong test: a
        sliding window legitimately passes more than 2,000 over a clock hour,
        because the earliest spending has aged out by the end of it. What must
        never happen is a trailing window holding more than its limit.
        """
        rate, clock = limiter()
        issued = 0
        while clock.now < 1000.0 + 4 * 3600:
            try:
                rate.reserve(50, max_wait=300)
                issued += 50
            except RateLimited:
                clock.advance(60)
            for window in rate.snapshot():
                self.assertLessEqual(window['used'], window['limit'],
                                     'breached %s' % window['window'])
        # Four hours of pushing as hard as the limits allow: near the hourly
        # allowance every hour, and never past it.
        self.assertGreater(issued, 6000)

    def test_a_day_of_traffic_stays_inside_the_daily_allowance(self):
        rate, clock = limiter()
        while clock.now < 1000.0 + 86400:
            try:
                rate.reserve(50, max_wait=300)
            except RateLimited:
                clock.advance(300)
            for window in rate.snapshot():
                self.assertLessEqual(window['used'], window['limit'])

    def test_the_window_edge_is_exclusive_so_spending_ages_out_exactly(self):
        rate, clock = limiter([(10, 50)])
        rate.reserve(50)
        clock.advance(10)
        # At exactly ten seconds the earlier spending is outside the window.
        self.assertEqual(rate.snapshot()[0]['used'], 0)
        self.assertEqual(rate.reserve(50, max_wait=0), 0.0)

    def test_a_long_wait_is_refused_rather_than_blocking_the_run(self):
        rate, clock = limiter([(3600, 100)])
        rate.reserve(100)
        with self.assertRaises(RateLimited) as caught:
            rate.reserve(50, max_wait=60)
        self.assertGreater(caught.exception.retry_after, 3000)
        self.assertEqual(caught.exception.window, (3600.0, 100))
        self.assertIn('TrustedParts limit reached', str(caught.exception))
        self.assertEqual(clock.slept, [])

    def test_a_request_larger_than_any_window_is_refused_immediately(self):
        rate, _ = limiter()
        with self.assertRaises(RateLimited) as caught:
            rate.reserve(51)
        self.assertIn('at most 50 parts', str(caught.exception))
        self.assertIsNone(caught.exception.retry_after)

    def test_nothing_is_reserved_for_an_empty_request(self):
        rate, clock = limiter()
        self.assertEqual(rate.reserve(0), 0.0)
        self.assertEqual(rate.snapshot()[0]['used'], 0)

    def test_the_caller_is_told_it_is_waiting(self):
        rate, clock = limiter()
        seen = []
        rate.reserve(50)
        rate.reserve(50, on_wait=lambda seconds, window: seen.append((round(seconds), window)))
        self.assertEqual(len(seen), 1)
        self.assertEqual(seen[0][1], (10.0, 50))

    def test_a_snapshot_says_what_is_left_in_each_window(self):
        rate, clock = limiter()
        rate.reserve(30)
        snapshot = rate.snapshot()
        self.assertEqual([w['limit'] for w in snapshot], [50, 150, 2000, 20000])
        self.assertEqual([w['used'] for w in snapshot], [30, 30, 30, 30])
        self.assertEqual(snapshot[0]['remaining'], 20)
        clock.advance(11)
        self.assertEqual(rate.snapshot()[0]['used'], 0)
        self.assertEqual(rate.snapshot()[1]['used'], 30)

    def test_concurrent_callers_share_one_allowance(self):
        import threading
        rate = SlidingWindowLimiter([(10, 50)], name='TrustedParts')
        granted = []
        errors = []

        def draw():
            try:
                rate.reserve(10, max_wait=0)
                granted.append(1)
            except RateLimited:
                errors.append(1)

        threads = [threading.Thread(target=draw) for _ in range(12)]
        for thread in threads:
            thread.start()
        for thread in threads:
            thread.join()
        # Five of twelve fit; the rest are turned away rather than overrunning.
        self.assertEqual(len(granted), 5)
        self.assertEqual(len(errors), 7)


class PersistenceTests(unittest.TestCase):
    """An hourly limit that a restart resets is not a limit."""

    def setUp(self):
        import tempfile
        handle, self.path = tempfile.mkstemp(suffix='.json')
        os.close(handle)
        os.unlink(self.path)

    def tearDown(self):
        for path in (self.path, self.path + '.tmp'):
            if os.path.exists(path):
                os.unlink(path)

    def test_spending_survives_a_restart(self):
        first = SlidingWindowLimiter([(3600, 100)], store=self.path)
        first.reserve(80)
        second = SlidingWindowLimiter([(3600, 100)], store=self.path)
        self.assertEqual(second.snapshot()[0]['used'], 80)
        with self.assertRaises(RateLimited):
            second.reserve(50, max_wait=0)

    def test_spending_older_than_the_longest_window_is_forgotten(self):
        import json
        import time
        with open(self.path, 'w', encoding='utf-8') as handle:
            json.dump({'events': [[time.time() - 7200, 80]]}, handle)
        rate = SlidingWindowLimiter([(3600, 100)], store=self.path)
        self.assertEqual(rate.snapshot()[0]['used'], 0)

    def test_an_unreadable_ledger_does_not_stop_a_run(self):
        with open(self.path, 'w', encoding='utf-8') as handle:
            handle.write('{ not json')
        rate = SlidingWindowLimiter([(3600, 100)], store=self.path)
        self.assertEqual(rate.reserve(10), 0.0)

    def test_an_unwritable_ledger_does_not_stop_a_run_either(self):
        rate = SlidingWindowLimiter([(3600, 100)], store='/proc/nope/ledger.json')
        self.assertEqual(rate.reserve(10), 0.0)


class WordingTests(unittest.TestCase):
    def test_spans_read_the_way_the_supplier_states_them(self):
        self.assertEqual(describe_span(10), '10 seconds')
        self.assertEqual(describe_span(60), '1 minute')
        self.assertEqual(describe_span(3600), '1 hour')
        self.assertEqual(describe_span(86400), '24 hours')

    def test_waits_are_said_in_a_unit_somebody_can_act_on(self):
        self.assertEqual(describe_wait(4.2), '4 seconds')
        self.assertEqual(describe_wait(200), '3 minutes')
        self.assertEqual(describe_wait(7200), '2.0 hours')

    def test_windows_can_be_overridden_without_a_code_change(self):
        self.assertEqual(parse_windows('10:50,60:150', []), [(10.0, 50), (60.0, 150)])

    def test_a_malformed_override_falls_back_rather_than_removing_the_limit(self):
        self.assertEqual(parse_windows('nonsense', TRUSTEDPARTS), TRUSTEDPARTS)
        self.assertEqual(parse_windows('', TRUSTEDPARTS), TRUSTEDPARTS)


if __name__ == '__main__':
    unittest.main()


class ClientTests(unittest.TestCase):
    """The limiter where it actually sits: in front of the HTTP call."""

    def client(self, windows=None, max_wait=None):
        from bomlib import trustedparts
        from bomlib.trustedparts import TrustedPartsClient
        clock = FakeClock()
        rate = SlidingWindowLimiter(windows or TRUSTEDPARTS, name='TrustedParts',
                                    clock=clock, sleeper=clock.sleep)
        client = TrustedPartsClient(api_key='k', limiter=rate, max_wait=max_wait)

        sent = []
        original = trustedparts.request_json

        def fake(url, method='GET', headers=None, body=None, **kwargs):
            import json as _json
            sent.append(len(_json.loads(body)['Queries']))
            return {'status': 200, 'data': {'PartResults': []}}

        trustedparts.request_json = fake
        self.addCleanup(lambda: setattr(trustedparts, 'request_json', original))
        return client, clock, sent

    def parts(self, count, offset=0):
        return [{'mpn': 'PART-%04d' % (offset + i), 'quantity': 1} for i in range(count)]

    def test_a_request_reserves_one_unit_per_part_not_one_per_call(self):
        client, clock, sent = self.client()
        client.search(self.parts(30))
        self.assertEqual(sent, [30])
        self.assertEqual(client.limiter.snapshot()[0]['used'], 30)

    def test_a_second_full_batch_waits_out_the_ten_second_window(self):
        client, clock, sent = self.client()
        client.search(self.parts(50))
        client.search(self.parts(50, offset=50))
        self.assertEqual(sent, [50, 50])
        self.assertGreater(sum(clock.slept), 9.9)

    def test_parts_too_short_to_search_do_not_spend_the_allowance(self):
        client, clock, sent = self.client()
        client.search([{'mpn': 'A'}, {'mpn': ''}, {'mpn': 'REAL-PART'}])
        self.assertEqual(sent, [1])
        self.assertEqual(client.limiter.snapshot()[0]['used'], 1)

    def test_nothing_searchable_spends_nothing_and_sends_nothing(self):
        client, clock, sent = self.client()
        self.assertEqual(client.search([{'mpn': 'A'}]), {})
        self.assertEqual(sent, [])
        self.assertEqual(client.limiter.snapshot()[0]['used'], 0)

    def test_the_batch_size_never_exceeds_the_tightest_window(self):
        from bomlib.trustedparts import MAX_QUERIES_PER_REQUEST
        client, _, _ = self.client()
        self.assertEqual(client.batch_size, 50)
        tight, _, _2 = self.client(windows=[(10, 20), (60, 150)])
        self.assertEqual(tight.batch_size, 20)
        self.assertLessEqual(tight.batch_size, MAX_QUERIES_PER_REQUEST)

    def test_a_long_wait_is_reported_rather_than_held(self):
        client, clock, sent = self.client(windows=[(3600, 50)], max_wait=30)
        client.search(self.parts(50))
        with self.assertRaises(RateLimited) as caught:
            client.search(self.parts(50, offset=50))
        self.assertEqual(sent, [50])
        self.assertIn('TrustedParts limit reached', str(caught.exception))
        self.assertIn('50 parts per 1 hour', str(caught.exception))

    def test_the_wait_is_announced_to_whoever_is_watching(self):
        client, clock, sent = self.client()
        heard = []
        client.on_wait = lambda seconds, message: heard.append(message)
        client.search(self.parts(50))
        client.search(self.parts(50, offset=50))
        self.assertEqual(len(heard), 1)
        self.assertIn('TrustedParts allows 50 parts per 10 seconds', heard[0])
        self.assertIn('waiting', heard[0])


class LookupIntegrationTests(unittest.TestCase):
    """Rate limiting through the service, where a cache hit must cost nothing."""

    def service(self, windows, max_wait=0):
        from bomlib.cache import PartCache
        from bomlib.lookup import LookupService
        from bomlib.trustedparts import TrustedPartsClient

        clock = FakeClock()
        rate = SlidingWindowLimiter(windows, name='TrustedParts',
                                    clock=clock, sleeper=clock.sleep)
        client = TrustedPartsClient(api_key='k', limiter=rate, max_wait=max_wait)

        from bomlib import trustedparts
        original = trustedparts.request_json
        sent = []

        def fake(url, method='GET', headers=None, body=None, **kwargs):
            import json as _json
            queries = _json.loads(body)['Queries']
            sent.append(len(queries))
            return {'status': 200, 'data': {'PartResults': [{
                'PartNumber': q['SearchToken'], 'Manufacturer': 'Acme',
                'Distributors': [{'Name': 'Arrow', 'DistributorResults': [{
                    'DistributorPartNumber': 'D-' + q['SearchToken'],
                    'Stock': {'Quantity': 500},
                    'Pricing': {'MinimumQuantity': 1, 'CurrencyCode': 'USD',
                                'PriceBreaks': [{'Quantity': 1, 'Price': 1.0}]},
                }]}],
            } for q in queries]}}

        trustedparts.request_json = fake
        self.addCleanup(lambda: setattr(trustedparts, 'request_json', original))
        return LookupService(clients=[client], cache=PartCache(ttl_seconds=600, path=None),
                             concurrency=1, include_alternates=False), sent, clock

    def parts(self, count):
        return [{'row': i + 1, 'mpn': 'PART-%04d' % i, 'quantity': 1} for i in range(count)]

    def test_a_cached_part_costs_no_allowance(self):
        service, sent, _ = self.service([(3600, 60)])
        service.lookup_parts(self.parts(40))
        limiter = service.clients[0].limiter
        self.assertEqual(limiter.snapshot()[0]['used'], 40)

        # The same BOM again: served from cache, so nothing is spent and the
        # remaining allowance is untouched.
        service.lookup_parts(self.parts(40))
        self.assertEqual(sent, [40])
        self.assertEqual(limiter.snapshot()[0]['used'], 40)

    def test_parts_beyond_the_allowance_are_reported_not_silently_dropped(self):
        service, sent, _ = self.service([(3600, 60)], max_wait=0)
        result = service.lookup_parts(self.parts(100))
        reasons = [row['offers']['trustedparts'].get('reason') for row in result['rows']]
        answered = [r for r in reasons if not r]
        refused = [r for r in reasons if r and 'limit reached' in r]
        self.assertEqual(len(answered), 50)
        self.assertEqual(len(refused), 50)
        # And the refusal says when it could be retried.
        self.assertIn('minutes', refused[0])

    def test_a_refusal_is_not_cached_so_the_next_run_retries_it(self):
        service, sent, clock = self.service([(3600, 60)], max_wait=0)
        service.lookup_parts(self.parts(100))
        first = len(sent)
        clock.advance(3601)
        result = service.lookup_parts(self.parts(100))
        self.assertGreater(len(sent), first)
        found = [row for row in result['rows']
                 if row['offers']['trustedparts'].get('found')]
        self.assertEqual(len(found), 100)
