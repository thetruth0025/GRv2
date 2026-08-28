"""Run a BOM against every configured supplier and pair the answers up."""

import threading
from concurrent.futures import ThreadPoolExecutor

from .normalize import (
    Lifecycle,
    NoMatch,
    compare_offers,
    error_offer,
    missing_offer,
    record_to_offer,
    round_to,
)


def _entry_for(record):
    """A cache entry for whatever the client returned.

    A NoMatch is not a record but it is not nothing either: it says the search
    ran and the requested part was not among the answers, and it carries what
    came back instead. Stored as plain data so it survives the cache.
    """
    if isinstance(record, NoMatch):
        return {'notFound': True, 'miss': record.as_dict()}
    return {'record': record} if record else {'notFound': True}


def _miss_reason(entry, supplier, mpn):
    miss = entry.get('miss') if isinstance(entry, dict) else None
    if isinstance(miss, dict):
        return NoMatch(**miss).describe(supplier, mpn)
    return 'No %s match for %s' % (supplier, mpn)


def summarize_alternate(part, offers, offer_list):
    """One approved alternate, reduced to the question being asked of it.

    A buyer looking at an alternates column wants one thing first: could this
    cover the line if the primary could not. That is stock and lifecycle, so
    those lead; the full per-supplier detail rides along for the row that gets
    expanded.
    """
    comparison = compare_offers(offer_list, part.get('quantity'))
    found = [offer for offer in offer_list if offer and offer.get('found')]
    covers = bool(comparison.get('stockCovers'))
    healthy = comparison.get('lifecycleSeverity') not in ('bad',)

    return {
        'mpn': part.get('mpn'),
        'quantity': part.get('quantity'),
        'offers': offers,
        'comparison': comparison,
        'found': bool(found),
        # The headline: stocked somewhere, and not itself on the way out.
        'usable': bool(found) and covers and healthy,
        'coversQuantity': covers,
        'lifecycle': comparison.get('lifecycle'),
        'lifecycleSeverity': comparison.get('lifecycleSeverity'),
        'bestPrice': comparison.get('bestPrice'),
        'bestPriceSupplier': comparison.get('bestPriceSupplier'),
        'stock': sum(
            offer['stock'] for offer in found
            if isinstance(offer.get('stock'), (int, float))
        ) or None,
    }


def alternate_parts(part):
    """The BOM's approved alternates, shaped like parts so they can be looked up.

    They inherit the line's quantity, because the question an alternate answers
    is "could this cover the same build", which is a question about the same
    number of pieces. They do not inherit the manufacturer: an alternate is
    usually a second source, so a different one.
    """
    found = []
    for mpn in part.get('alternates') or []:
        mpn = str(mpn or '').strip()
        if not mpn:
            continue
        found.append({
            'mpn': mpn,
            'quantity': part.get('quantity') or 1,
            'manufacturer': None,
            'reference': part.get('reference'),
            'description': None,
        })
    return found


def cache_key(supplier_id, part):
    mpn = ' '.join(str(part.get('mpn') or '').upper().split())
    mfr = ' '.join(str(part.get('manufacturer') or '').upper().split())
    return '%s %s %s' % (supplier_id, mpn, mfr)


class LookupService:
    def __init__(self, clients=None, cache=None, concurrency=3, include_alternates=True):
        self.clients = [c for c in (clients or []) if c and c.configured]
        self.cache = cache
        self.concurrency = max(1, concurrency)
        self.include_alternates = bool(include_alternates)

    @property
    def suppliers(self):
        return [{'id': c.id, 'name': c.name} for c in self.clients]

    def lookup_parts(self, parts, on_progress=None):
        parts = list(parts or [])
        if not self.clients:
            raise RuntimeError('No supplier is configured. Set DigiKey and/or Mouser credentials in .env.')

        # The same part number often appears on several BOM lines; one API call
        # per distinct part per supplier is enough for all of them.
        # Alternates are looked up alongside their primary rather than instead
        # of it: knowing a second source is available before the primary goes
        # obsolete is the reason a BOM carries the column at all.
        targets = []
        for part in parts:
            if not part or not part.get('mpn'):
                continue
            targets.append(part)
            if self.include_alternates:
                targets.extend(alternate_parts(part))

        jobs = {}
        for part in targets:
            for client in self.clients:
                key = cache_key(client.id, part)
                if key not in jobs:
                    jobs[key] = (key, client, part)

        job_list = list(jobs.values())
        resolved = {}
        stats = {'apiCalls': 0, 'cacheHits': 0, 'errors': 0, 'lookups': len(job_list), 'completed': 0}
        lock = threading.Lock()

        def note_progress(client, part):
            with lock:
                stats['completed'] += 1
                progress = {
                    'completed': stats['completed'],
                    'total': len(job_list),
                    'supplier': client.name,
                    'mpn': part.get('mpn'),
                    'apiCalls': stats['apiCalls'],
                    'cacheHits': stats['cacheHits'],
                    'errors': stats['errors'],
                }
            if on_progress:
                on_progress(progress)

        def run_batch(batch):
            """One request covering several parts, for aggregators that accept
            a list. Cache hits are served first so only genuine misses go out."""
            client = batch[0][1]
            pending = []
            for key, _, part in batch:
                cached = self.cache.get(key) if self.cache is not None else None
                if cached is not None:
                    with lock:
                        stats['cacheHits'] += 1
                        resolved[key] = cached
                    note_progress(client, part)
                else:
                    pending.append((key, part))

            if not pending:
                return

            parts = [part for _, part in pending]
            # A batch is usually one request, but a client may split it — Nexar
            # falls back to per-part queries on a plan without the batched one.
            # A client that counts its own requests is believed, because the
            # number people read this for is quota spent, not batches sent.
            before = getattr(client, 'requests_made', None)

            def requests_spent():
                after = getattr(client, 'requests_made', None)
                if isinstance(before, int) and isinstance(after, int):
                    return max(1, after - before)
                return 1

            try:
                records = client.fetch_records(parts) or {}
                with lock:
                    stats['apiCalls'] += requests_spent()
                for key, part in pending:
                    entry = _entry_for(records.get(part.get('mpn')))
                    if self.cache is not None:
                        self.cache.set(key, entry)
                    with lock:
                        resolved[key] = entry
            except Exception as err:
                message = str(err) or repr(err)
                with lock:
                    stats['apiCalls'] += requests_spent()
                    stats['errors'] += 1
                    # One failed request fails every part it carried, but the
                    # failure is not cached, so the next run retries them.
                    for key, _ in pending:
                        resolved[key] = {'error': message}
            finally:
                for _, part in pending:
                    note_progress(client, part)

        def run(job):
            key, client, part = job
            try:
                if self.cache is not None:
                    cached = self.cache.get(key)
                    if cached is not None:
                        with lock:
                            stats['cacheHits'] += 1
                            resolved[key] = cached
                        return
                try:
                    record = client.fetch_record(part)
                    entry = _entry_for(record)
                    if self.cache is not None:
                        self.cache.set(key, entry)
                    with lock:
                        stats['apiCalls'] += 1
                        resolved[key] = entry
                except Exception as err:
                    # Errors are deliberately not cached: a rate limit or a
                    # network blip should not poison the next run.
                    with lock:
                        stats['apiCalls'] += 1
                        stats['errors'] += 1
                        resolved[key] = {'error': str(err) or repr(err)}
            finally:
                note_progress(client, part)

        # Split the work: clients that accept a list of parts get chunked
        # requests, the rest keep one request per part.
        batches = []
        singles = []
        by_client = {}
        for job in job_list:
            by_client.setdefault(job[1].id, []).append(job)

        for client_jobs in by_client.values():
            client = client_jobs[0][1]
            size = getattr(client, 'batch_size', 1) or 1
            if size > 1 and hasattr(client, 'fetch_records'):
                for start in range(0, len(client_jobs), size):
                    batches.append(client_jobs[start:start + size])
            else:
                singles.extend(client_jobs)

        units = [('batch', b) for b in batches] + [('single', s) for s in singles]
        if units:
            workers = min(self.concurrency, len(units))
            with ThreadPoolExecutor(max_workers=workers) as pool:
                list(pool.map(
                    lambda unit: run_batch(unit[1]) if unit[0] == 'batch' else run(unit[1]),
                    units,
                ))

        def offers_for(target):
            offers = {}
            for client in self.clients:
                entry = resolved.get(cache_key(client.id, target))
                if entry is None:
                    offers[client.id] = missing_offer(client.name, 'Not looked up')
                elif entry.get('error'):
                    offers[client.id] = error_offer(client.name, entry['error'])
                elif entry.get('notFound'):
                    offers[client.id] = missing_offer(
                        client.name, _miss_reason(entry, client.name, target.get('mpn'))
                    )
                else:
                    offers[client.id] = record_to_offer(entry['record'], target)
            return offers

        rows = []
        for index, part in enumerate(parts):
            offers = offers_for(part)

            alternates = []
            if self.include_alternates:
                for alternate in alternate_parts(part):
                    alternate_offers = offers_for(alternate)
                    alternates.append(summarize_alternate(
                        alternate, alternate_offers,
                        [alternate_offers[c.id] for c in self.clients],
                    ))

            offer_list = [offers[c.id] for c in self.clients]
            rows.append({
                'index': index,
                'row': part.get('row') or index + 1,
                'mpn': part.get('mpn'),
                'quantity': part.get('quantity') or 1,
                'reference': part.get('reference'),
                'manufacturer': part.get('manufacturer'),
                'description': part.get('description'),
                'offers': offers,
                'comparison': compare_offers(offer_list, part.get('quantity')),
                # The BOM's own approved alternates, each priced and compared
                # the same way, kept beside the primary rather than mixed into
                # it: they are a fallback, not a competing quote.
                'alternates': alternates,
            })

        return {'rows': rows, 'suppliers': self.suppliers, 'stats': stats}


def summarize_bom(rows, suppliers):
    """Roll-up across the whole BOM: what it costs to single-source from each
    supplier, what the cheapest per-line mix costs, and where the risks are."""
    supplier_totals = {
        s['id']: {
            'id': s['id'],
            'name': s['name'],
            'total': 0,
            'linesPriced': 0,
            'linesMissing': 0,
            'linesShort': 0,
            'complete': True,
        }
        for s in suppliers
    }

    summary = {
        'lines': len(rows),
        'totalQuantity': 0,
        'supplierTotals': supplier_totals,
        'bestMixTotal': 0,
        'bestMixLines': 0,
        'unpricedLines': 0,
        'notFoundLines': 0,
        'errorLines': 0,
        'lifecycleCounts': {},
        'riskLines': [],
        'currency': None,
        'cheapestSingleSource': None,
        'mixSavings': None,
    }

    for row in rows:
        summary['totalQuantity'] += row.get('quantity') or 0
        comparison = row['comparison']

        status = comparison.get('lifecycle') or Lifecycle.UNKNOWN
        summary['lifecycleCounts'][status] = summary['lifecycleCounts'].get(status, 0) + 1

        any_found = False
        any_error = False
        for supplier in suppliers:
            offer = row['offers'].get(supplier['id'])
            totals = supplier_totals[supplier['id']]
            if not offer or not offer.get('found'):
                if offer and offer.get('error'):
                    any_error = True
                totals['linesMissing'] += 1
                totals['complete'] = False
                continue
            any_found = True
            if offer.get('currency') and not summary['currency']:
                summary['currency'] = offer['currency']
            if offer.get('extendedPrice') is not None:
                totals['total'] = round_to(totals['total'] + offer['extendedPrice'], 4)
                totals['linesPriced'] += 1
            else:
                totals['complete'] = False
            if offer.get('stockSufficient') is False:
                totals['linesShort'] += 1

        if not any_found:
            summary['notFoundLines'] += 1
            if any_error:
                summary['errorLines'] += 1

        if comparison.get('bestPrice') is not None:
            summary['bestMixTotal'] = round_to(summary['bestMixTotal'] + comparison['bestPrice'], 4)
            summary['bestMixLines'] += 1
        elif any_found:
            summary['unpricedLines'] += 1

        severity = comparison.get('lifecycleSeverity')
        no_stock = not comparison.get('stockCovers')
        if not any_found or severity in ('bad', 'warn') or no_stock:
            summary['riskLines'].append({
                'row': row.get('row'),
                'mpn': row.get('mpn'),
                'quantity': row.get('quantity'),
                'lifecycle': comparison.get('lifecycle'),
                'flags': comparison.get('flags'),
            })

    single_source = [t for t in supplier_totals.values() if t['linesPriced'] > 0]
    if single_source:
        cheapest = min(single_source, key=lambda t: t['total'])
        summary['cheapestSingleSource'] = cheapest['id']
        # Only meaningful when both carts cover the same lines, which
        # `complete` and `linesPriced` let the UI check.
        summary['mixSavings'] = round_to(cheapest['total'] - summary['bestMixTotal'], 4)

    summary['currency'] = summary['currency'] or 'USD'
    return summary
