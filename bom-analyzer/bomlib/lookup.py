"""Run a BOM against every configured supplier and pair the answers up."""

import threading
from concurrent.futures import ThreadPoolExecutor

from .normalize import (
    Lifecycle,
    compare_offers,
    error_offer,
    missing_offer,
    record_to_offer,
    round_to,
)


def cache_key(supplier_id, part):
    mpn = ' '.join(str(part.get('mpn') or '').upper().split())
    mfr = ' '.join(str(part.get('manufacturer') or '').upper().split())
    return '%s %s %s' % (supplier_id, mpn, mfr)


class LookupService:
    def __init__(self, clients=None, cache=None, concurrency=3):
        self.clients = [c for c in (clients or []) if c and c.configured]
        self.cache = cache
        self.concurrency = max(1, concurrency)

    @property
    def suppliers(self):
        return [{'id': c.id, 'name': c.name} for c in self.clients]

    def lookup_parts(self, parts, on_progress=None):
        parts = list(parts or [])
        if not self.clients:
            raise RuntimeError('No supplier is configured. Set DigiKey and/or Mouser credentials in .env.')

        # The same part number often appears on several BOM lines; one API call
        # per distinct part per supplier is enough for all of them.
        jobs = {}
        for part in parts:
            if not part or not part.get('mpn'):
                continue
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
            try:
                records = client.fetch_records(parts) or {}
                with lock:
                    stats['apiCalls'] += 1
                for key, part in pending:
                    record = records.get(part.get('mpn'))
                    entry = {'record': record} if record else {'notFound': True}
                    if self.cache is not None:
                        self.cache.set(key, entry)
                    with lock:
                        resolved[key] = entry
            except Exception as err:
                message = str(err) or repr(err)
                with lock:
                    stats['apiCalls'] += 1
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
                    entry = {'record': record} if record else {'notFound': True}
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

        rows = []
        for index, part in enumerate(parts):
            offers = {}
            for client in self.clients:
                entry = resolved.get(cache_key(client.id, part))
                if entry is None:
                    offers[client.id] = missing_offer(client.name, 'Not looked up')
                elif entry.get('error'):
                    offers[client.id] = error_offer(client.name, entry['error'])
                elif entry.get('notFound'):
                    offers[client.id] = missing_offer(
                        client.name, 'No %s match for %s' % (client.name, part.get('mpn'))
                    )
                else:
                    offers[client.id] = record_to_offer(entry['record'], part)

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
        no_stock = not comparison.get('inStockSuppliers')
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
