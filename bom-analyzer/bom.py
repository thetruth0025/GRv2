#!/usr/bin/env python3
"""Command-line BOM supplier analysis.

Point it at a bill of materials and it queries DigiKey and Mouser for lead
time, cost, stock and lifecycle status, prints a summary, and writes a
spreadsheet with every supplier column filled in.

    python3 bom.py my-bom.csv -o comparison.xlsx

Standard library only. Credentials come from .env, exactly as the web app.
"""

import argparse
import os
import sys

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

from bomlib.env import load_env  # noqa: E402

load_env()

from bomlib.cache import PartCache  # noqa: E402
from bomlib.digikey import DigiKeyClient  # noqa: E402
from bomlib.lookup import LookupService, summarize_bom  # noqa: E402
from bomlib.mouser import MouserClient  # noqa: E402
from bomlib.trustedparts import TrustedPartsClient  # noqa: E402
from bomlib.report import (  # noqa: E402
    WRITERS,
    Palette,
    format_days,
    integer,
    lead_label,
    money,
    render_table,
    truncate,
)
from bomlib import dmsms as dmsms_module
from bomlib.prepare import (
    DEFAULT_IGNORE_PREFIXES,
    describe_exclusions,
    parse_prefixes,
    prepare_lines,
)
from bomlib.spreadsheet import (  # noqa: E402
    FIELD_ORDER,
    extract_bom,
    line_from_row,
    parse_workbook,
)

EXIT_OK = 0
EXIT_ERROR = 1
EXIT_FAILED_CHECK = 2

FORMAT_BY_EXTENSION = {'.xlsx': 'xlsx', '.csv': 'csv', '.json': 'json'}


def build_parser():
    parser = argparse.ArgumentParser(
        prog='bom.py',
        description='Compare DigiKey and Mouser on price, lead time, stock and lifecycle '
                    'status for every part in a bill of materials.',
        epilog='Credentials are read from .env (see .env.example). Either supplier alone works.',
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        'bom', nargs='?',
        help='BOM file (.csv, .tsv, .xlsx), or - to read part numbers from stdin, '
             'one per line with an optional quantity after a comma. Omit it when '
             'using --part.',
    )
    parser.add_argument(
        '-p', '--part', action='append', metavar='MPN', dest='parts',
        help='Look up a part number directly, without a BOM file. Repeatable. '
             'Add a quantity after a comma: --part "STM32F103C8T6,25".',
    )
    parser.add_argument('-o', '--output', help='Write the full comparison here (.xlsx, .csv or .json).')
    parser.add_argument('-f', '--format', choices=sorted(WRITERS), help='Override the format inferred from --output.')
    parser.add_argument(
        '-b', '--build-quantity', type=int, default=1, metavar='N',
        help='Number of units being built; multiplies every BOM quantity by N (default: 1).',
    )
    parser.add_argument(
        '-s', '--supplier', action='append', choices=['digikey', 'mouser', 'trustedparts'],
        help='Query only this supplier. Repeatable.',
    )
    parser.add_argument(
        '--limit', type=int, metavar='N', help='Analyze only the first N parts.',
    )

    screening = parser.add_argument_group(
        'screening', 'Drop lines that are not worth a supplier lookup. On by default.')
    screening.add_argument(
        '--ignore-prefix', action='append', metavar='PREFIX', dest='ignore_prefixes',
        help='Skip part numbers starting with PREFIX (case-insensitive). Repeatable. '
             'Replaces the default list: %s.' % ', '.join(DEFAULT_IGNORE_PREFIXES),
    )
    screening.add_argument(
        '--no-ignore-prefixes', action='store_true',
        help='Look up in-house part numbers too.',
    )
    screening.add_argument(
        '--no-merge-duplicates', action='store_true',
        help='Keep repeated part numbers as separate lines instead of adding their quantities.',
    )
    screening.add_argument(
        '--show-skipped', action='store_true',
        help='List every skipped line and why it was skipped.',
    )

    columns = parser.add_argument_group('column mapping', 'Override the automatic header detection. '
                                                          'Each takes a header name or a 0-based index.')
    for field in FIELD_ORDER:
        columns.add_argument('--%s-column' % field, metavar='COL', dest='%s_column' % field)
    columns.add_argument(
        '--list-columns', action='store_true',
        help='Print the detected headers and column mapping, then exit.',
    )

    obsolescence = parser.add_argument_group(
        'obsolescence', 'DMSMS case form for the parts whose supply is ending.')
    obsolescence.add_argument(
        '--dmsms', metavar='FILE',
        help='Write a DMSMS case form (.xlsx) covering every at-risk part: obsolete, '
             'discontinued, end of life, last time buy and NRND.',
    )
    obsolescence.add_argument(
        '--program', metavar='NAME',
        help='Program or platform the DMSMS form is for. Required with --dmsms.',
    )
    obsolescence.add_argument(
        '--dmsms-status', action='append', metavar='STATUS', dest='dmsms_statuses',
        help='Limit the form to these lifecycle statuses, e.g. --dmsms-status Obsolete. '
             'Repeatable. Defaults to every at-risk status.',
    )

    output = parser.add_argument_group('output')
    output.add_argument('--table', dest='table', action='store_true', default=None,
                        help='Always print the per-part table.')
    output.add_argument('--no-table', dest='table', action='store_false',
                        help='Never print the per-part table.')
    output.add_argument('-q', '--quiet', action='store_true', help='Only print errors.')
    output.add_argument('--no-color', action='store_true', help='Disable coloured output.')
    output.add_argument(
        '--fail-on', choices=['never', 'notfound', 'risk'], default='never',
        help='Exit non-zero when parts are missing (notfound) or any line carries a risk flag '
             '(risk). Useful in scripts and CI (default: never).',
    )

    cache = parser.add_argument_group('cache')
    cache.add_argument('--no-cache', action='store_true', help='Ignore cached supplier answers.')
    cache.add_argument('--clear-cache', action='store_true', help='Empty the cache before running.')
    return parser


def main(argv=None):
    args = build_parser().parse_args(argv)
    paint = Palette(_use_color(args))
    log = _make_logger(args.quiet)

    try:
        lines, source = read_bom(args)
    except (OSError, ValueError) as err:
        print('error: %s' % err, file=sys.stderr)
        return EXIT_ERROR

    if args.list_columns:
        return EXIT_OK

    if not lines:
        print('error: no part numbers found in %s. Try --list-columns to see the detected mapping.'
              % source, file=sys.stderr)
        return EXIT_ERROR

    # Validated before the falsy-zero guard below, so -b 0 is refused rather
    # than silently skipped.
    if args.build_quantity < 1:
        print('error: --build-quantity must be 1 or more', file=sys.stderr)
        return EXIT_ERROR
    if args.limit is not None and args.limit < 1:
        print('error: --limit must be 1 or more', file=sys.stderr)
        return EXIT_ERROR
    if args.dmsms and not args.program:
        print('error: --dmsms needs --program to name the form', file=sys.stderr)
        return EXIT_ERROR
    if args.program and not args.dmsms:
        print('error: --program only applies with --dmsms', file=sys.stderr)
        return EXIT_ERROR

    # Screened before --limit so the limit counts parts that will really be
    # looked up, not assembly and cable lines that never reach a supplier.
    # Naming a part on the command line is a direct question about that part,
    # so the in-house prefixes do not answer it with nothing — unless the prefix
    # list was given explicitly, which is somebody asking for it.
    manual = bool(args.parts) and not args.ignore_prefixes
    screened = prepare_lines(
        lines,
        ignore_prefixes=[] if (args.no_ignore_prefixes or manual)
                        else parse_prefixes(args.ignore_prefixes),
        merge_duplicates=not args.no_merge_duplicates,
    )
    lines = screened['lines']
    excluded = screened['excluded']
    note = describe_exclusions(excluded)
    if note:
        log(paint(note, 'dim'))
    if args.show_skipped and excluded:
        for entry in excluded:
            log('  %-6s %-28s %s' % (
                entry.get('row') or '', truncate(entry.get('mpn'), 28), entry.get('detail') or ''))
    if not lines:
        print('error: every line in %s was skipped. Use --no-ignore-prefixes to look them up anyway.'
              % source, file=sys.stderr)
        return EXIT_ERROR

    if args.limit:
        lines = lines[:args.limit]
    if args.build_quantity != 1:
        for line in lines:
            line['quantity'] = line['quantity'] * args.build_quantity

    service, cache = build_service(args)
    if not service.clients:
        print('error: no supplier credentials found. Copy .env.example to .env and add your '
              'API keys, or set DIGIKEY_CLIENT_ID / DIGIKEY_CLIENT_SECRET / MOUSER_API_KEY / '
              'TRUSTEDPARTS_API_KEY.',
              file=sys.stderr)
        return EXIT_ERROR

    if args.clear_cache and cache is not None:
        cache.clear()

    names = _join_names([c.name for c in service.clients])
    log('Analyzing %s from %s against %s…'
        % (_plural(len(lines), 'part'), source, names))

    progress = _Progress(len(lines) * len(service.clients), enabled=not args.quiet)
    try:
        result = service.lookup_parts(lines, on_progress=progress.update)
    except Exception as err:
        progress.finish()
        print('error: %s' % err, file=sys.stderr)
        return EXIT_ERROR
    finally:
        if cache is not None:
            cache.flush()
    progress.finish()

    result['excluded'] = excluded
    summary = summarize_bom(result['rows'], result['suppliers'])

    if not args.quiet:
        show_table = args.table if args.table is not None else not args.output
        if show_table:
            print(render_parts_table(result, summary, paint))
            print()
        print(render_summary(result, summary, paint))

    if args.output:
        fmt = args.format or FORMAT_BY_EXTENSION.get(os.path.splitext(args.output)[1].lower())
        if not fmt:
            print('error: cannot tell the output format from %r — pass --format.' % args.output,
                  file=sys.stderr)
            return EXIT_ERROR
        try:
            WRITERS[fmt](args.output, result, summary)
        except OSError as err:
            print('error: could not write %s: %s' % (args.output, err), file=sys.stderr)
            return EXIT_ERROR
        log('\nWrote %s' % paint(args.output, 'bold'))

    if args.dmsms:
        code = write_dmsms_form(args, result, source, log, paint)
        if code != EXIT_OK:
            return code

    return check_exit(args.fail_on, summary)


def write_dmsms_form(args, result, source, log, paint):
    """Every at-risk part goes on the form; the browser is where you pick."""
    wanted = args.dmsms_statuses
    rows = dmsms_module.candidate_rows(result)
    if wanted:
        allowed = [w.strip().lower() for w in wanted]
        rows = [r for r in rows
                if str(r['comparison'].get('lifecycle') or '').lower() in allowed]

    if not rows:
        print('error: no at-risk parts to put on a DMSMS form.', file=sys.stderr)
        return EXIT_ERROR

    for row in rows:
        row['assembly'] = source

    try:
        dmsms_module.write_form(args.dmsms, rows, {
            'program': args.program,
            'scope': source,
        })
    except OSError as err:
        print('error: could not write %s: %s' % (args.dmsms, err), file=sys.stderr)
        return EXIT_ERROR

    log('Wrote %s — %s on the DMSMS form'
        % (paint(args.dmsms, 'bold'), _plural(len(rows), 'part')))
    return EXIT_OK


# ── Input ───────────────────────────────────────────────────────────────────

def read_bom(args):
    """Return (lines, source-label). Honours any --*-column overrides."""
    if args.parts:
        lines = parse_pasted('\n'.join(args.parts))
        if not lines:
            raise ValueError('no usable part numbers in --part')
        if args.list_columns:
            print('--part takes part numbers directly; column mapping does not apply.')
        return lines, _plural(len(lines), 'part number')

    if args.bom is None:
        raise ValueError('give a BOM file, - to read stdin, or --part to look up a part number')

    if args.bom == '-':
        lines = parse_pasted(sys.stdin.read())
        if args.list_columns:
            print('stdin is read as one part number per line; column mapping does not apply.')
        return lines, 'stdin'

    with open(args.bom, 'rb') as handle:
        data = handle.read()
    grid = parse_workbook(data, args.bom)
    parsed = extract_bom(grid)
    mapping = apply_overrides(parsed['mapping'], parsed['headers'], args)

    if args.list_columns:
        print_columns(parsed, mapping)

    if mapping != parsed['mapping']:
        # Re-derive from the raw grid so overrides take effect.
        start = parsed['headerRow'] + 1
        lines = []
        for index, row in enumerate(grid[start:]):
            if not row or all(not str(cell or '').strip() for cell in row):
                continue
            line = line_from_row(row, mapping, start + index)
            if line['mpn']:
                lines.append(line)
    else:
        lines = parsed['lines']

    return lines, os.path.basename(args.bom)


def apply_overrides(mapping, headers, args):
    resolved = dict(mapping)
    squashed = [str(h or '').strip().lower() for h in headers]

    for field in FIELD_ORDER:
        value = getattr(args, '%s_column' % field, None)
        if value is None:
            continue
        if value == '':
            resolved.pop(field, None)
            continue
        try:
            index = int(value)
        except ValueError:
            try:
                index = squashed.index(str(value).strip().lower())
            except ValueError:
                raise ValueError(
                    'no column named %r. Available: %s'
                    % (value, ', '.join(h for h in headers if h) or '(none detected)')
                )
        if index < 0 or index >= max(len(headers), 1):
            raise ValueError('column index %d is out of range (0-%d)' % (index, len(headers) - 1))
        resolved[field] = index
    return resolved


def print_columns(parsed, mapping):
    print('Header detected on row %d.' % (parsed['headerRow'] + 1))
    print()
    print('  %-4s %-32s %s' % ('IDX', 'HEADER', 'MAPPED TO'))
    reverse = {index: field for field, index in mapping.items()}
    for index, header in enumerate(parsed['headers']):
        print('  %-4d %-32s %s' % (index, truncate(header or '(blank)', 32), reverse.get(index, '')))
    print()
    unmapped = [f for f in FIELD_ORDER if f not in mapping]
    if unmapped:
        print('Not mapped: %s' % ', '.join(unmapped))
    print('Override with e.g. --mpn-column "Mfr. Part #" or --mpn-column 3')


def parse_pasted(text):
    """One part number per line, optional quantity after a comma or tab."""
    lines = []
    for index, raw in enumerate(text.splitlines()):
        stripped = raw.strip()
        if not stripped or stripped.startswith('#'):
            continue
        parts = [p.strip() for p in stripped.replace('\t', ',').split(',')]
        mpn = parts[0]
        if not mpn:
            continue
        quantity = 1
        if len(parts) > 1:
            digits = ''.join(ch for ch in parts[1] if ch.isdigit())
            if digits:
                quantity = max(1, int(digits))
        lines.append({
            'row': index + 1, 'mpn': mpn, 'quantity': quantity,
            'reference': None,
            'manufacturer': parts[2] if len(parts) > 2 and parts[2] else None,
            'description': None,
        })
    return lines


def build_service(args):
    digikey = DigiKeyClient(
        client_id=os.environ.get('DIGIKEY_CLIENT_ID'),
        client_secret=os.environ.get('DIGIKEY_CLIENT_SECRET'),
        sandbox=str(os.environ.get('DIGIKEY_SANDBOX', '')).strip().lower() in ('1', 'true', 'yes'),
        site=os.environ.get('DIGIKEY_SITE'),
        language=os.environ.get('DIGIKEY_LANGUAGE'),
        currency=os.environ.get('DIGIKEY_CURRENCY'),
    )
    mouser = MouserClient(
        api_key=os.environ.get('MOUSER_API_KEY'),
        currency=os.environ.get('MOUSER_CURRENCY'),
    )

    trustedparts = TrustedPartsClient(
        api_key=os.environ.get('TRUSTEDPARTS_API_KEY'),
        currency=os.environ.get('TRUSTEDPARTS_CURRENCY') or os.environ.get('DIGIKEY_CURRENCY'),
        country=os.environ.get('TRUSTEDPARTS_COUNTRY'),
        language=os.environ.get('TRUSTEDPARTS_LANGUAGE'),
        user_agent=os.environ.get('TRUSTEDPARTS_USER_AGENT'),
        distributors=[d.strip() for d in str(os.environ.get('TRUSTEDPARTS_DISTRIBUTORS') or '').split(',') if d.strip()],
        in_stock_only=str(os.environ.get('TRUSTEDPARTS_IN_STOCK_ONLY', '')).strip().lower() in ('1', 'true', 'yes'),
        use_cached_data=str(os.environ.get('TRUSTEDPARTS_USE_CACHED_DATA', '')).strip().lower() in ('1', 'true', 'yes'),
    )

    clients = [digikey, mouser, trustedparts]
    if args.supplier:
        wanted = set(args.supplier)
        clients = [c for c in clients if c.id in wanted]

    cache = None
    if not args.no_cache:
        cache_file = os.environ.get('CACHE_FILE')
        path = None if cache_file == 'none' else (cache_file or os.path.join(BASE_DIR, '.cache', 'parts.json'))
        try:
            ttl = float(os.environ.get('CACHE_TTL_HOURS') or 6)
        except ValueError:
            ttl = 6.0
        cache = PartCache(ttl_seconds=ttl * 3600, path=path)

    try:
        concurrency = int(os.environ.get('LOOKUP_CONCURRENCY') or 3)
    except ValueError:
        concurrency = 3

    return LookupService(clients=clients, cache=cache, concurrency=concurrency), cache


# ── Output ──────────────────────────────────────────────────────────────────

class _Progress:
    """A single rewritten line on stderr, so stdout stays pipeable."""

    def __init__(self, total, enabled=True):
        self.total = max(1, total)
        self.enabled = enabled and sys.stderr.isatty()
        self.width = 0

    def update(self, progress):
        if not self.enabled:
            return
        done, total = progress['completed'], progress['total']
        bars = int(24 * done / max(1, total))
        text = '  [%s%s] %d/%d  %s' % (
            '#' * bars, '.' * (24 - bars), done, total,
            truncate(progress.get('mpn') or '', 28),
        )
        self.width = max(self.width, len(text))
        sys.stderr.write('\r' + text.ljust(self.width))
        sys.stderr.flush()

    def finish(self):
        if self.enabled and self.width:
            sys.stderr.write('\r' + ' ' * self.width + '\r')
            sys.stderr.flush()


def render_parts_table(result, summary, paint):
    suppliers = result['suppliers']
    currency = summary.get('currency') or 'USD'

    headers = ['ROW', 'PART', 'QTY']
    aligns = ['right', 'left', 'right']
    for supplier in suppliers:
        headers += ['%s STOCK' % supplier['name'].upper(), 'LEAD', 'EXT']
        aligns += ['right', 'left', 'right']
    headers += ['BEST', 'LIFECYCLE']
    aligns += ['left', 'left']

    rows = []
    for row in result['rows']:
        comparison = row['comparison']
        record = [str(row.get('row')), truncate(row.get('mpn'), 28), integer(row.get('quantity'))]
        for supplier in suppliers:
            offer = row['offers'].get(supplier['id'])
            if not offer or not offer.get('found'):
                dash = paint('—', 'dim')
                record += [dash, dash, dash]
                continue
            stock = integer(offer.get('stock'))
            if offer.get('stockSufficient') is False:
                stock = paint(stock, 'yellow')
            extended = money(offer.get('extendedPrice'), currency)
            if comparison.get('bestPriceSupplier') == supplier['name']:
                extended = paint(extended, 'green')
            record += [stock, lead_label(offer), extended]

        record.append(comparison.get('recommendedSupplier') or paint('—', 'dim'))
        severity = comparison.get('lifecycleSeverity')
        label = comparison.get('lifecycle') or 'Unknown'
        label = {'Not Recommended for New Designs': 'NRND'}.get(label, label)
        colour = {'bad': 'red', 'warn': 'yellow', 'unknown': 'dim'}.get(severity)
        record.append(paint(label, colour) if colour else label)
        rows.append(record)

    return render_table(headers, rows, aligns, paint)


def render_summary(result, summary, paint):
    currency = summary.get('currency') or 'USD'
    out = [paint('Summary', 'bold')]
    out.append('  %-22s %s across %s'
               % ('Lines analyzed', summary['lines'], _plural(summary['totalQuantity'], 'unit')))

    for supplier in result['suppliers']:
        totals = summary['supplierTotals'][supplier['id']]
        note = []
        if totals['linesMissing']:
            note.append('%d not carried' % totals['linesMissing'])
        if totals['linesShort']:
            note.append('%d short on stock' % totals['linesShort'])
        suffix = paint('  (%s)' % ', '.join(note), 'yellow') if note else ''
        out.append('  %-22s %s%s' % ('%s cart' % supplier['name'],
                                     money(totals['total'], currency), suffix))

    if len(result['suppliers']) > 1:
        savings = summary.get('mixSavings')
        extra = ''
        if savings is not None and savings > 0:
            extra = paint('  (saves %s vs. single-sourcing)' % money(savings, currency), 'green')
        out.append('  %-22s %s%s' % ('Cheapest per line', money(summary['bestMixTotal'], currency), extra))

    stock_risk = sum(1 for r in result['rows'] if not r['comparison'].get('inStockSuppliers'))
    lifecycle_risk = sum(1 for r in result['rows']
                         if r['comparison'].get('lifecycleSeverity') in ('bad', 'warn'))
    out.append('  %-22s %s' % ('Stock risk',
                               paint(stock_risk, 'red') if stock_risk else paint('0', 'green')))
    out.append('  %-22s %s' % ('Lifecycle risk',
                               paint(lifecycle_risk, 'yellow') if lifecycle_risk else paint('0', 'green')))
    if summary['notFoundLines']:
        out.append('  %-22s %s' % ('Not found', paint(summary['notFoundLines'], 'red')))

    stats = result['stats']
    out.append(paint('  %-22s %d live, %d cached%s'
                     % ('Queries', stats['apiCalls'], stats['cacheHits'],
                        ', %d failed' % stats['errors'] if stats['errors'] else ''), 'dim'))

    risks = summary.get('riskLines') or []
    if risks:
        out.append('')
        out.append(paint('Needs attention', 'bold'))
        for risk in risks[:20]:
            flags = '; '.join(f['text'] for f in risk.get('flags') or []) or 'see details'
            worst = max((f.get('level') for f in risk.get('flags') or []), default='info')
            colour = {'bad': 'red', 'warn': 'yellow'}.get(worst)
            marker = paint('•', colour) if colour else '•'
            out.append('  %s %-28s %s' % (marker, truncate(risk['mpn'], 28), truncate(flags, 78)))
        if len(risks) > 20:
            out.append(paint('  … and %d more' % (len(risks) - 20), 'dim'))

    return '\n'.join(out)


def check_exit(fail_on, summary):
    if fail_on == 'notfound' and summary['notFoundLines']:
        return EXIT_FAILED_CHECK
    if fail_on == 'risk' and summary.get('riskLines'):
        return EXIT_FAILED_CHECK
    return EXIT_OK


def _use_color(args):
    if args.no_color or os.environ.get('NO_COLOR'):
        return False
    return sys.stdout.isatty()


def _make_logger(quiet):
    def log(message):
        if not quiet:
            print(message)
    return log


def _join_names(names):
    """"A", "A and B", "A, B and C"."""
    if len(names) <= 1:
        return names[0] if names else ''
    return '%s and %s' % (', '.join(names[:-1]), names[-1])


def _plural(count, noun):
    return '%s %s%s' % (format(int(count), ','), noun, '' if count == 1 else 's')


if __name__ == '__main__':
    sys.exit(main())
