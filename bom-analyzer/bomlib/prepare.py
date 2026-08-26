"""Decide which BOM lines are worth a supplier lookup.

Four things are screened out before any API call goes out, because each one
costs a request per supplier and adds a row nobody reads:

  * lines the BOM itself marks as not for sourcing, in a skip-to-production
    column — the author of the BOM said so, which beats any rule here;
  * in-house part numbers — assemblies, cables, drawings and bare boards are
    not distributor stock, so no distributor will ever match them;
  * the same part listed twice inside one BOM, which is one purchase, not two;
  * a part already claimed by an earlier BOM in the same session.

Nothing is thrown away: every screened line comes back in `excluded` with the
reason, so the UI and the exports can still account for it.
"""

# In-house prefixes: assemblies, cables, drawings/designs and bare boards.
DEFAULT_IGNORE_PREFIXES = ('ASY0', 'CBL0', 'DES0', 'PCB0')

IGNORED = 'ignored'
MERGED = 'merged'
DUPLICATE = 'duplicate'
FLAGGED = 'flagged'

# What a spreadsheet means by yes in a tick-box column. Anything else — blank,
# NO, N/A, a date, a note — leaves the line in: dropping a part because of a
# value nobody recognised is the wrong way to be wrong.
SKIP_VALUES = ('YES', 'Y', 'TRUE', 'T', '1', 'X', '✓', '✔')


def fold(value):
    """Upper case, with runs of whitespace collapsed to one."""
    return ' '.join(str(value or '').upper().split())


def normalize_mpn(value):
    """The form two part numbers are compared in: case and spacing folded."""
    return fold(value)


def skip_requested(value):
    """True when a BOM's own skip column says to leave this line alone."""
    return fold(value) in SKIP_VALUES


def parse_prefixes(raw, default=DEFAULT_IGNORE_PREFIXES):
    """Read a comma- or whitespace-separated prefix list.

    `None` means "not configured" and falls back to the default; an empty
    string is an explicit "ignore nothing", which is why the two are not
    collapsed into one falsy check.
    """
    if raw is None:
        return list(default)
    if isinstance(raw, (list, tuple)):
        items = raw
    else:
        items = str(raw).replace(',', ' ').split()
    return [normalize_mpn(item) for item in items if str(item).strip()]


def matching_prefix(mpn, prefixes):
    """The prefix that rules this part out, or None."""
    candidate = normalize_mpn(mpn)
    if not candidate:
        return None
    for prefix in prefixes or []:
        if prefix and candidate.startswith(prefix):
            return prefix
    return None


def _excluded(line, reason, detail):
    return {
        'row': line.get('row'),
        'mpn': line.get('mpn'),
        'quantity': line.get('quantity'),
        'reference': line.get('reference'),
        'manufacturer': line.get('manufacturer'),
        'description': line.get('description'),
        'reason': reason,
        'detail': detail,
    }


def _join_references(*values):
    """Combine reference designators without repeating any."""
    seen = []
    for value in values:
        for piece in str(value or '').replace(';', ',').split(','):
            piece = piece.strip()
            if piece and piece not in seen:
                seen.append(piece)
    return ', '.join(seen) or None


def prepare_lines(lines, ignore_prefixes=DEFAULT_IGNORE_PREFIXES,
                  merge_duplicates=True, claimed=None, honour_skip_flag=True):
    """Screen a BOM down to the lines that should actually be looked up.

    `claimed` maps a normalized part number to the name of whatever already
    owns it — another BOM, usually. Those parts are reported as duplicates and
    never looked up again.

    Returns {'lines', 'excluded', 'claimed'}: the survivors, everything screened
    out with its reason, and the part numbers this BOM now owns.
    """
    prefixes = list(ignore_prefixes or [])
    already = dict(claimed or {})

    kept = []
    excluded = []
    owned = {}
    by_mpn = {}

    for line in lines or []:
        if not line or not line.get('mpn'):
            continue

        # Checked first: whoever wrote the BOM marked this line deliberately,
        # and that reason is more use than "in-house part number" would be.
        if honour_skip_flag and skip_requested(line.get('skip')):
            excluded.append(_excluded(
                line, FLAGGED,
                'Marked "%s" in the skip-to-production column' % line.get('skip'),
            ))
            continue

        prefix = matching_prefix(line.get('mpn'), prefixes)
        if prefix:
            excluded.append(_excluded(
                line, IGNORED,
                'In-house part number (starts with %s)' % prefix,
            ))
            continue

        key = normalize_mpn(line.get('mpn'))

        owner = already.get(key)
        if owner:
            excluded.append(_excluded(
                line, DUPLICATE,
                'Already listed in %s' % owner,
            ))
            continue

        first = by_mpn.get(key) if merge_duplicates else None
        if first is not None:
            # One part bought once: the quantities add up and the reference
            # designators join, so the surviving line still buys enough.
            before = first.get('quantity') or 0
            first['quantity'] = before + (line.get('quantity') or 0)
            first['reference'] = _join_references(first.get('reference'), line.get('reference'))
            first.setdefault('mergedRows', []).append(line.get('row'))
            if not first.get('description') and line.get('description'):
                first['description'] = line['description']
            if not first.get('manufacturer') and line.get('manufacturer'):
                first['manufacturer'] = line['manufacturer']
            excluded.append(_excluded(
                line, MERGED,
                'Same part as row %s — quantity added (%s → %s)'
                % (first.get('row'), before, first['quantity']),
            ))
            continue

        entry = dict(line)
        kept.append(entry)
        by_mpn[key] = entry
        owned[key] = None

    return {'lines': kept, 'excluded': excluded, 'claimed': list(owned.keys())}


def excluded_counts(excluded):
    """Tally by reason, for the one-line summaries the UI and CLI print."""
    counts = {}
    for entry in excluded or []:
        reason = entry.get('reason') or IGNORED
        counts[reason] = counts.get(reason, 0) + 1
    return counts


def describe_exclusions(excluded):
    """A short human sentence, or None when nothing was screened out."""
    counts = excluded_counts(excluded)
    pieces = []
    if counts.get(FLAGGED):
        pieces.append('%d marked skip to production' % counts[FLAGGED])
    if counts.get(IGNORED):
        pieces.append('%d in-house' % counts[IGNORED])
    if counts.get(MERGED):
        pieces.append('%d merged duplicate%s' % (counts[MERGED], '' if counts[MERGED] == 1 else 's'))
    if counts.get(DUPLICATE):
        pieces.append('%d already in another BOM' % counts[DUPLICATE])
    if not pieces:
        return None
    return '%d line%s skipped: %s' % (
        len(excluded), '' if len(excluded) == 1 else 's', ', '.join(pieces),
    )
