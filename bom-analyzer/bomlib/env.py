"""Minimal .env reader so the project stays dependency-free."""

import os


def load_env(path=None):
    """Read KEY=VALUE lines into os.environ.

    Values already present in the environment win, which lets shell exports and
    hosting-platform config override the file.
    """
    target = path or os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), '.env')
    parsed = {}
    try:
        with open(target, 'r', encoding='utf-8') as handle:
            raw = handle.read()
    except FileNotFoundError:
        return parsed

    for line in raw.splitlines():
        trimmed = line.strip()
        if not trimmed or trimmed.startswith('#'):
            continue
        if '=' not in trimmed:
            continue
        key, _, value = trimmed.partition('=')
        key = key.strip()
        if not key:
            continue
        value = value.strip()
        if len(value) > 1 and value[0] == value[-1] and value[0] in ('"', "'"):
            value = value[1:-1]
        parsed[key] = value
        os.environ.setdefault(key, value)
    return parsed
