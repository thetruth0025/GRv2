"""HTTP with the retry behaviour the supplier APIs need, on urllib."""

import json
import time
import urllib.error
import urllib.parse
import urllib.request


class HttpError(Exception):
    def __init__(self, message, status=0, body=None):
        super().__init__(message)
        self.status = status
        self.body = body


def _host_of(url):
    try:
        return urllib.parse.urlparse(url).netloc or url
    except ValueError:
        return url


def request_json(url, method='GET', headers=None, body=None, timeout=20.0, retries=2, base_delay=0.7):
    """Supplier APIs throttle aggressively on free tiers, so 429 and 5xx get a
    backoff rather than being surfaced as a hard failure on the first try."""
    payload = body.encode('utf-8') if isinstance(body, str) else body
    last_error = None

    for attempt in range(retries + 1):
        request = urllib.request.Request(url, data=payload, method=method)
        for key, value in (headers or {}).items():
            request.add_header(key, value)

        try:
            with urllib.request.urlopen(request, timeout=timeout) as response:
                text = response.read().decode('utf-8', errors='replace')
                return _decode(response.status, text)
        except urllib.error.HTTPError as err:
            text = ''
            try:
                text = err.read().decode('utf-8', errors='replace')
            except Exception:
                pass
            data = _safe_json(text)
            error = HttpError(
                'HTTP %d from %s' % (err.code, _host_of(url)),
                err.code,
                data if data is not None else _truncate(text, 400),
            )
            retryable = err.code == 429 or err.code >= 500
            if not retryable or attempt == retries:
                raise error
            last_error = error
            time.sleep(_retry_delay(err.headers.get('Retry-After'), base_delay, attempt))
        except Exception as err:  # timeouts, DNS, connection resets
            last_error = HttpError('%s: %s' % (_host_of(url), err), 0, None)
            if attempt == retries:
                raise last_error
            time.sleep(base_delay * (2 ** attempt))

    raise last_error or HttpError('Request failed', 0, None)


def _retry_delay(retry_after, base_delay, attempt):
    try:
        seconds = float(retry_after)
        if seconds > 0:
            return min(seconds, 10.0)
    except (TypeError, ValueError):
        pass
    return base_delay * (2 ** attempt)


def _decode(status, text):
    return {'status': status, 'data': _safe_json(text), 'text': text}


def _safe_json(text):
    if not text:
        return None
    try:
        return json.loads(text)
    except ValueError:
        return None


def _truncate(text, limit):
    if not isinstance(text, str):
        return text
    return text[:limit] + '…' if len(text) > limit else text
