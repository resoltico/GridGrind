"""Confirm a single publication write using bounded, read-only observations.

Four observations, separated by one, two and three seconds, cover visibility
lag without repeating a mutation. Each owner distinguishes its exact prior
state from a conflicting identity; this module never interprets remote state.
"""
import os
import re
import subprocess
import time
import urllib.error
import traceback
from pathlib import Path


class ReleaseFailure(Exception):
    def __init__(self, message, *, uncertain=False, category='refusal', status=None,
                 request_id=None, exit_code=None, reason=None):
        super().__init__(message)
        self.uncertain = uncertain
        self.details = {'category': category, 'type': type(self).__name__, 'message': redact(message)}
        if status is not None:
            self.details['httpStatus'] = status
        if request_id and re.fullmatch(r'[A-Za-z0-9:_-]{1,128}', request_id):
            self.details['requestId'] = redact(request_id)
        if exit_code is not None:
            self.details['exitCode'] = exit_code
        if reason:
            self.details['reason'] = redact(reason)


def _redacted(value):
    """Sanitize the full text before either classifying or bounding it."""
    if isinstance(value, bytes):
        value = value.decode('utf-8', errors='replace')
    text = str(value)
    for key, secret in os.environ.items():
        if any(word in key.upper() for word in ('TOKEN', 'PASSWORD', 'SECRET', 'AUTHORIZATION')) and secret:
            text = text.replace(secret, '[REDACTED]')
    text = re.sub(r'(?i)\b(?:Bearer|Basic)\s+\S+', '[REDACTED AUTH]', text)
    text = re.sub(r'(?im)\b(?:authorization|proxy-authorization)\s*[:=][^\r\n]*', '[REDACTED AUTH]', text)
    text = re.sub(r'(?i)https?://[^\s<>"\']+', '[REDACTED URL]', text)
    return ''.join(c if c.isprintable() or c in '\n\t' else '?' for c in text)


def redact(value):
    return _redacted(value)[:1024]


class RegistryReadError(ReleaseFailure, ValueError):
    def __init__(self, status, request_id=None):
        self.status = status
        super().__init__('Registry read failed with HTTP ' + str(status), category='http',
                         status=status, request_id=request_id)


def diagnostic(error):
    # Unexpected programming errors are named, never interpreted as publication outcomes.
    if isinstance(error, ReleaseFailure):
        return dict(error.details)
    if isinstance(error, urllib.error.HTTPError):
        result = {'category': 'http', 'type': type(error).__name__, 'httpStatus': error.code}
        request_id = error.headers.get('X-GitHub-Request-Id') if error.headers else None
        if request_id and re.fullmatch(r'[A-Za-z0-9:_-]{1,128}', request_id):
            result['requestId'] = redact(request_id)
        return result
    result = {'category': 'unexpected', 'type': type(error).__name__}
    if error.__traceback__:
        result['frames'] = [{'file': Path(frame.filename).name, 'line': frame.lineno, 'function': frame.name}
                            for frame in traceback.extract_tb(error.__traceback__)[-8:]]
    return result


def _record(record, evidence, primary=False):
    if record:
        try:
            record(evidence)
        except Exception:
            if not primary:
                raise


def confirm(operation, write, observe, record=None):
    evidence = {'operation': operation, 'request': 'acknowledged', 'observations': []}
    try:
        write()
    except ReleaseFailure as error:
        evidence.update(request='failed', requestError=diagnostic(error))
        if not error.uncertain:
            _record(record, evidence, primary=True)
            raise
    for attempt in range(4):
        if attempt:
            time.sleep(attempt)
        try:
            state = observe()
        except Exception as error:
            evidence['observations'].append({'result': 'refused', 'error': diagnostic(error)})
            _record(record, evidence, primary=True)
            raise
        evidence['observations'].append({'result': 'confirmed' if state is not None else 'pending'})
        if state is not None:
            evidence['outcome'] = 'state-confirmed'
            _record(record, evidence)
            return state
    evidence['outcome'] = 'unconfirmed'
    _record(record, evidence, primary=True)
    raise ReleaseFailure('Publication write outcome remains unconfirmed; inspect remote state', category='unconfirmed')


def tool_write(args, *, timeout=1800, **kwargs):
    """Run one known publication command; never repeat an ambiguous write.

Usage/auth/permission diagnostics are definitive failures. Other completed
nonzero results and timeouts can follow an applied write, so only exact remote
confirmation may resolve them. A failed process launch cannot have applied it.
"""
    try:
        result = subprocess.run(args, check=False, timeout=timeout, **kwargs)
    except subprocess.TimeoutExpired as error:
        raise ReleaseFailure('Publication tool timed out', uncertain=True, category='tool-timeout', reason=error.stderr) from None
    except OSError:
        raise ReleaseFailure('Publication tool could not start', category='tool-launch') from None
    if result.returncode:
        sanitized = _redacted(result.stderr or '')
        reason = sanitized[:1024]
        definitive = bool(re.search(r'(?i)unknown (?:flag|command)|usage:|unauthorized|authentication required|permission denied|denied:', sanitized))
        raise ReleaseFailure('Publication tool returned a failure', uncertain=not definitive,
                             category='tool-exit', exit_code=result.returncode, reason=reason)
    return result
