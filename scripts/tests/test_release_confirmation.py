"""Publication fault controls use only loopback HTTP and task-private fake CLI."""
import contextlib
import copy
import hashlib
import json
import os
from pathlib import Path
import socket
import sys
import tempfile
import threading
import unittest
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from unittest import mock
import urllib.request

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import release_candidate as candidate
import publish_prepared as publisher
from release_confirmation import RegistryReadError, ReleaseFailure, confirm, diagnostic, redact, tool_write

COMMIT = '1' * 40
ROOT = 'sha256:' + '2' * 64


@contextlib.contextmanager
def http_fixture(kind, fault):
    state = {'applied': False, 'stale': 0, 'writes': [], 'assets': [], 'latest': 'sha256:' + '3' * 64}
    draft = {'id': 7, 'tag_name': 'v1.2.3', 'target_commitish': COMMIT,
             'name': 'GridGrind 1.2.3', 'body': 'Synthetic notes.', 'draft': True, 'prerelease': True}
    class Handler(BaseHTTPRequestHandler):
        def log_message(self, *_): pass
        def send(self, code, value):
            raw = json.dumps(value).encode()
            self.send_response(code); self.send_header('Content-Length', str(len(raw))); self.end_headers(); self.wfile.write(raw)
        def do_GET(self):
            if self.path.endswith('/branches/main'):
                return self.send(200, {'protected': True, 'commit': {'sha': COMMIT}})
            visible = state['applied']
            relevant = (kind == 'tag' and '/git/ref/tags/' in self.path or kind == 'draft' and '/releases?' in self.path or kind == 'latest' and self.path == '/latest')
            if visible and relevant and fault == 'stale' and state['stale'] == 0:
                state['stale'] += 1; visible = False
            if fault == 'missing': visible = False
            if '/git/ref/tags/' in self.path:
                if not visible and kind == 'tag': return self.send(404, {})
                return self.send(200, {'object': {'type': 'commit', 'sha': '9' * 40 if fault == 'wrong' and kind == 'tag' else COMMIT}})
            if '/releases/tags/' in self.path: return self.send(404, {})
            if '/releases?' in self.path:
                row = dict(draft)
                if fault == 'wrong': row['target_commitish'] = '9' * 40
                return self.send(200, [row] if visible else [])
            if '/releases/7/assets?' in self.path: return self.send(200, state['assets'])
            if self.path == '/latest': return self.send(200, state['latest'] if visible else 'sha256:' + '3' * 64)
            self.send(404, {})
        def write(self):
            state['writes'].append((self.command, self.path))
            body = self.rfile.read(int(self.headers['Content-Length']))
            state['applied'] = True
            if self.command == 'PUT': state['latest'] = ROOT
            if fault == 'lost':
                self.connection.shutdown(socket.SHUT_RDWR); self.connection.close(); return
            self.send(201, draft if kind == 'draft' else {})
        do_POST = write
        do_PATCH = write
        do_PUT = write
    server = ThreadingHTTPServer(('127.0.0.1', 0), Handler)
    worker = threading.Thread(target=server.serve_forever, daemon=True); worker.start()
    try: yield state, 'http://127.0.0.1:' + str(server.server_port)
    finally: server.shutdown(); server.server_close(); worker.join()


class ConfirmationTest(unittest.TestCase):
    def setUp(self):
        patcher = mock.patch('release_confirmation.time.sleep')
        self.sleep = patcher.start()
        self.addCleanup(patcher.stop)

    def test_actual_http_tag_and_draft_stale_and_lost_ack_write_once(self):
        for kind in ('tag', 'draft'):
            for fault in ('normal', 'stale', 'lost'):
                with self.subTest(kind=kind, fault=fault), http_fixture(kind, fault) as (state, base):
                    api = candidate.Github('synthetic/repo', 'synthetic-token')
                    api.api.base = base
                    legacy = api.legacy
                    asset = legacy.Asset(Path('/synthetic.bin'), 'synthetic.bin', 1, 'sha256:' + hashlib.sha256(b'x').hexdigest())
                    def upload(*_):
                        state['assets'] = [{'name': asset.name, 'size': 1, 'digest': asset.digest, 'state': 'uploaded'}]
                    records = []
                    pub = legacy.Publication(api.api, 'synthetic/repo', '1.2.3', COMMIT, 'Synthetic notes.', (asset,), upload, records.append)
                    if kind == 'tag':
                        with mock.patch.object(publisher, 'guard_main'):
                            publisher.ensure_tag(api, pub, {'version': '1.2.3', 'sourceCommit': COMMIT}, records.append)
                    else: pub.run('prepare')
                    self.assertEqual(1, sum(method == 'POST' for method, _ in state['writes']))
                    self.assertEqual('state-confirmed', records[0]['outcome'])
                    self.assertEqual('failed' if fault == 'lost' else 'acknowledged', records[0]['request'])
                    if fault == 'lost': self.assertEqual('transport', records[0]['requestError']['category'])
                    if fault == 'stale': self.assertEqual('pending', records[0]['observations'][0]['result'])

    def test_actual_http_wrong_or_unconfirmed_draft_never_uploads_or_repeats(self):
        for fault in ('wrong', 'missing'):
            with self.subTest(fault=fault), http_fixture('draft', fault) as (state, base):
                api = candidate.Github('synthetic/repo', 'synthetic-token'); api.api.base = base
                uploader = mock.Mock()
                pub = api.legacy.Publication(api.api, 'synthetic/repo', '1.2.3', COMMIT, 'Synthetic notes.', (), uploader)
                with self.assertRaises(ReleaseFailure): pub.run('prepare')
                self.assertEqual(1, len(state['writes'])); uploader.assert_not_called()

    def test_actual_cli_applied_then_failed_or_stale_latest_does_not_repeat(self):
        for fault in ('normal', 'stale', 'late-exit'):
            with self.subTest(fault=fault), http_fixture('latest', fault) as (state, base), tempfile.TemporaryDirectory() as temp:
                cli = Path(temp) / 'regctl'
                cli.write_text('#!' + sys.executable + '\nimport sys,urllib.request\n' +
                               'r=urllib.request.urlopen(urllib.request.Request(' + repr(base + '/latest') + ',data=b"new",method="PUT"))\n' +
                               ('print("connection closed after upload",file=sys.stderr);sys.exit(37)\n' if fault == 'late-exit' else 'assert r.status==201\n'))
                cli.chmod(0o755)
                class Remote:
                    def tag(self, _):
                        with urllib.request.urlopen(base + '/latest') as response: return json.load(response)
                public = mock.Mock(); public.tag_digest.side_effect = lambda _: Remote().tag("latest"); records = []
                publisher.point_latest(cli, 'synthetic/repo', ROOT, Remote(), public, records.append)
                self.assertEqual([('PUT', '/latest')], state['writes'])
                self.assertGreaterEqual(public.tag_digest.call_count, 2)
                self.assertEqual('failed' if fault == 'late-exit' else 'acknowledged', records[0]['request'])
                if fault == 'late-exit': self.assertEqual(37, records[0]['requestError']['exitCode'])

    def test_programming_and_definitive_errors_never_observe(self):
        for error in (TypeError('bug'), ReleaseFailure('HTTP403', status=403), ReleaseFailure('usage', category='tool-exit')):
            observer = mock.Mock()
            with self.assertRaises(type(error)):
                confirm('operation', mock.Mock(side_effect=error), observer)
            observer.assert_not_called()

    def test_exact_prior_only_and_four_observations(self):
        write = mock.Mock(); observe = mock.Mock(return_value=None); records = []
        with self.assertRaisesRegex(ReleaseFailure, 'unconfirmed'): confirm('operation', write, observe, records.append)
        write.assert_called_once(); self.assertEqual(4, observe.call_count)
        self.assertEqual('unconfirmed', records[0]['outcome'])
        self.assertEqual([mock.call(1), mock.call(2), mock.call(3)], self.sleep.call_args_list)

    def test_safe_diagnostics_redact_before_truncation_and_preserve_primary(self):
        secret = 'private-' + 'x' * 1200
        with mock.patch.dict(os.environ, {'GH_TOKEN': secret}):
            result = redact(secret + ' Authorization: Bearer abc\nhttps://host/path?X-Signature=private')
        for forbidden in ('private', 'Bearer', 'Signature', 'xxxx'): self.assertNotIn(forbidden, result)
        error = ReleaseFailure('fixed', uncertain=False, category='http', status=403, request_id='REQ:123')
        with self.assertRaises(ReleaseFailure) as raised:
            confirm('operation', mock.Mock(side_effect=error), mock.Mock(), mock.Mock(side_effect=OSError('disk')))
        self.assertIs(error, raised.exception)
        self.assertEqual({'category': 'unexpected', 'type': 'TypeError'}, diagnostic(TypeError(secret)))

    def test_bad_cli_usage_fails_without_confirmation_even_after_long_prefix(self):
        for prefix in (b'', b'x' * 1100):
            with self.subTest(prefix_length=len(prefix)):
                result = mock.Mock(returncode=2, stderr=prefix + b'usage: regctl image copy: unknown flag', stdout=b'')
                observer = mock.Mock(return_value=ROOT)
                with mock.patch('release_confirmation.subprocess.run', return_value=result):
                    with self.assertRaises(ReleaseFailure):
                        confirm('copy', lambda: tool_write(['pinned-regctl'], capture_output=True), observer)
                observer.assert_not_called()

    def test_public_only_lag_and_third_latest_root_are_not_conflated(self):
        old = 'sha256:' + '3' * 64
        for final, succeeds in ((ROOT, True), ('sha256:' + '9' * 64, False)):
            with self.subTest(final=final):
                remote = mock.Mock(); remote.tag.side_effect = [old, ROOT, ROOT]
                public = mock.Mock(); public.tag_digest.side_effect = [old, old, final]
                records = []
                with mock.patch.object(publisher, 'tool_write') as write:
                    if succeeds:
                        publisher.point_latest('pinned-regctl', 'synthetic/repo', ROOT, remote, public, records.append)
                    else:
                        with self.assertRaisesRegex(candidate.CandidateError, 'unexpected root'):
                            publisher.point_latest('pinned-regctl', 'synthetic/repo', ROOT, remote, public, records.append)
                    write.assert_called_once()
                self.assertEqual('pending', records[0]['observations'][0]['result'])
                self.assertEqual('confirmed' if succeeds else 'refused', records[0]['observations'][1]['result'])

    def test_http_status_request_id_and_unknown_failure_diagnostics_are_distinct(self):
        legacy = candidate.publication_module()
        for status in (403, 422, 503):
            error = legacy.ApiError(status, 'POST', 'REQ:123')
            details = diagnostic(error)
            self.assertEqual(status, details['httpStatus']); self.assertEqual('REQ:123', details['requestId'])
            self.assertEqual(status == 503, error.uncertain)
        self.assertNotIn('credential', json.dumps(diagnostic(TypeError('credential'))))

    def test_expired_registry_token_renews_one_fixed_scope_get_only(self):
        import release_oci
        import urllib.error
        import urllib.parse
        verify = publisher.verify_module()
        payload = b'{"schemaVersion":2,"synthetic":true}'
        expected = 'sha256:' + hashlib.sha256(payload).hexdigest()
        for owner in ('public', 'authenticated', 'streaming'):
            for mode in ('expired', 'second401', 'wrong', 'absent'):
                with self.subTest(owner=owner, mode=mode):
                    state = {'tokens': 0, 'reads': 0, 'methods': []}
                    class Handler(BaseHTTPRequestHandler):
                        def log_message(self, *_): pass
                        def do_GET(self):
                            state['methods'].append(self.command)
                            query = urllib.parse.urlsplit(self.path)
                            if query.path == '/token':
                                scope = urllib.parse.parse_qs(query.query)['scope']
                                self.server.test.assertEqual(['repository:synthetic/repo:' + ('pull,push' if owner == 'authenticated' else 'pull')], scope)
                                state['tokens'] += 1; raw = json.dumps({'token': 'synthetic-' + str(state['tokens'])}).encode(); code = 200
                            else:
                                self.server.test.assertEqual('/v2/synthetic/repo/manifests/latest', query.path)
                                state['reads'] += 1
                                if mode == 'absent': code, raw = 404, b'{}'
                                elif mode == 'second401' or self.headers.get('Authorization') == 'Bearer synthetic-1': code, raw = 401, b'{}'
                                else: code, raw = 200, payload if mode != 'wrong' else b'wrong'
                            self.send_response(code)
                            # An untrusted challenge must not change token-service identity.
                            if code == 401: self.send_header('WWW-Authenticate', 'Bearer realm="https://untrusted.invalid/token"')
                            self.send_header('Content-Length', str(len(raw))); self.end_headers(); self.wfile.write(raw)
                    server = ThreadingHTTPServer(('127.0.0.1', 0), Handler); server.test = self
                    worker = threading.Thread(target=server.serve_forever, kwargs={'poll_interval': .01}, daemon=True); worker.start()
                    real = urllib.request.build_opener(); base = 'http://127.0.0.1:' + str(server.server_port)
                    class Routing:
                        def open(self, request, timeout=30):
                            url = request if isinstance(request, str) else request.full_url
                            parts = urllib.parse.urlsplit(url)
                            self_outer.assertEqual('https', parts.scheme); self_outer.assertEqual('ghcr.io', parts.hostname)
                            headers = {} if isinstance(request, str) else dict(request.header_items())
                            local = urllib.request.Request(base + parts.path + ('?' + parts.query if parts.query else ''), headers=headers)
                            try:
                                return real.open(local, timeout=timeout)
                            except urllib.error.HTTPError as error:
                                error.url = url
                                raise
                    self_outer = self
                    try:
                        with mock.patch.object(urllib.request, 'build_opener', return_value=Routing()):
                            client = release_oci.Ghcr('synthetic/repo', 'synthetic', 'not-a-credential') if owner == 'authenticated' else verify.Registry('synthetic/repo')
                        def read():
                            if owner == 'authenticated': self.assertEqual(expected, client.tag('latest'))
                            elif owner == 'streaming':
                                with client.open('manifests', 'latest') as response:
                                    self.assertEqual(expected, 'sha256:' + hashlib.sha256(response.read()).hexdigest())
                            else: client.verify_tag('latest', expected)
                        if mode == 'expired': read()
                        else:
                            with self.assertRaises((ValueError, candidate.CandidateError, AssertionError)): read()
                        self.assertEqual(1 if mode == 'absent' else 2, state['tokens'])
                        self.assertEqual(1 if mode == 'absent' else 2, state['reads'])
                        self.assertEqual({'GET'}, set(state['methods']))
                    finally: server.shutdown(); server.server_close(); worker.join()

    def test_registry_http_status_and_request_id_survive_diagnostics(self):
        error = RegistryReadError(401, 'REQ:401')
        self.assertIsInstance(error, ValueError)
        self.assertEqual({'category': 'http', 'type': 'RegistryReadError', 'httpStatus': 401,
                          'requestId': 'REQ:401', 'message': 'Registry read failed with HTTP 401'}, diagnostic(error))

    def test_redirected_cdn_401_does_not_renew_repository_token(self):
        import urllib.error
        verify = publisher.verify_module()
        for operation in ('metadata', 'streaming'):
            with self.subTest(operation=operation):
                client = verify.Registry.__new__(verify.Registry)
                client.repository = 'synthetic/repo'; client.token = 'synthetic'
                client.refresh = mock.Mock()
                opener = mock.Mock()
                opener.open.side_effect = urllib.error.HTTPError('https://cdn.invalid/blob?signature=private', 401, 'unauthorized', {}, None)
                client.metadata = client.blobs = opener
                with self.assertRaises(RegistryReadError) as raised:
                    if operation == 'metadata': client.tag_digest('latest')
                    else: client.open('blobs', ROOT)
                self.assertEqual(401, diagnostic(raised.exception)['httpStatus'])
                client.refresh.assert_not_called(); opener.open.assert_called_once()
