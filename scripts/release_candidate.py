#!/usr/bin/env python3
"""Immutable release candidate, service receipt and reviewed-pair contracts.

These files describe CI candidates visible to authenticated artifact readers.
An independent acknowledgement records the dispatcher's reviewed evidence;
it does not claim that GitHub executed the independent review.
"""
import argparse
import hashlib
import importlib.util
import json
import os
from pathlib import Path, PurePosixPath
import re
import stat
import sys
import urllib.error
import urllib.parse
import urllib.request
import zipfile

MAX_ARCHIVE = 16 * 1024**3
MAX_EXPANDED = 32 * 1024**3
MAX_FILES = 20000
WORKFLOW_PATH = '.github/workflows/release.yml'
REQUIRED_JOBS = {'Prepare candidate', 'Bind candidate QA receipt'}
SHA = re.compile(r'sha256:[0-9a-f]{64}\Z')
COMMIT = re.compile(r'[0-9a-f]{40}\Z')
VERSION = re.compile(r'(?:0|[1-9][0-9]*)\.(?:0|[1-9][0-9]*)\.(?:0|[1-9][0-9]*)\Z')

class CandidateError(ValueError):
    """An unproven candidate cannot acquire publication authority."""

def require(condition, message):
    if not condition:
        raise CandidateError(message)

def canonical(value):
    return json.dumps(value, sort_keys=True, separators=(',', ':'), ensure_ascii=False, allow_nan=False).encode('utf-8')

def sha_bytes(value):
    return 'sha256:' + hashlib.sha256(value).hexdigest()

def sha_file(path):
    h = hashlib.sha256()
    with Path(path).open('rb') as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b''):
            h.update(block)
    return 'sha256:' + h.hexdigest()

def digest(value):
    require(isinstance(value, str) and SHA.fullmatch(value), 'Expected a typed sha256 digest')
    return value

def positive(value):
    require(type(value) is int and value > 0, 'Expected a positive service identity')
    return value

def exact(value, keys):
    require(isinstance(value, dict) and set(value) == set(keys), 'Unexpected or missing contract fields')
    return value

def read_json(path, limit=8 * 1024 * 1024):
    path = Path(path)
    require(path.is_file() and not path.is_symlink() and path.stat().st_size <= limit, 'Missing or oversized JSON input')
    def pairs(rows):
        result = {}
        for key, value in rows:
            require(key not in result, 'Duplicate JSON member')
            result[key] = value
        return result
    return json.loads(path.read_text(encoding='utf-8'), object_pairs_hook=pairs,
                      parse_constant=lambda _: (_ for _ in ()).throw(CandidateError('Nonfinite JSON number')))

def write_json(path, value):
    path = Path(path)
    require(not path.exists(), 'Refusing to overwrite an immutable receipt')
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(canonical(value) + b'\n')

def relative(value):
    require(isinstance(value, str) and value and '\\' not in value and not any(ord(c) < 32 or ord(c) == 127 for c in value),
            'Invalid relative inventory path')
    p = PurePosixPath(value)
    require(not p.is_absolute() and all(part not in ('', '.', '..') for part in value.split('/'))
            and p.as_posix() == value, 'Noncanonical inventory path')
    return value

def inventory(root, excluded=()):
    root = Path(root)
    require(root.is_dir() and not root.is_symlink(), 'Payload must be a real directory')
    rows = []
    for path in sorted(root.rglob('*')):
        require(not path.is_symlink(), 'Payload symlinks are forbidden')
        if path.is_dir():
            continue
        require(path.is_file() and path.stat().st_nlink == 1, 'Payload must contain regular singly-linked files')
        name = relative(path.relative_to(root).as_posix())
        if name in excluded:
            continue
        rows.append({'path': name, 'size': path.stat().st_size, 'sha256': sha_file(path)})
    require(rows and len(rows) <= MAX_FILES, 'Empty or excessive delivery inventory')
    return rows

def validate_inventory(rows):
    require(isinstance(rows, list) and 0 < len(rows) <= MAX_FILES, 'Invalid inventory')
    names = []
    total = 0
    for row in rows:
        exact(row, ('path', 'size', 'sha256'))
        names.append(relative(row['path']))
        require(type(row['size']) is int and 0 <= row['size'] <= MAX_ARCHIVE, 'Invalid inventory size')
        total += row['size']
        digest(row['sha256'])
    require(names == sorted(set(names)) and total <= MAX_EXPANDED, 'Duplicate, unsorted or excessive inventory')

def identity(value):
    exact(value, ('repository', 'version', 'sourceCommit', 'epoch', 'workflowId', 'workflowPath', 'workflowCommit', 'runId', 'runAttempt'))
    require(isinstance(value['repository'], str) and re.fullmatch(r'[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+', value['repository']), 'Invalid repository')
    require(isinstance(value['version'], str) and VERSION.fullmatch(value['version']), 'Invalid stable version')
    require(isinstance(value['sourceCommit'], str) and COMMIT.fullmatch(value['sourceCommit']), 'Invalid source commit')
    require(value['workflowCommit'] == value['sourceCommit'], 'Workflow controls and prepared source must be the same main commit')
    require(value['workflowPath'] == WORKFLOW_PATH, 'Unexpected release workflow')
    require(type(value['epoch']) is int and value['epoch'] > 0, 'Invalid commit epoch')
    for key in ('workflowId', 'runId', 'runAttempt'):
        positive(value[key])
    return value

def validate_checks(value):
    exact(value, ('formatVersion', 'status', 'suites', 'negativeControls'))
    require(type(value['formatVersion']) is int and value['formatVersion'] == 1 and value['status'] == 'PASS', 'Automated artifact checks did not pass')
    required = {'deterministic', 'active-fuzz', 'bounded-pit', 'runtime-amd64', 'runtime-arm64', 'devcontainer',
                'delivery-content', 'oci-source-sbom'}
    require(isinstance(value['suites'], list), 'Missing executed suite records')
    names = []
    for row in value['suites']:
        exact(row, ('id', 'status', 'tests', 'failures', 'errors', 'skipped', 'evidence'))
        names.append(row['id'])
        require(row['status'] == 'PASS' and type(row['tests']) is int and row['tests'] > 0
                and all(type(row[k]) is int and row[k] == 0 for k in ('failures', 'errors', 'skipped')),
                'Missing, failed, empty or skipped portable artifact suite')
        require(isinstance(row['evidence'], list) and row['evidence'], 'Suite has no retained evidence')
        for entry in row['evidence']:
            relative(entry)
    require(len(names) == len(set(names)) and set(names) == required, 'Required suite discovery mismatch')
    require(isinstance(value['negativeControls'], list) and value['negativeControls'], 'Missing negative-control evidence')
    for name in value['negativeControls']:
        relative(name)
    return value

def candidate(root):
    root = Path(root)
    value = read_json(root / 'candidate.json')
    exact(value, ('formatVersion', 'role', 'identity', 'inventory', 'inventorySha256', 'inputHashes', 'oci', 'checks'))
    require(type(value['formatVersion']) is int and value['formatVersion'] == 1 and value['role'] == 'delivery', 'Wrong candidate role/version')
    identity(value['identity'])
    validate_inventory(value['inventory'])
    require(value['inventorySha256'] == sha_bytes(canonical(value['inventory'])), 'Inventory digest mismatch')
    require(value['inventory'] == inventory(root, ('candidate.json',)), 'Payload missing, extra, changed or indirect files')
    require(isinstance(value['inputHashes'], list) and value['inputHashes'], 'Missing authoritative input hashes')
    validate_inventory(value['inputHashes'])
    exact(value['oci'], ('runtime', 'source'))
    for role in ('runtime', 'source'):
        row = value['oci'][role]
        exact(row, ('path', 'root', 'graphSha256'))
        relative(row['path']); digest(row['root']); digest(row['graphSha256'])
        require(row['path'] in {r['path'] for r in value['inventory']}, 'OCI payload is not inventoried')
    checks = validate_checks(value['checks'])
    names = {row['path'] for row in value['inventory']}
    for row in checks['suites']:
        require(set(row['evidence']) <= names, 'Suite evidence is absent')
    require(set(checks['negativeControls']) <= names, 'Negative control evidence is absent')
    return value

def service_artifact(value, expected_name, expected_identity):
    require(isinstance(value, dict), 'Invalid artifact service record')
    positive(value.get('id')); digest(value.get('digest'))
    require(value.get('name') == expected_name and value.get('expired') is False, 'Wrong role, attempt, or expired artifact')
    require(type(value.get('size_in_bytes')) is int and 0 < value['size_in_bytes'] <= MAX_ARCHIVE, 'Invalid artifact transport size')
    run = value.get('workflow_run')
    require(isinstance(run, dict) and run.get('id') == expected_identity['runId']
            and run.get('head_sha') == expected_identity['sourceCommit']
            and run.get('head_branch') == 'main', 'Artifact belongs to another run/source/branch')
    return {'id': value['id'], 'digest': value['digest']}

def role_name(role, value):
    return '%s-%s-%s-%s' % (role, value['version'], value['runId'], value['runAttempt'])

def validate_receipt(root, delivery, delivery_service):
    root = Path(root)
    value = read_json(root / 'artifact-qa.json')
    exact(value, ('formatVersion', 'role', 'identity', 'delivery', 'candidateSha256', 'inventorySha256', 'checks', 'transportVerified'))
    require(type(value['formatVersion']) is int and value['formatVersion'] == 1 and value['role'] == 'artifact-qa' and value['transportVerified'] is True, 'Invalid QA receipt')
    require(value['identity'] == delivery['identity'] and value['delivery'] == delivery_service, 'QA receipt service/run/attempt binding mismatch')
    require(value['inventorySha256'] == delivery['inventorySha256'] and value['checks'] == delivery['checks'], 'QA changed earlier completed checks')
    require({p.name for p in root.iterdir()} == {'artifact-qa.json'}, 'Unexpected QA role payload')
    digest(value['candidateSha256'])
    return value

def reviewed_pair(ident, delivery_service, candidate_sha, qa_service, receipt_sha):
    identity(ident)
    require(delivery_service['id'] != qa_service['id'], 'Delivery and QA must be distinct immutable artifacts')
    for service in (delivery_service, qa_service):
        exact(service, ('id', 'digest')); positive(service['id']); digest(service['digest'])
    return {'formatVersion': 1, 'identity': ident, 'delivery': delivery_service, 'candidateSha256': digest(candidate_sha),
            'qa': qa_service, 'receiptSha256': digest(receipt_sha)}

def independent_acceptance(value, expected_sha, pair_sha, jar_sha):
    require(len(canonical(value)) <= 32768 and sha_bytes(canonical(value)) == digest(expected_sha), 'Independent acknowledgement digest mismatch')
    exact(value, ('formatVersion', 'role', 'pairSha256', 'jarSha256', 'status', 'evidenceSha256'))
    require(type(value['formatVersion']) is int and value['formatVersion'] == 1 and value['role'] == 'independent-acceptance'
            and value['pairSha256'] == pair_sha and value['jarSha256'] == jar_sha and value['status'] == 'PASS',
            'Independent acknowledgement binds another pair/JAR or did not pass')
    digest(value['evidenceSha256'])
    return value


def extract_zip(archive, output, expected_digest):
    archive, output = Path(archive), Path(output)
    require(sha_file(archive) == digest(expected_digest) and archive.stat().st_size <= MAX_ARCHIVE, 'Service artifact transport digest mismatch')
    require(not output.exists(), 'Artifact extraction destination must be absent')
    with zipfile.ZipFile(archive) as z:
        entries = z.infolist(); names = []; total = 0
        require(0 < len(entries) <= MAX_FILES, 'Excessive or empty artifact archive')
        for entry in entries:
            name = relative(entry.filename.rstrip('/')); names.append(name)
            mode = (entry.external_attr >> 16) & 0xffff
            require(not entry.flag_bits & 1 and not stat.S_ISLNK(mode)
                    and (stat.S_IFMT(mode) in (0, stat.S_IFREG, stat.S_IFDIR)), 'Artifact links, devices and encryption forbidden')
            total += entry.file_size
            require(total <= MAX_EXPANDED and entry.file_size <= MAX_ARCHIVE, 'Artifact expansion exceeds bound')
        require(len(names) == len(set(names)), 'Duplicate archive path')
        # Validate the complete namespace before creating any files.
        files = {entry.filename for entry in entries if not entry.is_dir()}
        require(not any(str(parent) in files for name in names for parent in PurePosixPath(name).parents if str(parent) != '.'), 'Archive file/directory collision')
        output.mkdir(parents=True)
        for entry in entries:
            target = output / entry.filename
            if entry.is_dir():
                target.mkdir(parents=True, exist_ok=True); continue
            target.parent.mkdir(parents=True, exist_ok=True)
            with z.open(entry) as source, target.open('xb') as dest:
                copied = 0
                for block in iter(lambda: source.read(1024 * 1024), b''):
                    copied += len(block); require(copied <= entry.file_size, 'Archive member exceeds declared size'); dest.write(block)
            require(copied == entry.file_size, 'Truncated archive member')

# Reuse the existing bounded API/no-ambiguous-mutation implementation.
def publication_module():
    spec = importlib.util.spec_from_file_location('_gridgrind_publication', Path(__file__).with_name('release-publication.py'))
    module = importlib.util.module_from_spec(spec); sys.modules[spec.name] = module; spec.loader.exec_module(module)
    return module

class Github:
    def __init__(self, repo, token):
        require(re.fullmatch(r'[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+', repo) and all(p not in ('.', '..') for p in repo.split('/')), 'Invalid GitHub repository')
        self.legacy = publication_module(); self.api = self.legacy.Github(repo, token)
    def get(self, path):
        return self.api.request('GET', path)
    def mutate(self, method, path, body):
        return self.api.request(method, path, body)
    def artifact(self, identifier):
        return self.get('/actions/artifacts/' + str(positive(identifier)))
    def download(self, identifier, output, expected_digest):
        require(not Path(output).exists(), 'Refusing to overwrite downloaded artifact')
        req = urllib.request.Request(self.api.base + '/actions/artifacts/%s/zip' % positive(identifier),
                                     headers={'Authorization': 'Bearer ' + self.api.token, 'Accept': 'application/vnd.github+json',
                                              'X-GitHub-Api-Version': '2026-03-10'})
        class NoFollow(urllib.request.HTTPRedirectHandler):
            def redirect_request(self, *_):
                return None
        try:
            urllib.request.build_opener(NoFollow()).open(req, timeout=30)
        except urllib.error.HTTPError as error:
            require(error.code == 302, 'Artifact download did not return the expected service redirect')
            location = error.headers.get('Location'); error.close()
        else:
            raise CandidateError('Artifact transport unexpectedly omitted its service redirect')
        # Authorization goes only to GitHub API, never to signed storage URLs.
        class HttpsOnly(urllib.request.HTTPRedirectHandler):
            def redirect_request(self, request, fp, code, message, headers, url):
                require(urllib.parse.urlsplit(url).scheme == 'https', 'Artifact redirect must remain HTTPS')
                return super().redirect_request(request, fp, code, message, headers, url)
        parsed = urllib.parse.urlsplit(location or '')
        require(parsed.scheme == 'https' and parsed.hostname and not parsed.username and not parsed.fragment, 'Invalid signed service redirect')
        total = 0; h = hashlib.sha256()
        with urllib.request.build_opener(HttpsOnly()).open(location, timeout=60) as response, Path(output).open('xb') as dest:
            for block in iter(lambda: response.read(1024 * 1024), b''):
                total += len(block); require(total <= MAX_ARCHIVE, 'Artifact exceeds transfer bound'); h.update(block); dest.write(block)
        require('sha256:' + h.hexdigest() == digest(expected_digest), 'Downloaded artifact service digest mismatch')

    def validate_run(self, ident, completed):
        wf = self.get('/actions/workflows/release.yml')
        require(wf.get('id') == ident['workflowId'] and wf.get('path') == WORKFLOW_PATH, 'Unexpected trusted workflow identity')
        run = self.get('/actions/runs/' + str(ident['runId']))
        require(run.get('workflow_id') == ident['workflowId'] and run.get('event') == 'workflow_dispatch'
                and run.get('head_branch') == 'main' and run.get('head_sha') == ident['sourceCommit']
                and run.get('run_attempt') == ident['runAttempt'], 'Wrong workflow/event/branch/source or stale attempt')
        if completed:
            require(run.get('status') == 'completed' and run.get('conclusion') == 'success', 'Preparation run did not finish successfully')
            jobs = []
            for page in range(1, 101):
                data = self.get('/actions/runs/%s/attempts/%s/jobs?per_page=100&page=%s' % (ident['runId'], ident['runAttempt'], page))
                rows = data.get('jobs'); require(isinstance(rows, list), 'Invalid job service record'); jobs.extend(rows)
                if len(rows) < 100: break
            else: raise CandidateError('Job pagination exceeds bound')
            for name in REQUIRED_JOBS:
                matches = [j for j in jobs if j.get('name') == name]
                require(len(matches) == 1 and matches[0].get('status') == 'completed'
                        and matches[0].get('conclusion') == 'success', 'Required preparation/QA job did not succeed')
        return run

def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument('command', choices=('pair',))
    p.add_argument('--candidate', type=Path, required=True); p.add_argument('--delivery-service', type=Path, required=True)
    p.add_argument('--qa-service', type=Path, required=True); p.add_argument('--qa-receipt', type=Path, required=True)
    args = p.parse_args(); delivery = candidate(args.candidate)
    pair = reviewed_pair(delivery['identity'], read_json(args.delivery_service), sha_file(args.candidate/'candidate.json'),
                         read_json(args.qa_service), sha_file(args.qa_receipt))
    print(json.dumps({'pair': pair, 'pairSha256': sha_bytes(canonical(pair))}, indent=2))

if __name__ == '__main__':
    try: main()
    except (CandidateError, OSError, ValueError, KeyError, zipfile.BadZipFile):
        raise SystemExit('Release candidate validation failed; inspect retained inputs without mutation') from None
