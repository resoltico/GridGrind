#!/usr/bin/env python3
"""Inspect immutable OCI roots and copy them with the qualified workflow-only client.

The publication identity is the selected content-addressed index, not index.json
locator bytes. Regctl may maintain bounded OCI-subject referrer indexes/tags;
those are metadata, never replacements for the reviewed publication graph.
"""
import argparse
import base64
import gzip
import hashlib
import io
import json
import os
from pathlib import Path
import re
import subprocess
import tarfile
import urllib.error
import urllib.parse
import urllib.request
from release_confirmation import RegistryReadError, confirm, tool_write
from release_candidate import (CandidateError, canonical, digest, read_json, require, sha_bytes, sha_file, write_json)

INDEX = 'application/vnd.oci.image.index.v1+json'
MANIFEST = 'application/vnd.oci.image.manifest.v1+json'
CONFIG = 'application/vnd.oci.image.config.v1+json'
ATTESTATION = 'application/vnd.in-toto+json'
MAX_BLOB = 8 * 1024**3
MAX_GRAPH = 10000

def command(args, **kwargs):
    result = subprocess.run(args, check=False, timeout=1800, **kwargs)
    require(result.returncode == 0, 'OCI consumer failed; inspect state before selecting a retry')
    return result

def properties(path):
    return dict(line.split('=', 1) for line in Path(path).read_text().splitlines() if line and not line.startswith('#'))

def install_client(pins_path, output):
    pins = properties(pins_path)
    arch = os.uname().machine
    require(arch in ('x86_64', 'aarch64'), 'Release OCI consumer requires qualified Linux amd64/arm64')
    arch = 'amd64' if arch == 'x86_64' else 'arm64'
    url = 'https://github.com/regclient/regclient/releases/download/v%s/regctl-linux-%s' % (pins['REGCTL_VERSION'], arch)
    output = Path(output); require(not output.exists(), 'OCI client output must be absent')
    output.parent.mkdir(parents=True, exist_ok=True)
    with urllib.request.urlopen(url, timeout=60) as response, output.open('xb') as dest:
        size = 0
        for block in iter(lambda: response.read(1024 * 1024), b''):
            size += len(block); require(size <= 64 * 1024 * 1024, 'Oversized OCI tool'); dest.write(block)
    require(sha_file(output) == 'sha256:' + pins['REGCTL_' + arch.upper() + '_SHA256'], 'OCI tool publisher checksum mismatch')
    output.chmod(0o755)
    result = command([str(output), 'version'], capture_output=True, text=True)
    require('VCSTag:     v' + pins['REGCTL_VERSION'] in result.stdout
            and pins['REGCTL_SOURCE_COMMIT'] in result.stdout, 'OCI tool release/source identity mismatch')

def unpack(archive, output):
    archive, output = Path(archive), Path(output)
    require(not output.exists(), 'OCI layout destination must be absent')
    with tarfile.open(archive) as tar:
        entries = tar.getmembers(); total = 0; seen = set()
        require(len(entries) <= MAX_GRAPH + 10, 'Excessive OCI archive entries')
        for member in entries:
            name = member.name.rstrip('/')
            require(name not in seen, 'Duplicate OCI archive path'); seen.add(name)
            require(member.isdir() or member.isfile(), 'OCI links/devices forbidden')
            valid = name in {'index.json', 'oci-layout', 'blobs', 'blobs/sha256'} or re.fullmatch(r'blobs/sha256/[0-9a-f]{64}', name)
            require(valid and member.size <= MAX_BLOB, 'Unexpected OCI archive entry')
            total += member.size
            require(total <= 32 * 1024**3, 'OCI expansion exceeds bound')
        output.mkdir(parents=True)
        tar.extractall(output, filter='data')
    require(read_json(output/'oci-layout') == {'imageLayoutVersion': '1.0.0'}, 'Unsupported OCI layout')
    top = read_json(output/'index.json')
    require(isinstance(top.get('manifests'), list) and len(top['manifests']) == 1, 'Producer OCI must select exactly one root')
    root = top['manifests'][0]
    require(root.get('mediaType') == INDEX, 'Expected multi-platform OCI publication index')
    return root

def inspect(layout, root, role):
    layout = Path(layout); nodes = {}; edges = []
    def blob(desc):
        name = digest(desc.get('digest')); size = desc.get('size')
        require(type(size) is int and 0 < size <= MAX_BLOB, 'Invalid OCI descriptor size')
        path = layout/'blobs/sha256'/name[7:]
        require(path.is_file() and not path.is_symlink() and path.stat().st_size == size and sha_file(path) == name, 'Missing/mutated OCI blob')
        return path
    def walk(desc, parent=None, slot='root'):
        require(len(nodes) <= MAX_GRAPH, 'OCI graph exceeds bound')
        path = blob(desc); name = desc['digest']; edges.append({'parent': parent, 'slot': slot, 'descriptor': desc})
        if name in nodes: return
        nodes[name] = {'digest': name, 'size': desc['size'], 'mediaType': desc['mediaType']}
        if desc['mediaType'] in (INDEX, MANIFEST):
            value = read_json(path)
            require(value.get('schemaVersion') == 2 and value.get('mediaType') == desc['mediaType'], 'OCI manifest type mismatch')
            for key in ('manifests', 'layers'):
                for i, child in enumerate(value.get(key, [])): walk(child, name, key + ':' + str(i))
            for key in ('config', 'subject'):
                if key in value: walk(value[key], name, key)
    selected = {key: root[key] for key in ('mediaType', 'digest', 'size')}
    walk(selected)
    index = read_json(blob(root)); runtime = {}; attestations = {}
    for desc in index['manifests']:
        platform = desc.get('platform', {})
        if platform.get('os') == 'linux':
            name = 'linux/' + platform.get('architecture', '')
            require(name in {'linux/amd64', 'linux/arm64'} and name not in runtime, 'Unexpected/duplicate runtime platform')
            runtime[name] = desc
        else:
            ann = desc.get('annotations', {})
            require(role == 'runtime' and ann.get('vnd.docker.reference.type') == 'attestation-manifest', 'Unexpected OCI index member')
            target = digest(ann.get('vnd.docker.reference.digest'))
            require(target not in attestations, 'Duplicate platform attestation')
            attestations[target] = desc
    require(set(runtime) == {'linux/amd64', 'linux/arm64'}, 'Both runtime platforms are required')
    statements = {}
    for platform, desc in runtime.items():
        manifest = read_json(blob(desc)); config = read_json(blob(manifest['config']))
        require(config.get('os') == 'linux' and 'linux/' + config.get('architecture', '') == platform, 'OCI config/platform mismatch')
        if role == 'source':
            settings = config.get('config', {})
            require(not settings.get('Entrypoint') and not settings.get('Cmd'), 'Source image must remain data-only')
        if role == 'runtime':
            require(desc['digest'] in attestations, 'Platform is missing indexed attestation')
            att = read_json(blob(attestations[desc['digest']]))
            require(att.get('subject', {}).get('digest') == desc['digest'], 'Attestation manifest subject mismatch')
            kinds = {}
            for layer in att['layers']:
                require(layer.get('mediaType') == ATTESTATION, 'Unsupported attestation layer')
                value = read_json(blob(layer), 64 * 1024 * 1024)
                require(value.get('_type') in {'https://in-toto.io/Statement/v0.1', 'https://in-toto.io/Statement/v1'}
                        and any(s.get('digest', {}).get('sha256') == desc['digest'][7:] for s in value.get('subject', [])), 'In-toto statement subject mismatch')
                kind = value.get('predicateType'); require(kind not in kinds, 'Duplicate attestation predicate'); kinds[kind] = value
            require(set(kinds) == {'https://spdx.dev/Document', 'https://slsa.dev/provenance/v1'}, 'Required actual SPDX/SLSA predicates missing')
            require(kinds['https://slsa.dev/provenance/v1'].get('predicate', {}).get('buildDefinition', {}).get('buildType'), 'Empty provenance build definition')
            statements[platform] = kinds
    require(set(attestations) == ({d['digest'] for d in runtime.values()} if role == 'runtime' else set()), 'Unindexed or unrelated attestation')
    physical = {p.name for p in (layout/'blobs/sha256').iterdir()}
    # Producer layouts may contain only the selected root closure, never inferred external referrers.
    require(physical == {name[7:] for name in nodes}, 'Producer contains unselected OCI blobs')
    graph = {'root': selected, 'nodes': nodes, 'edges': edges}
    return {'graph': graph, 'graphSha256': sha_bytes(canonical(graph)), 'platforms': runtime, 'statements': statements}

def check_sbom(inspection, components):
    expected = {'pkg:maven/%s/%s@%s' % (c['group'], c['artifact'], c['version']) for c in components}
    require(expected and any('/native-provider@' in p for p in expected), 'Missing resolved runtime/capsule inventory')
    forbidden = ('pkg:maven/org.junit', 'pkg:maven/org.pitest', 'pkg:maven/edu.berkeley.cs.jqf', 'pkg:maven/net.sourceforge.pmd',
                 'pkg:maven/com.tngtech.archunit', 'pkg:maven/com.diffplug.spotless', 'pkg:maven/com.uber.nullaway',
                 'pkg:maven/biz.aQute', 'pkg:maven/org.osgi/', 'pkg:maven/org.gradle/', 'pkg:maven/com.google.errorprone/error_prone_core')
    for platform, statements in inspection['statements'].items():
        spdx = statements['https://spdx.dev/Document']['predicate']
        require(isinstance(spdx.get('packages'), list) and spdx['packages'], 'Empty actual scanner package inventory')
        actual = {ref['referenceLocator'] for package in spdx['packages'] for ref in package.get('externalRefs', []) if ref.get('referenceType') == 'purl'}
        require(expected <= actual, 'Actual scanner omitted resolved runtime components: ' + platform)
        require(not any(p.startswith(forbidden) for p in actual), 'Build/test tooling appears in runtime SBOM')

def source_payload(layout, inspection, expected_files):
    expected = {p.name: (p.stat().st_size, sha_file(p)) for p in expected_files}
    for desc in inspection['platforms'].values():
        manifest = read_json(Path(layout)/'blobs/sha256'/desc['digest'][7:]); found = set()
        require(len(manifest['layers']) == 1, 'Data-only source image must have one payload layer')
        layer = Path(layout)/'blobs/sha256'/manifest['layers'][0]['digest'][7:]
        with tarfile.open(layer) as archive:
            for entry in archive:
                if entry.isdir(): continue
                require(entry.isfile() and entry.name.startswith('sources/'), 'Unexpected source-image filesystem entry')
                name = entry.name[len('sources/'):]
                require(name in expected and name not in found, 'Unexpected/duplicate source payload')
                stream = archive.extractfile(entry); h = hashlib.sha256(); total = 0
                for block in iter(lambda: stream.read(1024 * 1024), b''): total += len(block); h.update(block)
                require((total, 'sha256:' + h.hexdigest()) == expected[name], 'Source image payload changed')
                found.add(name)
        require(found == set(expected), 'Source image payload incomplete')

def load_for_testing(archive, inspection, tag, server_version):
    version = command(['docker', 'version', '--format', '{{.Server.Version}}'], capture_output=True, text=True).stdout.strip()
    require(version == server_version, 'Unqualified Docker server version for OCI import')
    driver = command(['docker', 'info', '--format', '{{json .DriverStatus}}'], capture_output=True, text=True).stdout
    require('io.containerd.snapshotter.v1' in driver, 'Descriptor-preserving containerd store required')
    command(['docker', 'image', 'load', '--input', str(archive)], capture_output=True)
    actual = command(['docker', 'image', 'inspect', tag, '--format', '{{.Id}}'], capture_output=True, text=True).stdout.strip()
    require(actual == inspection['graph']['root']['digest'], 'Docker import changed reviewed index')
    check = Path(archive).with_suffix('.store-check.tar')
    require(not check.exists(), 'Store verification export must be absent')
    command(['docker', 'image', 'save', '--output', str(check), tag], capture_output=True)
    with tarfile.open(check) as saved:
        locator = json.load(saved.extractfile('index.json'))
        require(any(d['digest'] == actual for d in locator['manifests']), 'Store export lost selected root')
        for name, row in inspection['graph']['nodes'].items():
            stream = saved.extractfile('blobs/sha256/' + name[7:]); require(stream is not None, 'Store export lost reviewed blob')
            h = hashlib.sha256(); size = 0
            for block in iter(lambda: stream.read(1024 * 1024), b''): size += len(block); h.update(block)
            require(size == row['size'] and 'sha256:' + h.hexdigest() == name, 'Store import/export changed a reviewed blob')
    check.unlink()
    return actual

class Ghcr:
    """Read-only, authenticated tag inspection; no custom registry uploader."""
    def __init__(self, repository, actor, token):
        require(re.fullmatch(r'[a-z0-9._-]+/[a-z0-9._-]+', repository), 'Invalid GHCR repository')
        self.repository = repository
        class NoRedirect(urllib.request.HTTPRedirectHandler):
            def redirect_request(self, *_): raise CandidateError('Registry metadata must not redirect credentials')
        self.opener = urllib.request.build_opener(NoRedirect())
        query = urllib.parse.urlencode({'service': 'ghcr.io', 'scope': 'repository:' + repository + ':pull,push'})
        auth = base64.b64encode((actor + ':' + token).encode()).decode()
        self.auth_request = urllib.request.Request('https://ghcr.io/token?' + query, headers={'Authorization': 'Basic ' + auth})
        self.refresh()
    def refresh(self):
        with self.opener.open(self.auth_request, timeout=30) as response: value = json.load(response)
        self.token = value.get('token'); require(isinstance(self.token, str) and self.token, 'No registry service token')
    def tag(self, tag):
        require(re.fullmatch(r'[A-Za-z0-9_][A-Za-z0-9_.-]{0,127}', tag), 'Invalid registry tag')
        for attempt in range(2):
            request = urllib.request.Request('https://ghcr.io/v2/' + self.repository + '/manifests/' + tag,
                                            headers={'Authorization': 'Bearer ' + self.token, 'Accept': INDEX})
            try:
                with self.opener.open(request, timeout=30) as response: raw = response.read(4 * 1024 * 1024 + 1)
                break
            except urllib.error.HTTPError as error:
                status = error.code; request_id = error.headers.get("X-GitHub-Request-Id") if error.headers else None
                same_resource = error.geturl() == request.full_url; error.close()
                if status == 401 and same_resource and attempt == 0:
                    self.refresh()
                    continue
                if status == 404: return None
                raise RegistryReadError(status, request_id) from None
        require(len(raw) <= 4 * 1024 * 1024, 'Oversized registry index')
        return sha_bytes(raw)

def promote(client, layout, inspected, repository, tag, actor, token, verification_dir, recorder=None):
    root = inspected['graph']['root']['digest']; remote = Ghcr(repository, actor, token)
    existing = remote.tag(tag)
    require(existing in (None, root), 'Refusing to replace a different version index')
    # Only the declared root is copied. Regctl may maintain its reviewed bounded subject metadata.
    def observed():
        actual = remote.tag(tag)
        if actual == root:
            return actual
        require(actual == existing, 'Published tag changed to an unexpected index')
        return None
    confirm('version-index-copy',
            lambda: tool_write([str(client), 'image', 'copy', '--force-recursive',
                                'ocidir://' + str(Path(layout).resolve()) + '@' + root,
                                'ghcr.io/' + repository + ':' + tag], capture_output=True),
            observed, recorder)
    output = Path(verification_dir)
    command([str(client), 'image', 'copy', '--force-recursive', 'ghcr.io/' + repository + '@' + root,
             'ocidir://' + str(output.resolve()) + ':verified'], capture_output=True)
    # The copy consumer creates locator/referrer metadata. Read only the already reviewed root closure.
    original_names = set(inspected['graph']['nodes'])
    for name in original_names:
        path = output/'blobs/sha256'/name[7:]
        require(path.is_file() and sha_file(path) == name
                and path.stat().st_size == inspected['graph']['nodes'][name]['size'], 'Published graph/blob differs')
    return root

def main():
    p = argparse.ArgumentParser(description=__doc__); sub = p.add_subparsers(dest='command', required=True)
    tool = sub.add_parser('install'); tool.add_argument('--pins', type=Path, required=True); tool.add_argument('--output', type=Path, required=True)
    args = p.parse_args()
    if args.command == 'install': install_client(args.pins, args.output)

if __name__ == '__main__':
    try: main()
    except (CandidateError, OSError, ValueError, KeyError, tarfile.TarError):
        raise SystemExit('OCI verification stopped; inspect retained state and do not rebuild as a fallback') from None
