#!/usr/bin/env python3
"""Prepare owned-project sources from the exact Git commit before a tag exists."""
import gzip
import hashlib
from pathlib import Path
import re
import subprocess
import tarfile
from release_candidate import require, relative, sha_file

def git(*args):
    result = subprocess.run(['git', *args], capture_output=True, check=True)
    return result.stdout

def tracked(commit):
    require(re.fullmatch(r'[0-9a-f]{40}', commit), 'Owned sources require an immutable commit')
    rows = {}
    for record in git('ls-tree', '-rz', '--full-tree', commit).split(b'\0'):
        if not record: continue
        metadata, raw_name = record.split(b'\t', 1)
        mode, kind, oid = metadata.decode('ascii').split()
        name = relative(raw_name.decode('utf-8'))
        require(kind == 'blob' and mode in ('100644', '100755'), 'Owned source links/submodules are not supported silently')
        rows[name] = {'mode': mode, 'gitBlob': oid}
    require(rows, 'Git source tree is empty')
    return rows

def verify_archive(path, commit, version):
    expected = tracked(commit); found = set(); prefix = 'gridgrind-' + version + '/'
    with tarfile.open(path) as tar:
        for member in tar:
            require(member.isdir() or member.isfile(), 'Owned source links/devices forbidden')
            if member.isdir(): continue
            require(member.name.startswith(prefix), 'Unexpected owned archive prefix')
            name = relative(member.name[len(prefix):])
            require(name in expected and name not in found, 'Unexpected/duplicate owned source path')
            require(member.size <= 2 * 1024**3, 'Oversized owned source member')
            h = hashlib.sha1(b'blob ' + str(member.size).encode('ascii') + b'\0')
            stream = tar.extractfile(member)
            for block in iter(lambda: stream.read(1024 * 1024), b''): h.update(block)
            require(h.hexdigest() == expected[name]['gitBlob'], 'Owned source content differs from Git commit')
            require(bool(member.mode & 0o111) == (expected[name]['mode'] == '100755'), 'Owned source executable mode differs')
            found.add(name)
    require(found == set(expected), 'Owned source archive omitted committed material')
    return {'commit': commit, 'files': len(found), 'sha256': sha_file(path), 'name': Path(path).name}

def create_archive(output, commit, version, epoch):
    output = Path(output)
    require(output.name == 'gridgrind-' + version + '-sources.tar.gz' and not output.exists(), 'Expected absent canonical owned source asset')
    output.parent.mkdir(parents=True, exist_ok=True)
    # No working-tree walk, tag archive, export-ignore fallback, or later regeneration.
    process = subprocess.Popen(['git', 'archive', '--format=tar', '--prefix=gridgrind-' + version + '/', commit], stdout=subprocess.PIPE)
    with output.open('xb') as dest, gzip.GzipFile(filename='', mode='wb', fileobj=dest, mtime=epoch) as compressed:
        for block in iter(lambda: process.stdout.read(1024 * 1024), b''): compressed.write(block)
    process.stdout.close()
    require(process.wait() == 0 and output.stat().st_size < 2 * 1024**3, 'Owned source archive failed or exceeds asset limit')
    return verify_archive(output, commit, version)
