#!/usr/bin/env python3
"""Project commit-owned source bytes using committed Git export attributes."""
import gzip
import hashlib
import os
from pathlib import Path, PurePosixPath
import re
import subprocess
import tarfile
from release_candidate import require, relative, sha_file


def git(*args):
    result = subprocess.run(['git', '--no-replace-objects', *args], capture_output=True, check=True)
    return result.stdout


def tracked(commit):
    """All tracked regular blobs, including nonexported control metadata."""
    require(re.fullmatch(r'[0-9a-f]{40}', commit), 'Owned sources require an immutable commit')
    rows = {}
    for record in git('ls-tree', '-rz', '--full-tree', commit).split(b'\0'):
        if not record: continue
        metadata, raw_name = record.split(b'\t', 1)
        mode, kind, oid = metadata.decode('ascii').split()
        name = relative(raw_name.decode('utf-8'))
        require(kind == 'blob' and mode in ('100644', '100755'), 'Owned source links/submodules are not supported silently')
        require(re.fullmatch(r'[0-9a-f]{40}', oid), 'Invalid canonical Git blob identity')
        rows[name] = {'mode': mode, 'gitBlob': oid}
    require(rows, 'Git source tree is empty')
    return rows


def export_environment():
    """No host attribute overrides and no configuration mutations."""
    info = Path(os.fsdecode(git('rev-parse', '--path-format=absolute', '--git-path', 'info/attributes').rstrip(b'\n')))
    require(not info.is_symlink() and (not info.exists() or info.is_file() and info.stat().st_size == 0),
            'Owned source export refuses nonempty or indirect repository info attributes')
    env = dict(os.environ)
    env['GIT_ATTR_NOSYSTEM'] = '1'
    env.pop('GIT_ATTR_SOURCE', None)
    return env


def export_command(commit, *args):
    return ['git', '--no-replace-objects', '-c', 'core.attributesFile=/dev/null',
            '-c', 'core.autocrlf=false', '-c', 'core.eol=lf', '--attr-source='+commit, *args]


def export_attributes(commit, names):
    paths = set(names)
    for name in names:
        paths.update(parent.as_posix()+'/' for parent in PurePosixPath(name).parents if parent.as_posix() != '.')
    paths = sorted(paths)
    attributes = ('export-ignore', 'export-subst', 'filter', 'ident', 'working-tree-encoding', 'text', 'eol')
    data = ('\0'.join(paths)+'\0').encode('utf-8')
    result = subprocess.run(export_command(commit, 'check-attr', '--source='+commit, '--stdin', '-z', *attributes),
                            input=data, capture_output=True, check=True, env=export_environment())
    parts = result.stdout.split(b'\0')
    require(parts[-1] == b'' and (len(parts)-1) % 3 == 0, 'Malformed committed attribute response framing')
    rows = {name:{} for name in paths}
    for i in range(0, len(parts)-1, 3):
        name, attribute, value = (p.decode('utf-8') for p in parts[i:i+3])
        require(name in rows and attribute in attributes and attribute not in rows[name], 'Unexpected or duplicate committed attribute record')
        rows[name][attribute] = value
    require(all(set(row) == set(attributes) for row in rows.values()), 'Incomplete committed attribute response')
    return rows


def _raw_blobs(rows, commit):
    """One framed raw batch verifies canonical Git objects without per-file processes."""
    result = {}
    process = subprocess.Popen(export_command(commit, 'cat-file', '--batch'), env=export_environment(),
                               stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.DEVNULL)
    try:
        for name, row in rows.items():
            process.stdin.write((row['gitBlob']+'\n').encode('ascii')); process.stdin.flush()
            header = process.stdout.readline(256)
            fields = header.rstrip(b'\n').split(b' ')
            require(header.endswith(b'\n') and len(fields) == 3 and fields[0].decode('ascii') == row['gitBlob']
                    and fields[1] == b'blob' and fields[2].isdigit(), 'Malformed canonical Git blob response')
            size = int(fields[2]); require(size < 2 * 1024**3, 'Oversized owned source blob')
            oid = hashlib.sha1(b'blob '+str(size).encode('ascii')+b'\0'); sha = hashlib.sha256(); remaining = size; carriage = False
            while remaining:
                block = process.stdout.read(min(remaining, 1024*1024))
                require(block, 'Truncated canonical Git blob'); remaining -= len(block)
                oid.update(block); sha.update(block); carriage |= b'\r' in block
            require(process.stdout.read(1) == b'\n' and oid.hexdigest() == row['gitBlob'], 'Canonical Git blob content mismatch')
            result[name] = {'size':size, 'sha256':'sha256:'+sha.hexdigest(), 'hasCarriageReturn':carriage}
        process.stdin.close()
        require(process.stdout.read(1) == b'' and process.wait() == 0, 'Canonical Git batch did not finish cleanly')
    finally:
        if process.poll() is None: process.kill(); process.wait()
        if not process.stdin.closed: process.stdin.close()
        process.stdout.close()
    return result


def _filtered_blob(commit, name, oid):
    # Batch --filters headers report UNFILTERED size, so they cannot frame changed
    # payload lengths. Use native EOF-framed filtering only where EOL may change.
    process = subprocess.Popen(export_command(commit, 'cat-file', '--filters', '--path='+name, oid),
                               env=export_environment(), stdout=subprocess.PIPE, stderr=subprocess.DEVNULL)
    h = hashlib.sha256(); size = 0
    try:
        for block in iter(lambda:process.stdout.read(1024*1024), b''):
            size += len(block); require(size < 2 * 1024**3, 'Oversized projected source blob'); h.update(block)
        require(process.wait() == 0, 'Git declared source projection failed')
    finally:
        if process.poll() is None: process.kill(); process.wait()
        process.stdout.close()
    return {'size':size, 'sha256':'sha256:'+h.hexdigest()}


def source_projection(commit):
    """All input/control blobs plus separately derived archive membership."""
    rows = tracked(commit); attributes = export_attributes(commit, rows); raw = _raw_blobs(rows, commit); projected = {}
    for name, row in rows.items():
        attr = attributes[name]
        require(attr['filter'] in ('unset', 'unspecified') and attr['ident'] in ('unset', 'unspecified')
                and attr['working-tree-encoding'] in ('unset', 'unspecified') and attr['export-subst'] in ('unset', 'unspecified'),
                'Owned source projection does not support custom filters, ident, encoding or export-subst')
        require(attr['text'] in ('set', 'unset', 'unspecified', 'auto') and attr['eol'] in ('lf', 'crlf', 'unset', 'unspecified'),
                'Unsupported committed source text attributes')
        candidates = [name]+[parent.as_posix()+'/' for parent in PurePosixPath(name).parents if parent.as_posix() != '.']
        exported = all(attributes[p]['export-ignore'] != 'set' for p in candidates)
        body = raw[name]
        # Explicit binary text=unset has no EOL conversion. With controlled LF
        # defaults, raw bytes without CR already are the native LF projection.
        if attr['text'] != 'unset' and (attr['eol'] == 'crlf' or body['hasCarriageReturn']):
            body = _filtered_blob(commit, name, row['gitBlob'])
        projected[name] = {**row, 'size':body['size'], 'sha256':body['sha256'], 'exported':exported}
    require(any(row['exported'] for row in projected.values()), 'Committed source export selection is empty')
    return projected


def working_input_hashes(commit, projection=None):
    projection = source_projection(commit) if projection is None else projection
    result = []
    for name, row in sorted(projection.items()):
        path = Path(name)
        require(path.is_file() and not path.is_symlink() and path.stat().st_size == row['size'] and sha_file(path) == row['sha256'],
                'Source input bytes differ from committed Git projection')
        require(bool(path.stat().st_mode & 0o111) == (row['mode'] == '100755'), 'Source input mode differs from committed Git mode')
        result.append({'path':name, 'size':row['size'], 'sha256':row['sha256']})
    return result


def exported_source_rows(projection):
    # .gitattributes owns selection. This independent completeness floor permits
    # only its pre-existing root agent-control metadata scope to be omitted;
    # it must never bless ignored application/helper/repair/build/legal/docs.
    omitted = [name for name, row in projection.items() if not row['exported']]
    require(all(name == 'AGENTS.md' or name == '.codex' or name.startswith('.codex/') for name in omitted),
            'Committed export policy excludes required project source material')
    return {name:row for name,row in projection.items() if row['exported']}


def verify_archive(path, commit, version, projection=None):
    projection = source_projection(commit) if projection is None else projection
    expected = exported_source_rows(projection); found = set(); prefix = 'gridgrind-' + version + '/'
    with tarfile.open(path) as tar:
        for member in tar:
            require(member.isdir() or member.isfile(), 'Owned source links/devices forbidden')
            if member.isdir(): continue
            require(member.name.startswith(prefix), 'Unexpected owned archive prefix')
            name = relative(member.name[len(prefix):])
            require(name in expected and name not in found, 'Unexpected/duplicate owned source path')
            require(member.size == expected[name]['size'], 'Owned source projected size differs')
            h = hashlib.sha256(); stream = tar.extractfile(member)
            for block in iter(lambda: stream.read(1024 * 1024), b''): h.update(block)
            require('sha256:'+h.hexdigest() == expected[name]['sha256'], 'Owned source content differs from committed Git projection')
            require(bool(member.mode & 0o111) == (expected[name]['mode'] == '100755'), 'Owned source executable mode differs')
            found.add(name)
    require(found == set(expected), 'Owned source archive omitted exported committed material')
    return {'commit': commit, 'files': len(found), 'sha256': sha_file(path), 'name': Path(path).name, 'projection':projection}


def create_archive(output, commit, version, epoch):
    output = Path(output)
    require(output.name == 'gridgrind-' + version + '-sources.tar.gz' and not output.exists(), 'Expected absent canonical owned source asset')
    projection = source_projection(commit)
    exported_source_rows(projection)
    output.parent.mkdir(parents=True, exist_ok=True)
    process = subprocess.Popen(export_command(commit, 'archive', '--format=tar', '--prefix=gridgrind-' + version + '/', commit),
                               env=export_environment(), stdout=subprocess.PIPE, stderr=subprocess.DEVNULL)
    try:
        with output.open('xb') as dest, gzip.GzipFile(filename='', mode='wb', fileobj=dest, mtime=epoch) as compressed:
            for block in iter(lambda: process.stdout.read(1024 * 1024), b''): compressed.write(block)
        require(process.wait() == 0 and output.stat().st_size < 2 * 1024**3, 'Owned source archive failed or exceeds asset limit')
    finally:
        if process.poll() is None: process.kill(); process.wait()
        process.stdout.close()
    return verify_archive(output, commit, version, projection)
