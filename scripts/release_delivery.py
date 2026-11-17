#!/usr/bin/env python3
"""Prepare and inspect immutable delivery files before any Actions candidate upload."""
import argparse
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import tarfile
import tempfile
import traceback
import xml.etree.ElementTree as ET
import zipfile
from release_candidate import (CandidateError, Github, canonical, candidate, identity, inventory, read_json,
                               require, relative, sha_bytes, sha_file, validate_checks, write_json)
from release_oci import (command, inspect, load_for_testing, properties, source_payload, unpack, check_sbom)
from release_sources import create_archive, git, tracked, source_projection, working_input_hashes, exported_source_rows

WORK = Path('build/release-work')
DELIVERY = Path('build/release-candidate')
PREFIX = 'META-INF/gridgrind/native-provider/'


def begin(version, source):
    repo = os.environ['GITHUB_REPOSITORY']; token = os.environ['GH_TOKEN']; api = Github(repo, token)
    require(os.environ['GITHUB_EVENT_NAME'] == 'workflow_dispatch' and os.environ['GITHUB_REF'] == 'refs/heads/main', 'Release controls run only by dispatch on main')
    workflow_commit = os.environ['RELEASE_WORKFLOW_COMMIT']
    require(source == workflow_commit == git('rev-parse', 'HEAD').decode().strip(), 'Checkout/source/workflow commit mismatch')
    require(not git('status', '--porcelain', '--untracked-files=all').strip(), 'Release checkout must start clean')
    branch = api.get('/branches/main')
    require(branch.get('protected') is True and branch.get('commit', {}).get('sha') == source, 'Exact protected main commit required')
    wf = api.get('/actions/workflows/release.yml')
    value = identity({'repository': repo, 'version': version, 'sourceCommit': source,
                      'epoch': int(git('show', '-s', '--format=%ct', source).decode().strip()),
                      'workflowId': wf['id'], 'workflowPath': wf['path'], 'workflowCommit': workflow_commit,
                      'runId': int(os.environ['GITHUB_RUN_ID']), 'runAttempt': int(os.environ['GITHUB_RUN_ATTEMPT'])})
    api.validate_run(value, False)
    legacy = api.legacy
    try: api.get('/git/ref/tags/v' + version)
    except legacy.ApiError as error:
        require(error.status == 404, 'Tag preflight failed')
    else: raise CandidateError('Preparation requires an untagged version; inspect/resume an existing reviewed publication')
    publication = legacy.Publication(api.api, repo, version, source, '', ())
    require(publication.release() is None, 'Existing release/draft must be inspected instead of preparing another candidate')
    WORK.mkdir(parents=True, exist_ok=True)
    write_json(WORK/'identity.json', value)
    # The exact property is checked again at seal. Refuse unsupported export
    # policy/representations before paying for the complete versioned graph.
    projection = source_projection(source)
    exported_source_rows(projection)
    working_input_hashes(source, projection)
    if os.environ.get('GITHUB_OUTPUT'):
        with open(os.environ['GITHUB_OUTPUT'], 'a', encoding='utf-8') as output:
            for key, item in {'version':version, 'commit':source, 'epoch':value['epoch'],
                              'image':'ghcr.io/'+repo.split('/')[0].lower()+'/gridgrind'}.items():
                output.write(str(key)+'='+str(item)+'\n')
    return value


def run_logged(args, name):
    path = WORK/name
    require(not path.exists(), 'Refusing to overwrite check evidence')
    with path.open('wb') as log:
        result = subprocess.run(args, stdout=log, stderr=subprocess.STDOUT, timeout=5400)
    require(result.returncode == 0, 'Artifact command failed; retain evidence and do not upload a candidate: ' + name)


def xml_counts(root):
    files = sorted(Path(root).glob('TEST-*.xml'))
    require(files, 'Required test lane produced no XML: ' + str(root))
    row = {'tests': 0, 'failures': 0, 'errors': 0, 'skipped': 0}
    for file in files:
        tree = ET.parse(file).getroot()
        for key in row: row[key] += int(tree.attrib.get(key, 0))
    require(row['tests'] > 0 and row['failures'] == row['errors'] == row['skipped'] == 0, 'Empty/failed/skipped portable test lane')
    return row


def copy_file(source, target):
    source, target = Path(source), Path(target)
    require(source.is_file() and not source.is_symlink() and not target.exists(), 'Invalid or occupied delivery file')
    target.parent.mkdir(parents=True, exist_ok=True); shutil.copyfile(source, target)


def legal_names(names):
    require(isinstance(names, list) and names and len(names) == len(set(names))
            and all(isinstance(n, str) and re.fullmatch(r'[A-Za-z0-9._-]+', n) for n in names), 'Invalid legal inventory')
    return names


def verify_capsule_sources(archive, rows):
    require(isinstance(rows, list) and rows, 'Empty capsule source provenance')
    seen = set()
    for row in rows:
        role = row['component']
        require(role in ('poi', 'poi-ooxml', 'helper'), 'Unknown capsule source role')
        name = row.get('compiledSourcePath', row['sourcePath'])
        require(isinstance(name, str) and not name.startswith('/') and '..' not in name.split('/'), 'Invalid capsule source path')
        path = PREFIX + 'sources/modified/' + role + '/' + name
        require(path not in seen, 'Duplicate capsule source provenance'); seen.add(path)
        expected = row['sha256'] if role == 'helper' else row['modifiedSourceSha256']
        require(isinstance(expected, str) and re.fullmatch(r'[0-9a-f]{64}', expected), 'Capsule source requires raw64 SHA256')
        require(hashlib.sha256(archive.read(path)).hexdigest() == expected, 'Capsule source mismatch')
        if 'patchPath' in row:
            require(role != 'helper', 'Helper must not impersonate a patched upstream source')
            patch = archive.read(PREFIX + 'sources/' + row['patchPath'])
            require(hashlib.sha256(patch).hexdigest() == row['patchSha256'], 'Capsule patch mismatch')
    return len(seen)


def inspect_container_sources(assets, version, manifest):
    """Independently hash delivered material, not just the collector's plan."""
    require(manifest.get('version') == version, 'Container source version differs')
    expected = {}
    for row in manifest['sources']:
        for name, facts in row['files'].items():
            path = relative('ubuntu/'+row['package']+'/'+row['version']+'/'+name)
            expected[path] = facts
    for row in manifest['gridgrindBuildSources']:
        relative(row['sourcePath']); relative(row['archivePath'])
        require(sha_file(row['sourcePath']) == 'sha256:'+row['sha256'], 'Container recipe/legal source differs from current commit')
        expected[row['archivePath']] = {'sha256':row['sha256'],'size':row['size']}
    expected[relative(manifest['vendorSourceArchive'])] = {'sha256':manifest['vendorSourceSha256']}
    expected[relative(manifest['vendorBuildArchive'])] = {'sha256':manifest['vendorBuildSha256']}
    for row in manifest['vendorMetadata']: expected[relative(row['path'])] = {'sha256':row['sha256']}
    external = assets/('gridgrind-container-sources-'+version+'.manifest.json')
    expected['manifest.json'] = {'sha256':sha_file(external)[7:],'size':external.stat().st_size}
    seen = set(); names = set(); prefix = 'gridgrind-container-sources-'+version+'/'
    archive = assets/('gridgrind-container-sources-'+version+'.tar.gz')
    with tarfile.open(archive) as tar:
        for entry in tar:
            require(entry.isdir() or entry.isfile(), 'Container source archive contains links/devices')
            relative(entry.name.rstrip('/'))
            require(entry.name not in names, 'Duplicate container source entry'); names.add(entry.name)
            if entry.isdir(): continue
            require(entry.name.startswith(prefix), 'Container source archive prefix differs')
            name = relative(entry.name[len(prefix):])
            if name not in expected: continue
            stream=tar.extractfile(entry); h=hashlib.sha256(); size=0
            for block in iter(lambda:stream.read(1024*1024),b''): size+=len(block); h.update(block)
            facts=expected[name]
            require(h.hexdigest()==facts['sha256'] and ('size'not in facts or size==facts['size']), 'Delivered container source payload hash/size mismatch')
            seen.add(name)
    require(seen==set(expected), 'Delivered container source material incomplete')
    expected_sums=sha_file(archive)[7:]+'  '+archive.name+'\n'+sha_file(external)[7:]+'  '+external.name+'\n'
    require((assets/('gridgrind-container-sources-'+version+'.sha256')).read_text()==expected_sums, 'Container source checksums differ')
    return {'verifiedPayloads':len(seen),'sourceTuples':len(manifest['sources'])}


def inspect_content(jar, components, version, assets, legal_files):
    legal = legal_names(legal_files); stock_count = 0
    with zipfile.ZipFile(jar) as z:
        names = z.namelist(); require(len(names) == len(set(names)), 'Duplicate JAR entries')
        for name in legal: require(z.read('META-INF/' + name) == Path(name).read_bytes(), 'JAR legal bytes differ')
        manifest = json.loads(z.read(PREFIX + 'manifest.json'))
        recipe = z.read(PREFIX + 'repairs.json')
        require(hashlib.sha256(recipe).hexdigest() == manifest['recipeManifestSha256'], 'Capsule recipe hash mismatch')
        require(manifest['component'] == {'group': 'dev.erst.gridgrind', 'artifact': 'native-provider', 'version': version}, 'Capsule own identity mismatch')
        classes = manifest['classes']; require(isinstance(classes, dict) and classes, 'Empty capsule repair list')
        required = z.read(PREFIX + 'required-classes.list').decode().splitlines()
        require(set(required) == {n.removesuffix('.class').replace('/', '.') for n in classes}, 'Required repair list mismatch')
        for name, expected in classes.items():
            require(name.startswith('org/apache/poi/') and '..' not in name.split('/'), 'Invalid repair path')
            data = z.read(PREFIX + 'classes/' + name)
            require(hashlib.sha256(data).hexdigest() == expected and int.from_bytes(data[6:8], 'big') == 70, 'Capsule class hash/target mismatch')
        actual_repairs = {n[len(PREFIX+'classes/'):] for n in names if n.startswith(PREFIX+'classes/') and n.endswith('.class')}
        require(actual_repairs == set(classes), 'Extra/missing repair bytes')
        verify_capsule_sources(z, manifest['sources'])
        for name in ('LICENSE-APACHE-2.0', 'LICENSE-AGPL-3.0-only', 'NOTICE'):
            require(z.read(PREFIX + name).strip(), 'Missing capsule legal material')
        props = z.read('META-INF/maven/dev.erst.gridgrind/native-provider/pom.properties').decode()
        require('version=' + version + '\n' in props and 'artifactId=native-provider\n' in props, 'Capsule publisher metadata missing')
        # Compare original stock classes with actual resolved publisher JARs, not source mirrors.
        for component in components:
            if component['group'] != 'org.apache.poi': continue
            raw = component['sha256']
            require(isinstance(raw, str) and re.fullmatch(r'[0-9a-f]{64}', raw), 'Runtime component requires raw64 SHA256')
            path = Path(component['file']); require(sha_file(path) == 'sha256:' + raw, 'Resolved vendor artifact changed')
            with zipfile.ZipFile(path) as upstream:
                for name in upstream.namelist():
                    if name.startswith('org/apache/poi/') and name.endswith('.class'):
                        require(z.read(name) == upstream.read(name), 'Parent stock provider class was modified'); stock_count += 1
        require(stock_count > 0, 'No original stock provider class comparison executed')
        project = [n for n in names if n.startswith('dev/erst/') and n.endswith('.class')]
        require(project and all(int.from_bytes(z.read(n)[6:8], 'big') == 70 for n in project), 'Project bytecode target differs')
        require(not any(n.startswith(('org/junit/', 'org/pitest/', 'com/tngtech/archunit/')) for n in names), 'Test tools leaked into delivery JAR')
    jar_sha = sha_file(jar)
    for suffix in ('zip', 'tar'):
        path = assets/('gridgrind-' + version + '.' + suffix); wanted = 'gridgrind-' + version + '/'
        if suffix == 'zip':
            with zipfile.ZipFile(path) as archive:
                members = archive.namelist(); jars = [n for n in members if n.startswith(wanted+'lib/') and n.endswith('.jar')]
                require(len(jars) == 1 and sha_bytes(archive.read(jars[0])) == jar_sha, 'Launcher ZIP JAR mismatch')
                for name in legal: require(archive.read(wanted+name) == Path(name).read_bytes(), 'Launcher ZIP legal mismatch')
        else:
            with tarfile.open(path) as archive:
                members = archive.getmembers(); jars = [m for m in members if m.isfile() and m.name.startswith(wanted+'lib/') and m.name.endswith('.jar')]
                require(len(jars) == 1, 'Launcher TAR JAR missing/duplicated')
                h = hashlib.sha256()
                stream = archive.extractfile(jars[0])
                for block in iter(lambda: stream.read(1024 * 1024), b''): h.update(block)
                require('sha256:' + h.hexdigest() == jar_sha, 'Launcher TAR JAR mismatch')
                for name in legal: require(archive.extractfile(wanted+name).read() == Path(name).read_bytes(), 'Launcher TAR legal mismatch')
    return {'stockProviderClassesCompared': stock_count, 'repairClasses': len(classes), 'sourceUnits': len(manifest['sources']), 'jarSha256': jar_sha}


def build_oci(value, source=False):
    role = 'source' if source else 'runtime'
    archive = DELIVERY/'oci'/(role+'.oci.tar'); archive.parent.mkdir(parents=True, exist_ok=True)
    tag = 'gridgrind-candidate:%s-%s-%s%s' % (value['version'], value['runId'], value['runAttempt'], '-sources' if source else '')
    args = ['docker', 'buildx', 'build', '--platform', 'linux/amd64,linux/arm64', '--tag', tag,
            '--build-arg', 'GRIDGRIND_VERSION='+value['version'], '--build-arg', 'GRIDGRIND_REVISION='+value['sourceCommit'],
            '--build-arg', 'SOURCE_DATE_EPOCH='+str(value['epoch']), '--output', 'type=oci,dest='+str(archive)]
    if source:
        args += ['--file', 'release/Dockerfile.sources', '--provenance=false', '--sbom=false', str(WORK/'source-image-payload')]
    else:
        args += ['--provenance=mode=max', '--attest', 'type=sbom,generator=docker.io/docker/buildkit-syft-scanner@sha256:873a7458c3c8ac8e32c7b796c10e1600b07b963d2aa000a72a2d97aaaa2d6315', '.']
    run_logged(args, role+'-oci-build.log')
    layout = WORK/(role+'-layout'); root = unpack(archive, layout); result = inspect(layout, root, role)
    # Keep full statements for immediate scanner/content checks in memory. Their
    # immutable raw bytes already live in the OCI; do not duplicate multi-MiB
    # predicates into the small graph summary consumed by seal.
    write_json(WORK/(role+'-inspection.json'), {key:result[key] for key in ('graph','platforms','graphSha256')})
    return archive, layout, result, tag


def verify_runtime_images(value, runtime, jar_sha):
    version = value['version']
    bindings = {}
    platform_receipts = {}
    retained_root = runtime['graph']['root']['digest']
    for arch in ('amd64', 'arm64'):
        platform = 'linux/'+arch; expected = runtime['platforms'][platform]['digest']
        result = command(['docker', 'image', 'inspect', '--platform', platform, retained_root], capture_output=True, text=True)
        selected = json.loads(result.stdout)[0]
        require(selected['Os']+'/'+selected['Architecture'] == platform and selected.get('Descriptor', {}).get('digest') == expected, 'Selected test platform/config does not bind reviewed manifest')
        output = command(['docker', 'run', '--rm', '--platform', platform, '--entrypoint', 'sha256sum', retained_root, '/app/gridgrind.jar'], capture_output=True, text=True).stdout
        require(output.split()[0] == jar_sha[7:], 'Image JAR differs from retained JAR')
        bindings[platform] = retained_root
        run_logged(['./gradlew', ':cli:containerTest', '-Pversion='+version, '-PcontainerImage='+retained_root, '-PcontainerPlatform='+platform], 'container-'+arch+'.log')
        xml_counts('cli/build/test-results/containerTest')
        shutil.copytree('cli/build/test-results/containerTest', WORK/('container-results-'+arch))
        platform_receipts[platform] = {'containerPlatformProperty': platform, 'retainedIndex': runtime['graph']['root']['digest'],
                                      'operationalImageRoot': retained_root, 'selectedInspectionId': selected['Id'], 'platformManifest': expected,
                                      'selectedConfigArchitecture': selected['Architecture'], 'jarSha256': jar_sha}
        require(sha_file('cli/build/libs/gridgrind.jar') == jar_sha, 'Runtime checks rebuilt or changed delivery JAR')
    return bindings, platform_receipts


def assemble(component_file, component_files):
    value = read_json(WORK/'identity.json'); version = value['version']
    require(not DELIVERY.exists(), 'Delivery staging must be absent')
    assets = DELIVERY/'assets'; assets.mkdir(parents=True)
    for source, name in [(Path('cli/build/libs/gridgrind.jar'), 'gridgrind.jar'),
                         (Path('cli/build/distributions')/('gridgrind-'+version+'.zip'), 'gridgrind-'+version+'.zip'),
                         (Path('cli/build/distributions')/('gridgrind-'+version+'.tar'), 'gridgrind-'+version+'.tar')]: copy_file(source, assets/name)
    owned = assets/('gridgrind-'+version+'-sources.tar.gz')
    owned_proof = create_archive(owned, value['sourceCommit'], version, value['epoch'])
    components = read_json(component_file)
    require(components.get('formatVersion') == 1 and components.get('components'), 'Resolved runtime inventory is empty')
    legal_files = legal_names(components.get('legalFiles'))
    local_files = read_json(component_files)
    require(local_files.get('formatVersion') == 1 and local_files.get('components'), 'Resolved vendor file inventory is empty')
    components = components['components']
    files = local_files['components']
    require([{k:v for k,v in row.items() if k != 'file'} for row in files] == components, 'Runtime file map differs from public component facts')
    content = inspect_content(assets/'gridgrind.jar', files, version, assets, legal_files)
    write_json(WORK/'delivery-content.json', {'status': 'PASS', **content, 'ownedSource': owned_proof})
    runtime_archive, runtime_layout, runtime, tag = build_oci(value)
    check_sbom(runtime, components)
    store_version = properties('release/transfer-tools.properties')['DOCKER_SERVER_VERSION']
    load_for_testing(runtime_archive, runtime, tag, store_version)
    jar_sha = sha_file(assets/'gridgrind.jar')
    bindings, platform_receipts = verify_runtime_images(value, runtime, jar_sha)
    run_logged(['./gradlew', ':cli:devcontainerTest', '-Pversion='+version], 'devcontainer.log')
    write_json(WORK/'runtime-test-bindings.json', platform_receipts)
    pins = properties('release/container-runtime.properties'); vendor = WORK/'vendor'; vendor.mkdir()
    import urllib.request
    vendor_inputs = []
    for role, url_key, hash_key in [('source', 'TEMURIN_SOURCE_URL', 'TEMURIN_SOURCE_SHA256'), ('build', 'TEMURIN_BUILD_URL', 'TEMURIN_BUILD_SHA256')]:
        output = vendor/(role+'.tar.gz')
        with urllib.request.urlopen(pins[url_key], timeout=60) as source_stream, output.open('xb') as dest:
            size = 0
            for block in iter(lambda: source_stream.read(1024 * 1024), b''):
                size += len(block); require(size < 2 * 1024**3, 'Vendor source exceeds acquisition bound'); dest.write(block)
        require(sha_file(output) == 'sha256:'+pins[hash_key], 'Full vendor source checksum mismatch'); vendor_inputs.append(output)
    source_dir = WORK/'container-sources'
    # Explicit platform selectors avoid Docker's host-architecture default for a shared retained index.
    run_logged(['python3', 'scripts/prepare-container-sources.py', '--image', 'linux/amd64='+bindings['linux/amd64'],
                '--image', 'linux/arm64='+bindings['linux/arm64'], '--version', version, '--output', str(source_dir),
                '--vendor-source', str(vendor_inputs[0]), '--vendor-build-source', str(vendor_inputs[1])], 'source-collection.log')
    for suffix in ('.tar.gz', '.manifest.json', '.sha256'):
        name = 'gridgrind-container-sources-'+version+suffix; copy_file(source_dir/name, assets/name)
    correspondence = read_json(assets/('gridgrind-container-sources-'+version+'.manifest.json'))
    require({r['platform']: r['manifestDigest'] for r in correspondence['images']} == {p: d['digest'] for p,d in runtime['platforms'].items()}, 'Source correspondence binds another platform manifest')
    source_content = inspect_container_sources(assets, version, correspondence)
    payload = WORK/'source-image-payload'; payload.mkdir()
    source_files = [owned] + [assets/('gridgrind-container-sources-'+version+s) for s in ('.tar.gz', '.manifest.json', '.sha256')]
    for path in source_files: copy_file(path, payload/path.name)
    source_archive, source_layout, source_image, _ = build_oci(value, True)
    source_payload(source_layout, source_image, source_files)
    write_json(WORK/'oci-source-sbom.json', {'status': 'PASS', 'runtimeGraphSha256': runtime['graphSha256'],
                                           'sourceGraphSha256': source_image['graphSha256'], 'componentCount': len(components),
                                           'deliveredSourceContent': source_content,
                                           'sourceManifestSha256': sha_file(assets/('gridgrind-container-sources-'+version+'.manifest.json'))})
    # Discriminating controls consume copies, never corrupt retained producer bytes.
    victim = next(name for name, row in runtime['graph']['nodes'].items() if row['mediaType'] not in {'application/vnd.oci.image.index.v1+json', 'application/vnd.oci.image.manifest.v1+json'})
    original = runtime_layout/'blobs/sha256'/victim[7:]
    backup = WORK/'negative-blob-backup'; shutil.copyfile(original, backup)
    negatives = []
    try:
        original.unlink()
        try: inspect(runtime_layout, runtime['graph']['root'], 'runtime')
        except (CandidateError, OSError): negatives.append('missing-blob-rejected')
        original.write_bytes(b'changed')
        try: inspect(runtime_layout, runtime['graph']['root'], 'runtime')
        except (CandidateError, OSError): negatives.append('mutated-blob-rejected')
    finally: shutil.copyfile(backup, original); backup.unlink()
    require(len(negatives) == 2, 'OCI validator negative control did not reject corruption')
    write_json(WORK/'oci-negative-controls.json', {'status': 'PASS', 'cases': negatives})
    inspect(runtime_layout, runtime['graph']['root'], 'runtime')
    sums = ''.join(sha_file(p)[7:]+'  '+p.name+'\n' for p in sorted(assets.iterdir()))
    (assets/'SHA256SUMS').write_text(sums)
    copy_file('build/release-notes.md', DELIVERY/'release-notes.md')
    # Preserve actual resolver facts while portable consumers use GAVs, not cache filenames.
    copy_file(component_file, DELIVERY/'evidence/runtime-components.json')
    return value


def seal():
    value = read_json(WORK/'identity.json')
    evidence = DELIVERY/'evidence'; evidence.mkdir(exist_ok=True)
    lanes = ['build/test-results/fuzzReplay', 'core/build/test-results/test', 'json/build/test-results/test',
             'engine/build/test-results/test', 'engine/build/test-results/parityTest', 'engine/build/test-results/publicNativeTest',
             'engine/build/test-results/publicNativeLifetimeTest', 'cli/build/test-results/test', 'cli/build/test-results/documentationTest',
             'architecture-tests/build/test-results/test', 'architecture-tests/build/test-results/architectureTest', 'gradle/build-logic/build/test-results/test']
    totals = {k: 0 for k in ('tests','failures','errors','skipped')}
    with tarfile.open(evidence/'deterministic-results.tar.gz', 'w:gz') as archive:
        for lane in lanes:
            counts = xml_counts(lane)
            for key in totals: totals[key] += counts[key]
            for file in sorted(Path(lane).glob('TEST-*.xml')): archive.add(file, arcname=file.as_posix())
    suites = [{'id': 'deterministic', 'status':'PASS', **totals, 'evidence':['evidence/deterministic-results.tar.gz', 'evidence/check.log']}]
    fuzz = (WORK/'active-fuzz.log').read_text()
    match = re.search(r'JQF completed ([0-9]+) trials; seeds executed=([0-9]+)/([0-9]+); product coverage probes=([0-9]+)', fuzz)
    require(match and int(match[1]) > 0 and match[2] == match[3] and int(match[4]) > 0, 'Missing actual active fuzz discovery/coverage')
    suites.append({'id':'active-fuzz','status':'PASS','tests':int(match[1]),'failures':0,'errors':0,'skipped':0,'evidence':['evidence/active-fuzz.log']})
    mutations = ET.parse('core/build/reports/pitest/mutations.xml').getroot().findall('mutation')
    require(mutations and all(m.get('status') == 'KILLED' for m in mutations), 'Bounded mutation smoke produced no killed controls')
    copy_file('core/build/reports/pitest/mutations.xml', evidence/'mutations.xml')
    suites.append({'id':'bounded-pit','status':'PASS','tests':len(mutations),'failures':0,'errors':0,'skipped':0,'evidence':['evidence/mutations.xml','evidence/bounded-pit.log']})
    for arch in ('amd64','arm64'):
        folder = WORK/('container-results-'+arch); counts = xml_counts(folder)
        with tarfile.open(evidence/('runtime-'+arch+'-results.tar.gz'), 'w:gz') as archive:
            for f in folder.glob('TEST-*.xml'): archive.add(f, arcname=f.name)
        suites.append({'id':'runtime-'+arch,'status':'PASS',**counts,'evidence':['evidence/runtime-'+arch+'-results.tar.gz','evidence/container-'+arch+'.log','evidence/runtime-test-bindings.json']})
    counts = xml_counts('cli/build/test-results/devcontainerTest')
    with tarfile.open(evidence/'devcontainer-results.tar.gz','w:gz') as archive:
        for f in Path('cli/build/test-results/devcontainerTest').glob('TEST-*.xml'): archive.add(f,arcname=f.name)
    suites.append({'id':'devcontainer','status':'PASS',**counts,'evidence':['evidence/devcontainer-results.tar.gz','evidence/devcontainer.log']})
    for name in ('delivery-content','oci-source-sbom'):
        report = read_json(WORK/(name+'.json')); require(report['status']=='PASS','Missing completed artifact QA')
        suites.append({'id':name,'status':'PASS','tests':1,'failures':0,'errors':0,'skipped':0,'evidence':['evidence/'+name+'.json']})
    for file in sorted(WORK.glob('*.log')): copy_file(file,evidence/file.name)
    for name in ('delivery-content.json','oci-source-sbom.json','oci-negative-controls.json','runtime-inspection.json','source-inspection.json','runtime-test-bindings.json'):
        copy_file(WORK/name,evidence/name)
    checks = validate_checks({'formatVersion':1,'status':'PASS','suites':suites,'negativeControls':['evidence/control-tests.log','evidence/oci-negative-controls.json']})
    control = (WORK/'control-tests.log').read_text()
    require(re.search(r'Ran [1-9][0-9]* tests',control) and re.search(r'\nOK\s*\Z',control), 'Release-control negative tests did not execute successfully')
    require(not git('status', '--porcelain', '--untracked-files=all').strip(), 'Release checks changed committed source inputs')
    projection = source_projection(value['sourceCommit'])
    owned_projection = read_json(WORK/'delivery-content.json')['ownedSource']['projection']
    require(projection == owned_projection, 'Source projection changed since owned archive creation')
    inputs = working_input_hashes(value['sourceCommit'], projection)
    rows = inventory(DELIVERY)
    oci = {}
    for role in ('runtime','source'):
        result = read_json(WORK/(role+'-inspection.json'))
        oci[role] = {'path':'oci/'+role+'.oci.tar','root':result['graph']['root']['digest'],'graphSha256':result['graphSha256']}
    write_json(DELIVERY/'candidate.json',{'formatVersion':1,'role':'delivery','identity':value,'inventory':rows,
               'inventorySha256':sha_bytes(canonical(rows)),'inputHashes':inputs,'oci':oci,'checks':checks})
    candidate(DELIVERY)


ACTIVE_PHASE = None


def preparation_failure(error):
    """Keep known safe guards, never exception payloads/locals/credentials."""
    known = isinstance(error, CandidateError)
    reason = str(error) if known else 'Unexpected failure; inspect retained stack locations'
    frames = [{'file':Path(frame.filename).name, 'line':frame.lineno, 'function':frame.name}
              for frame in traceback.extract_tb(error.__traceback__)[-16:]]
    value = {'formatVersion':1, 'status':'STOPPED', 'phase':ACTIVE_PHASE,
             'failureType':type(error).__name__, 'reason':reason, 'frames':frames}
    retained = True
    try:
        WORK.mkdir(parents=True, exist_ok=True)
        (WORK/'preparation-error.json').write_bytes(canonical(value)+b'\n')
    except OSError:
        retained = False
    suffix = '' if retained else ' Diagnostic file could not be retained.'
    return 'Release preparation stopped: '+value['failureType']+': '+reason+'. No candidate/publication is permitted.'+suffix


def main():
    p=argparse.ArgumentParser(description=__doc__);sub=p.add_subparsers(dest='command',required=True)
    b=sub.add_parser('begin');b.add_argument('--version',required=True);b.add_argument('--source',required=True)
    a=sub.add_parser('assemble');a.add_argument('--components',type=Path,required=True);a.add_argument('--component-files',type=Path,required=True)
    sub.add_parser('seal');args=p.parse_args()
    global ACTIVE_PHASE
    ACTIVE_PHASE = args.command
    if args.command=='begin': begin(args.version,args.source)
    elif args.command=='assemble': assemble(args.components,args.component_files)
    else: seal()

if __name__=='__main__':
    try: main()
    except Exception as error:
        raise SystemExit(preparation_failure(error)) from None
