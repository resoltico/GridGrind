#!/usr/bin/env python3
"""Publish only an exact independently reviewed immutable CI artifact pair.

No compiler/build invocation exists here. Stable releases are never replaced;
API mutations are not retried after ambiguous failures. Regctl's qualified
subject metadata may protect older graphs: five semver roots is not a storage bound.
"""
import argparse
import importlib.util
import json
import os
from pathlib import Path
import re
import subprocess
import sys
import urllib.error
import urllib.parse
import urllib.request
from release_candidate import (CandidateError, Github, canonical, candidate, digest, independent_acceptance,
                               read_json, require, sha_bytes, sha_file, write_json)
from release_pair import expected_identity, download_pair
from release_oci import (command, inspect, unpack, promote, check_sbom, source_payload, Ghcr)
from release_confirmation import confirm, diagnostic, redact, tool_write
from release_sources import verify_archive, tracked, source_projection, working_input_hashes


class PublicationProgress:
    """One bounded current-state receipt, including an ambiguous active phase."""
    def __init__(self, pair, ident, jar_sha, inspections):
        self.path=Path('build/publication-receipt.json')
        require(not self.path.exists(),'Publication state receipt already exists')
        self.value={'formatVersion':1,'status':'IN_PROGRESS','pairSha256':pair,'sourceCommit':ident['sourceCommit'],
                    'version':ident['version'],'jarSha256':jar_sha,
                    'expectedRoots':{k:v['graph']['root']['digest']for k,v in inspections.items()},
                    'activePhase':None,'completedPhases':[]}
        self.save()
    def save(self):
        self.path.parent.mkdir(parents=True,exist_ok=True)
        temporary=self.path.with_suffix('.json.partial')
        with temporary.open('wb')as output:output.write(canonical(self.value)+b'\n')
        temporary.replace(self.path)
    def observation(self, value):
        self.value.setdefault('writeObservations', []).append(value);self.save()
    def start(self,phase):
        self.value['activePhase']=phase;self.save()
    def complete(self):
        self.value['completedPhases'].append(self.value['activePhase']);self.value['activePhase']=None;self.save()
    def finish(self,status):
        self.value['status']=status;self.save()


def guard_main(api, ident):
    require(os.environ['GITHUB_EVENT_NAME']=='workflow_dispatch' and os.environ['GITHUB_REF']=='refs/heads/main'
            and os.environ['RELEASE_WORKFLOW_COMMIT']==ident['workflowCommit'], 'Publication controls must be the reviewed main workflow')
    branch=api.get('/branches/main')
    require(branch.get('protected') is True and branch.get('commit',{}).get('sha')==ident['sourceCommit'], 'Main moved or lost protection; reprepare instead of substituting controls')


def source_files(delivery,version):
    assets=delivery/'assets'
    return [assets/('gridgrind-'+version+'-sources.tar.gz')] + [assets/('gridgrind-container-sources-'+version+s) for s in ('.tar.gz','.manifest.json','.sha256')]


def assets_for(delivery,version):
    assets=delivery/'assets'
    names={'gridgrind.jar','gridgrind-'+version+'.zip','gridgrind-'+version+'.tar','SHA256SUMS'}|{p.name for p in source_files(delivery,version)}
    require({p.name for p in assets.iterdir()}==names, 'Release asset set differs from declared delivery contract')
    paths=sorted(assets.iterdir())
    expected=''.join(sha_file(p)[7:]+'  '+p.name+'\n' for p in paths if p.name!='SHA256SUMS')
    require((assets/'SHA256SUMS').read_text()==expected,'Release checksums do not describe the same retained asset bytes')
    return paths


def public_source_assets(repository,version,paths):
    # These anonymous GETs hash full response bytes; a HEAD or future tag archive is insufficient.
    from release_candidate import MAX_ARCHIVE
    class HttpsOnly(urllib.request.HTTPRedirectHandler):
        def redirect_request(self,request,fp,code,message,headers,url):
            require(urllib.parse.urlsplit(url).scheme=='https','Public source redirect must remain HTTPS')
            return super().redirect_request(request,fp,code,message,headers,url)
    opener=urllib.request.build_opener(HttpsOnly())
    import hashlib
    for path in paths:
        url='https://github.com/'+repository+'/releases/download/v'+version+'/'+path.name
        size=0;h=hashlib.sha256()
        with opener.open(url,timeout=60)as response:
            for block in iter(lambda:response.read(1024*1024),b''):
                size+=len(block);require(size<=MAX_ARCHIVE and size<=path.stat().st_size,'Oversized anonymous source asset');h.update(block)
        require(size==path.stat().st_size and 'sha256:'+h.hexdigest()==sha_file(path),'Anonymous source asset differs from reviewed bytes')


def verify_module():
    spec=importlib.util.spec_from_file_location('_gridgrind_public_verify',Path(__file__).with_name('verify-container-release.py'))
    module=importlib.util.module_from_spec(spec);spec.loader.exec_module(module);return module


def anonymous_graph(repository,expected,output):
    # Read every reviewed graph blob anonymously, including actual layer bytes.
    import hashlib
    module=verify_module();registry=module.Registry(repository);output=Path(output);require(not output.exists(),'Anonymous verification destination must be absent')
    (output/'blobs/sha256').mkdir(parents=True)
    for name,row in expected['graph']['nodes'].items():
        manifest=row['mediaType']in('application/vnd.oci.image.index.v1+json','application/vnd.oci.image.manifest.v1+json')
        path=output/'blobs/sha256'/name[7:];size=0;h=hashlib.sha256()
        with registry.open('manifests'if manifest else'blobs',name,timeout=60)as response,path.open('xb')as dest:
            for block in iter(lambda:response.read(1024*1024),b''):
                size+=len(block);require(size<=row['size'],'Anonymous source blob exceeds reviewed size');h.update(block);dest.write(block)
        require(size==row['size']and 'sha256:'+h.hexdigest()==name,'Anonymous source image graph differs')
    return output


def anonymous_source_image(repository,expected,output,files):
    source_payload(anonymous_graph(repository,expected,output),expected,files)


def ensure_tag(api,publication,ident,recorder=None):
    guard_main(api,ident)
    try:
        api.get('/git/ref/tags/v'+ident['version'])
    except api.legacy.ApiError as error:
        require(error.status==404,'Version tag preflight failed')
        def observed():
            try:
                publication.validate_tag()
            except api.legacy.ApiError as failure:
                if failure.status==404:
                    return None
                raise
            return ident['sourceCommit']
        confirm('version-tag-create',
                lambda: api.mutate('POST','/git/refs',{'ref':'refs/tags/v'+ident['version'],'sha':ident['sourceCommit']}),
                observed, recorder)
    publication.validate_tag()


def point_latest(client,repository,runtime_root,remote,public,recorder=None):
    prior_latest=remote.tag('latest')
    prior_public=public.tag_digest('latest')
    def latest_observed():
        actual=remote.tag('latest')
        anonymous=public.tag_digest('latest')
        require(actual in (prior_latest,runtime_root),'Latest index changed to an unexpected root')
        require(anonymous in (prior_public,runtime_root),'Public latest index changed to an unexpected root')
        return runtime_root if actual==anonymous==runtime_root else None
    confirm('latest-index-copy',lambda:tool_write([str(client),'image','copy','--force-recursive','ghcr.io/'+repository+'@'+runtime_root,'ghcr.io/'+repository+':latest'],capture_output=True),latest_observed,recorder)


def publish(args):
    repo=os.environ['GITHUB_REPOSITORY'];api=Github(repo,os.environ['GH_TOKEN'])
    ident=expected_identity(api,repo,args.version,args.source,args.run_id,args.attempt)
    guard_main(api,ident)
    record,manifest=download_pair(api,ident,args.delivery_id,args.delivery_digest,args.qa_id,args.qa_digest,args.download)
    require(record['pair']['candidateSha256']==digest(args.candidate_sha)
            and record['pair']['receiptSha256']==digest(args.receipt_sha)
            and record['pairSha256']==digest(args.pair_sha),'Requested reviewed pair does not match downloaded service roles')
    delivery=args.download/'delivery';jar_sha=sha_file(delivery/'assets/gridgrind.jar')
    raw=os.environ.get('INDEPENDENT_ACCEPTANCE','');require(len(raw.encode('utf-8'))<=32768,'Independent acknowledgement exceeds bound')
    def pairs(rows):
        result={}
        for key,value in rows:
            require(key not in result,'Duplicate independent acknowledgement field');result[key]=value
        return result
    ack=json.loads(raw,object_pairs_hook=pairs,parse_constant=lambda _:(_ for _ in()).throw(CandidateError('Nonfinite acknowledgement')))
    independent_acceptance(ack,args.acceptance_sha,record['pairSha256'],jar_sha)
    write_json(args.download/'independent-acceptance.json',ack)
    # Current controls and all declared commit inputs must still match, never artifact-supplied executable code.
    require({r['path']for r in manifest['inputHashes']}==set(tracked(args.source)), 'Input inventory omits or invents committed source/control files')
    projection=source_projection(args.source)
    require(working_input_hashes(args.source,projection)==manifest['inputHashes'], 'Current source/control projection differs from prepared inputs')
    files=assets_for(delivery,args.version)
    verify_archive(delivery/'assets'/('gridgrind-'+args.version+'-sources.tar.gz'),args.source,args.version,projection)
    inspections={};layouts={}
    for role in ('runtime','source'):
        selected=manifest['oci'][role];layout=args.download/(role+'-layout');root=unpack(delivery/selected['path'],layout);result=inspect(layout,root,role)
        require(result['graph']['root']['digest']==selected['root']and result['graphSha256']==selected['graphSha256'],'Reviewed OCI graph changed')
        layouts[role]=layout;inspections[role]=result
    components=read_json(delivery/'evidence/runtime-components.json')
    check_sbom(inspections['runtime'],components['components'])
    source_payload(layouts['source'],inspections['source'],source_files(delivery,args.version))
    notes=(delivery/'release-notes.md').read_text()
    legacy=api.legacy;publication=legacy.Publication(api.api,repo,args.version,args.source,notes,legacy.local_assets(files))
    existing=publication.release()
    repository=repo.split('/')[0].lower()+'/gridgrind'
    # These reads occur before ANY mutation, including Git tag creation.
    remote=Ghcr(repository,os.environ['GITHUB_ACTOR'],os.environ['GH_TOKEN'])
    for role,tag in [('source',args.version+'-sources'),('runtime',args.version)]:
        require(remote.tag(tag)in(None,inspections[role]['graph']['root']['digest']),'Different registry version graph already exists')
    progress=PublicationProgress(record['pairSha256'],ident,jar_sha,inspections)
    publication.recorder=progress.observation
    if existing is not None and not existing['draft']and not existing['prerelease']:
        progress.start('read-only-stable-postflight')
        publication.validate_tag();publication.verify(existing,'stable')
        require(remote.tag(args.version)==inspections['runtime']['graph']['root']['digest'],'Stable runtime index differs')
        require(remote.tag(args.version+'-sources')==inspections['source']['graph']['root']['digest'],'Stable source index differs')
        require(remote.tag('latest')==inspections['runtime']['graph']['root']['digest'],'Stable latest index differs')
        public_source_assets(repo,args.version,source_files(delivery,args.version))
        anonymous_source_image(repository,inspections['source'],args.download/'stable-source-anonymous',source_files(delivery,args.version))
        anonymous_graph(repository,inspections['runtime'],args.download/'stable-runtime-anonymous')
        require(api.get('/releases/latest').get('id')==existing['id'],'Stable GitHub latest differs')
        progress.complete();progress.finish('READ_ONLY_STABLE_POSTFLIGHT')
        return
    # Private config is workflow-owned; never write credentials into a user's default config.
    config=Path(os.environ['REGCTL_CONFIG']).resolve();temporary=Path(os.environ['RUNNER_TEMP']).resolve()
    require(config.is_relative_to(temporary)and not config.exists(),'Expected absent workflow-owned registry config')
    config.parent.mkdir(parents=True,exist_ok=True)
    login=subprocess.run([str(args.client),'registry','login','ghcr.io','--user',os.environ['GITHUB_ACTOR'],'--pass-stdin'],
                         input=os.environ['GH_TOKEN'].encode(),capture_output=True,timeout=60)
    require(login.returncode==0,'Registry login failed')
    guard_main(api,ident);progress.start('version-tag');ensure_tag(api,publication,ident,progress.observation);progress.complete()
    progress.start('release-draft');publication.run('prepare');progress.complete()
    guard_main(api,ident);progress.start('release-expose');publication.run('expose');progress.complete()
    progress.start('anonymous-source-assets')
    public_source_assets(repo,args.version,source_files(delivery,args.version))
    progress.complete()
    guard_main(api,ident)
    progress.start('source-index-copy')
    source_root=promote(args.client,layouts['source'],inspections['source'],repository,args.version+'-sources',os.environ['GITHUB_ACTOR'],os.environ['GH_TOKEN'],args.download/'published-source-authenticated',progress.observation)
    progress.complete();progress.start('anonymous-source-index')
    anonymous_source_image(repository,inspections['source'],args.download/'published-source-anonymous',source_files(delivery,args.version))
    progress.complete()
    guard_main(api,ident)
    progress.start('runtime-index-copy')
    runtime_root=promote(args.client,layouts['runtime'],inspections['runtime'],repository,args.version,os.environ['GITHUB_ACTOR'],os.environ['GH_TOKEN'],args.download/'published-runtime-authenticated',progress.observation)
    progress.complete();progress.start('public-runtime-verification')
    module=verify_module();public=module.Registry(repository)
    source_manifest=read_json(delivery/'assets'/('gridgrind-container-sources-'+args.version+'.manifest.json'))
    module.verify(public.fetch(runtime_root),public.fetch,source_manifest,public.fetch_blob)
    public.verify_tag(args.version,runtime_root)
    progress.complete()
    # Preserve the conservative existing policy, including protection of non-semver subject metadata.
    guard_main(api,ident)
    progress.start('five-semver-root-retention')
    command([sys.executable,str(Path(__file__).with_name('prune-container-versions.py')),'--owner',repo.split('/')[0],'--package','gridgrind','--keep','5','--apply'])
    progress.complete()
    guard_main(api,ident)
    progress.start('latest-index-pointer')
    point_latest(args.client,repository,runtime_root,remote,public,progress.observation)
    progress.complete();guard_main(api,ident);progress.start('stable-release-finalization');publication.run('finalize');progress.complete()
    progress.value.update({'independentAcceptanceSha256':args.acceptance_sha,'sourceIndex':source_root,'runtimeIndex':runtime_root,
               'retention':'Five semver version roots; protected subject/fallback/unassociated metadata can retain old graphs/storage indefinitely. No physical storage bound or concurrent external referrer merge guarantee.'})
    progress.finish('PUBLISHED_AND_VERIFIED')


def main():
    p=argparse.ArgumentParser(description=__doc__)
    for name in ('version','source','delivery-digest','candidate-sha','qa-digest','receipt-sha','pair-sha','acceptance-sha'):p.add_argument('--'+name,required=True)
    for name in ('run-id','attempt','delivery-id','qa-id'):p.add_argument('--'+name,type=int,required=True)
    p.add_argument('--download',type=Path,required=True);p.add_argument('--client',type=Path,required=True)
    publish(p.parse_args())

def stopped(error):
    details=diagnostic(error)
    if isinstance(error,CandidateError):
        details.update(category='refusal',message=redact(str(error)))
    path=Path('build/publication-receipt.json')
    try:
        if path.is_file():
            value=read_json(path);value['status']='STOPPED_INSPECT_REMOTE_STATE';value['failure']=details
            path.write_bytes(canonical(value)+b'\n')
    except Exception:
        # Failure evidence is best effort and must never replace the actual error.
        pass
    return 'Prepared publication stopped: '+json.dumps(details,sort_keys=True)+'. No automatic mutation retry; inspect remote state.'


if __name__=='__main__':
    try:main()
    except Exception as error:
        raise SystemExit(stopped(error))from None
