#!/usr/bin/env python3
"""Download exact immutable artifact roles; bind transport, candidate and reviewed pair."""
import argparse
import datetime
import os
from pathlib import Path
from release_candidate import (CandidateError, Github, canonical, candidate, digest, identity, independent_acceptance,
                               read_json, require, reviewed_pair, role_name, service_artifact, sha_bytes, sha_file,
                               validate_receipt, extract_zip, write_json, COMMIT, VERSION, positive)


def expected_identity(api, repo, version, source, run, attempt):
    require(COMMIT.fullmatch(source) and VERSION.fullmatch(version), 'Invalid source/version before API reads')
    positive(run); positive(attempt)
    wf = api.get('/actions/workflows/release.yml')
    commit = api.get('/git/commits/' + source)
    require(commit.get('sha') == source, 'Git service returned another source object')
    date = commit.get('committer', {}).get('date')
    require(isinstance(date, str), 'Missing source commit epoch')
    epoch = int(datetime.datetime.fromisoformat(date.replace('Z', '+00:00')).timestamp())
    return identity({'repository':repo,'version':version,'sourceCommit':source,'epoch':epoch,
                     'workflowId':wf['id'],'workflowPath':wf['path'],'workflowCommit':source,
                     'runId':run,'runAttempt':attempt})


def download_role(api, ident, identifier, expected_digest, role, output):
    metadata = api.artifact(identifier)
    service = service_artifact(metadata, role_name(role,ident),ident)
    if expected_digest is not None:
        require(service['digest'] == digest(expected_digest), 'Reviewed artifact service digest changed')
    archive = Path(output).with_suffix('.zip')
    api.download(identifier,archive,service['digest'])
    extract_zip(archive,output,service['digest'])
    return service


def download_pair(api, ident, delivery_id, delivery_digest, qa_id, qa_digest, output):
    api.validate_run(ident, True)
    output=Path(output);require(not output.exists(),'Pair destination must be absent');output.mkdir(parents=True)
    delivery_service=download_role(api,ident,delivery_id,delivery_digest,'delivery',output/'delivery')
    delivery=candidate(output/'delivery');require(delivery['identity']==ident,'Candidate source/workflow/run identity mismatch')
    qa_service=download_role(api,ident,qa_id,qa_digest,'artifact-qa',output/'qa')
    receipt=validate_receipt(output/'qa',delivery,delivery_service)
    candidate_sha=sha_file(output/'delivery/candidate.json')
    require(receipt['candidateSha256']==candidate_sha,'QA receipt binds another candidate manifest')
    pair=reviewed_pair(ident,delivery_service,candidate_sha,qa_service,sha_file(output/'qa/artifact-qa.json'))
    record={'pair':pair,'pairSha256':sha_bytes(canonical(pair))}
    write_json(output/'reviewed-pair.json',record)
    return record,delivery


def create_receipt(repo,version,source,delivery_id,output,download):
    api=Github(repo,os.environ['GH_TOKEN'])
    require(os.environ['GITHUB_EVENT_NAME']=='workflow_dispatch'and os.environ['GITHUB_REF']=='refs/heads/main'
            and os.environ['RELEASE_WORKFLOW_COMMIT']==source,'QA must use current trusted main controls')
    ident=expected_identity(api,repo,version,source,int(os.environ['GITHUB_RUN_ID']),int(os.environ['GITHUB_RUN_ATTEMPT']))
    api.validate_run(ident,False)
    delivery_service=download_role(api,ident,delivery_id,None,'delivery',download)
    value=candidate(download);require(value['identity']==ident,'QA downloaded another source/attempt')
    output=Path(output);require(not output.exists(),'QA role must be absent');output.mkdir(parents=True)
    write_json(output/'artifact-qa.json',{'formatVersion':1,'role':'artifact-qa','identity':ident,'delivery':delivery_service,
               'candidateSha256':sha_file(Path(download)/'candidate.json'),'inventorySha256':value['inventorySha256'],
               'checks':value['checks'],'transportVerified':True})
    validate_receipt(output,value,delivery_service)


def main():
    p=argparse.ArgumentParser(description=__doc__);sub=p.add_subparsers(dest='command',required=True)
    q=sub.add_parser('receipt');q.add_argument('--repo',required=True);q.add_argument('--version',required=True);q.add_argument('--source',required=True)
    q.add_argument('--delivery-id',required=True,type=int);q.add_argument('--output',required=True,type=Path);q.add_argument('--download',required=True,type=Path)
    d=sub.add_parser('download');d.add_argument('--repo',required=True);d.add_argument('--version',required=True);d.add_argument('--source',required=True)
    d.add_argument('--run-id',required=True,type=int);d.add_argument('--attempt',required=True,type=int)
    d.add_argument('--delivery-id',required=True,type=int);d.add_argument('--delivery-digest',required=True)
    d.add_argument('--qa-id',required=True,type=int);d.add_argument('--qa-digest',required=True);d.add_argument('--output',required=True,type=Path)
    args=p.parse_args()
    if args.command=='receipt':create_receipt(args.repo,args.version,args.source,args.delivery_id,args.output,args.download)
    else:
        api=Github(args.repo,os.environ['GH_TOKEN']);ident=expected_identity(api,args.repo,args.version,args.source,args.run_id,args.attempt)
        record,_=download_pair(api,ident,args.delivery_id,args.delivery_digest,args.qa_id,args.qa_digest,args.output)
        print(record['pairSha256'])

if __name__=='__main__':
    try:main()
    except (CandidateError,OSError,ValueError,KeyError):
        raise SystemExit('Immutable artifact pair verification failed; no publication is permitted')from None
