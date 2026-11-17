"""Pure candidate/source/OCI boundary controls; no GitHub, registry or product execution."""
import copy
import hashlib
import io
import json
import os
from pathlib import Path
import stat
import sys
import tarfile
import tempfile
import unittest
from unittest import mock
import zipfile

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import release_candidate as c
import release_delivery as delivery
import release_oci as oci
import release_sources as sources

IDENTITY = {'repository':'example/GridGrind','version':'0.78.0','sourceCommit':'1'*40,'epoch':1,
            'workflowId':2,'workflowPath':c.WORKFLOW_PATH,'workflowCommit':'1'*40,'runId':3,'runAttempt':4}
D = 'sha256:'+'2'*64
Q = 'sha256:'+'3'*64


def checks():
    names=('deterministic','active-fuzz','bounded-pit','runtime-amd64','runtime-arm64','devcontainer','delivery-content','oci-source-sbom')
    return {'formatVersion':1,'status':'PASS','suites':[{'id':n,'status':'PASS','tests':1,'failures':0,'errors':0,'skipped':0,'evidence':['evidence/pass.txt']}for n in names],
            'negativeControls':['evidence/negative.txt']}


def fixture(root):
    for name in ('evidence/pass.txt','evidence/negative.txt','oci/runtime.tar','oci/source.tar'):
        p=root/name;p.parent.mkdir(parents=True,exist_ok=True);p.write_text(name)
    rows=c.inventory(root)
    value={'formatVersion':1,'role':'delivery','identity':copy.deepcopy(IDENTITY),'inventory':rows,'inventorySha256':c.sha_bytes(c.canonical(rows)),
           'inputHashes':[{'path':'source.txt','size':1,'sha256':D}],
           'oci':{name:{'path':'oci/'+name+'.tar','root':D,'graphSha256':D}for name in ('runtime','source')},'checks':checks()}
    c.write_json(root/'candidate.json',value)
    return value


class CandidateContractTest(unittest.TestCase):
    def test_actual_payload_mutation_and_extra_file_refuse(self):
        with tempfile.TemporaryDirectory()as temp:
            root=Path(temp);value=fixture(root);self.assertEqual(value,c.candidate(root))
            (root/'evidence/pass.txt').write_text('changed')
            with self.assertRaisesRegex(c.CandidateError,'Payload'):c.candidate(root)
        with tempfile.TemporaryDirectory()as temp:
            root=Path(temp);fixture(root);(root/'extra').write_text('unlisted')
            with self.assertRaisesRegex(c.CandidateError,'Payload'):c.candidate(root)

    def test_payload_symlink_and_hardlink_are_not_file_aliases(self):
        for kind in ('symlink','hardlink'):
            with self.subTest(kind=kind),tempfile.TemporaryDirectory()as temp:
                root=Path(temp);fixture(root)
                if kind=='symlink':(root/'alias').symlink_to(root/'evidence/pass.txt')
                else:os.link(root/'evidence/pass.txt',root/'alias')
                with self.assertRaises(c.CandidateError):c.candidate(root)

    def test_identity_rejects_bool_ids_other_workflow_and_control_commit(self):
        for key,value in [('runId',True),('workflowPath','.github/workflows/ci.yml'),('workflowCommit','9'*40),('epoch',False)]:
            with self.subTest(key=key):
                bad=copy.deepcopy(IDENTITY);bad[key]=value
                with self.assertRaises(c.CandidateError):c.identity(bad)

    def test_checks_refuse_missing_empty_failed_skipped_or_future_results(self):
        for change in ('missing','empty','failed','skipped','boolean','future'):
            with self.subTest(change=change):
                value=checks()
                if change=='missing':value['suites'].pop()
                elif change=='empty':value['suites'][0]['tests']=0
                elif change=='failed':value['suites'][0]['failures']=1
                elif change=='skipped':value['suites'][0]['skipped']=1
                elif change=='boolean':value['suites'][0]['failures']=False
                else:value['futureArtifactId']=1
                with self.assertRaises(c.CandidateError):c.validate_checks(value)

    def test_service_role_cannot_mix_attempt_source_or_expiration(self):
        value={'id':7,'digest':D,'name':c.role_name('delivery',IDENTITY),'expired':False,'size_in_bytes':1,
               'workflow_run':{'id':3,'head_sha':'1'*40,'head_branch':'main'}}
        self.assertEqual({'id':7,'digest':D},c.service_artifact(value,value['name'],IDENTITY))
        for key,bad in [('name','delivery-0.78.0-3-2'),('expired',True),('digest',None)]:
            with self.subTest(key=key):
                row=copy.deepcopy(value);row[key]=bad
                with self.assertRaises(c.CandidateError):c.service_artifact(row,value['name'],IDENTITY)
        row=copy.deepcopy(value);row['workflow_run']['head_sha']='9'*40
        with self.assertRaises(c.CandidateError):c.service_artifact(row,value['name'],IDENTITY)

    def test_pair_excludes_later_acceptance_and_distinguishes_service_roles(self):
        pair=c.reviewed_pair(IDENTITY,{'id':7,'digest':D},D,{'id':8,'digest':Q},Q)
        before=c.sha_bytes(c.canonical(pair));self.assertNotIn('independent_acceptance',pair)
        self.assertEqual(before,c.sha_bytes(c.canonical(copy.deepcopy(pair))))
        with self.assertRaises(c.CandidateError):c.reviewed_pair(IDENTITY,{'id':7,'digest':D},D,{'id':7,'digest':Q},Q)

    def test_ack_binds_exact_pair_jar_status_and_evidence_digest(self):
        pair=c.sha_bytes(b'pair')
        ack={'formatVersion':1,'role':'independent-acceptance','pairSha256':pair,'jarSha256':D,
             'status':'PASS','evidenceSha256':D}
        c.independent_acceptance(ack,c.sha_bytes(c.canonical(ack)),pair,D)
        for key,value in [('pairSha256',Q),('jarSha256',Q),('status','FAIL'),('evidenceSha256','not-a-digest')]:
            bad=copy.deepcopy(ack);bad[key]=value
            with self.subTest(key=key),self.assertRaises(c.CandidateError):
                c.independent_acceptance(bad,c.sha_bytes(c.canonical(bad)),pair,D)


class ArtifactZipTest(unittest.TestCase):
    def test_namespace_rejected_before_destination_creation(self):
        for names in ([('../escape',b'x')],[('/absolute',b'x')],[('a',b'x'),('a',b'y')],[('a',b'x'),('a/b',b'y')]):
            with self.subTest(names=names),tempfile.TemporaryDirectory()as temp:
                root=Path(temp);path=root/'artifact.zip'
                with zipfile.ZipFile(path,'w')as z:
                    for name,data in names:z.writestr(name,data)
                with self.assertRaises(c.CandidateError):c.extract_zip(path,root/'out',c.sha_file(path))
                self.assertFalse((root/'out').exists())

    def test_symlink_mode_and_transport_digest_refuse(self):
        with tempfile.TemporaryDirectory()as temp:
            root=Path(temp);path=root/'artifact.zip';entry=zipfile.ZipInfo('link');entry.create_system=3;entry.external_attr=(stat.S_IFLNK|0o777)<<16
            with zipfile.ZipFile(path,'w')as z:z.writestr(entry,'outside')
            with self.assertRaises(c.CandidateError):c.extract_zip(path,root/'out',c.sha_file(path))
            with self.assertRaisesRegex(c.CandidateError,'digest'):c.extract_zip(path,root/'out',D)


class CapsuleSourceTest(unittest.TestCase):
    def test_helper_and_upstream_roles_use_their_distinct_declared_hash_fields(self):
        h=hashlib.sha256(b'helper').hexdigest();p=hashlib.sha256(b'patched').hexdigest()
        rows=[{'component':'helper','sourcePath':'src/helper/java/Helper.java','compiledSourcePath':'Helper.java','sha256':h},
              {'component':'poi','sourcePath':'Patched.java','modifiedSourceSha256':p}]
        class Archive:
            def read(self,name):return {'META-INF/gridgrind/native-provider/sources/modified/helper/Helper.java':b'helper',
                                        'META-INF/gridgrind/native-provider/sources/modified/poi/Patched.java':b'patched'}[name]
        self.assertEqual(2,delivery.verify_capsule_sources(Archive(),rows))
        bad=copy.deepcopy(rows);bad[0]['sha256']='0'*64
        with self.assertRaisesRegex(c.CandidateError,'source mismatch'):delivery.verify_capsule_sources(Archive(),bad)
        bad=copy.deepcopy(rows);bad[0]['component']='unknown'
        with self.assertRaisesRegex(c.CandidateError,'source role'):delivery.verify_capsule_sources(Archive(),bad)


class OwnedSourceTest(unittest.TestCase):
    def test_git_owned_archive_content_and_completeness_not_merely_tar_parse(self):
        data=b'owned source';oid=hashlib.sha1(b'blob '+str(len(data)).encode()+b'\0'+data).hexdigest()
        row=b'100644 blob '+oid.encode()+b'\tbuild.gradle.kts\0'
        with tempfile.TemporaryDirectory()as temp:
            path=Path(temp)/'sources.tar.gz'
            with tarfile.open(path,'w:gz')as t:
                member=tarfile.TarInfo('gridgrind-0.78.0/build.gradle.kts');member.size=len(data);member.mode=0o644;t.addfile(member,io.BytesIO(data))
            with mock.patch.object(sources,'git',return_value=row):self.assertEqual(1,sources.verify_archive(path,'1'*40,'0.78.0')['files'])
            with mock.patch.object(sources,'git',return_value=row+b'100644 blob '+oid.encode()+b'\tmissing.java\0'):
                with self.assertRaisesRegex(c.CandidateError,'omitted'):sources.verify_archive(path,'1'*40,'0.78.0')


class OciProfileTest(unittest.TestCase):
    def layout(self,root):
        blobs=root/'blobs/sha256';blobs.mkdir(parents=True)
        def put(value,media):
            raw=value if isinstance(value,bytes)else json.dumps(value,sort_keys=True,separators=(',',':')).encode()
            name='sha256:'+hashlib.sha256(raw).hexdigest();(blobs/name[7:]).write_bytes(raw)
            return {'mediaType':media,'digest':name,'size':len(raw)}
        layer=put(b'synthetic layer','application/vnd.oci.image.layer.v1.tar+gzip')
        members=[]
        for arch in ('amd64','arm64'):
            config=put({'os':'linux','architecture':arch,'config':{}},oci.CONFIG)
            runtime=put({'schemaVersion':2,'mediaType':oci.MANIFEST,'config':config,'layers':[layer]},oci.MANIFEST)
            members.append({**runtime,'platform':{'os':'linux','architecture':arch}})
            statements=[]
            for kind,predicate in [('https://spdx.dev/Document',{'packages':[{'name':'synthetic'}]}),('https://slsa.dev/provenance/v1',{'buildDefinition':{'buildType':'synthetic'}})]:
                statements.append(put({'_type':'https://in-toto.io/Statement/v1','subject':[{'digest':{'sha256':runtime['digest'][7:]}}],
                                       'predicateType':kind,'predicate':predicate},oci.ATTESTATION))
            unknown=put({'os':'unknown','architecture':'unknown'},oci.CONFIG)
            att=put({'schemaVersion':2,'mediaType':oci.MANIFEST,'config':unknown,'layers':statements,'subject':runtime},oci.MANIFEST)
            members.append({**att,'platform':{'os':'unknown','architecture':'unknown'},
                            'annotations':{'vnd.docker.reference.type':'attestation-manifest','vnd.docker.reference.digest':runtime['digest']}})
        return put({'schemaVersion':2,'mediaType':oci.INDEX,'manifests':members},oci.INDEX)

    def test_complete_indexed_subject_graph_rejects_missing_mutated_and_unselected_bytes(self):
        for kind in ('missing','mutated','extra'):
            with self.subTest(kind=kind),tempfile.TemporaryDirectory()as temp:
                root=Path(temp);desc=self.layout(root);good=oci.inspect(root,desc,'runtime')
                self.assertEqual({'linux/amd64','linux/arm64'},set(good['platforms']))
                victim=next((root/'blobs/sha256').iterdir())
                if kind=='missing':victim.unlink()
                elif kind=='mutated':victim.write_bytes(b'changed')
                else:(root/'blobs/sha256'/('0'*64)).write_bytes(b'unselected')
                with self.assertRaises((c.CandidateError,OSError)):oci.inspect(root,desc,'runtime')

    def test_actual_scanner_component_membership_is_not_nonempty_spdx(self):
        inspection={'statements':{'linux/amd64':{'https://spdx.dev/Document':{'predicate':{'packages':[{'name':'other'}]}}}}}
        required=[{'group':'dev.erst.gridgrind','artifact':'native-provider','version':'0.78.0'}]
        with self.assertRaisesRegex(c.CandidateError,'omitted'):oci.check_sbom(inspection,required)
        inspection['statements']['linux/amd64']['https://spdx.dev/Document']['predicate']['packages'][0]['externalRefs']=[
            {'referenceType':'purl','referenceLocator':'pkg:maven/dev.erst.gridgrind/native-provider@0.78.0'},
            {'referenceType':'purl','referenceLocator':'pkg:maven/org.junit.jupiter/junit-jupiter@6'}]
        with self.assertRaisesRegex(c.CandidateError,'tooling'):oci.check_sbom(inspection,required)


class PublisherBoundaryTest(unittest.TestCase):
    def test_invalid_independent_ack_precedes_any_mutation(self):
        import publish_prepared as pub
        from types import SimpleNamespace
        with tempfile.TemporaryDirectory()as temp:
            root=Path(temp);(root/'delivery/assets').mkdir(parents=True);(root/'delivery/assets/gridgrind.jar').write_bytes(b'jar')
            jar=c.sha_bytes(b'jar');pair=c.sha_bytes(b'pair')
            ack={'formatVersion':1,'role':'independent-acceptance','pairSha256':pair,'jarSha256':jar,'status':'FAIL','evidenceSha256':D}
            api=mock.Mock();api.get.return_value={'protected':True,'commit':{'sha':'1'*40}}
            args=SimpleNamespace(version='0.78.0',source='1'*40,run_id=3,attempt=4,delivery_id=7,delivery_digest=D,qa_id=8,qa_digest=Q,
                                 download=root,candidate_sha=D,receipt_sha=Q,pair_sha=pair,acceptance_sha=c.sha_bytes(c.canonical(ack)))
            env={'GITHUB_REPOSITORY':'example/GridGrind','GH_TOKEN':'synthetic-not-a-token','GITHUB_EVENT_NAME':'workflow_dispatch',
                 'GITHUB_REF':'refs/heads/main','RELEASE_WORKFLOW_COMMIT':'1'*40,'INDEPENDENT_ACCEPTANCE':json.dumps(ack)}
            record={'pair':{'candidateSha256':D,'receiptSha256':Q},'pairSha256':pair}
            with mock.patch.dict(os.environ,env),mock.patch.object(pub,'Github',return_value=api),\
                 mock.patch.object(pub,'expected_identity',return_value=IDENTITY),mock.patch.object(pub,'download_pair',return_value=(record,{})):
                with self.assertRaises(c.CandidateError):pub.publish(args)
            api.mutate.assert_not_called()
            self.assertFalse((root/'independent-acceptance.json').exists())

    def test_phase_receipt_retains_ambiguous_active_phase_without_secrets(self):
        import publish_prepared as pub
        with tempfile.TemporaryDirectory()as temp:
            before=Path.cwd()
            try:
                os.chdir(temp)
                roots={r:{'graph':{'root':{'digest':D}}}for r in ('runtime','source')}
                state=pub.PublicationProgress(Q,IDENTITY,D,roots)
                state.start('release-expose');state.complete();state.start('source-index-copy')
                actual=c.read_json('build/publication-receipt.json')
                self.assertEqual(['release-expose'],actual['completedPhases'])
                self.assertEqual('source-index-copy',actual['activePhase'])
                self.assertEqual(Q,actual['pairSha256'])
            finally:os.chdir(before)
