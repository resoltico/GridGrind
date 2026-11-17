"""Pure candidate/source/OCI boundary controls; no GitHub, registry or product execution."""
import copy
import contextlib
import hashlib
import io
import json
import os
from pathlib import Path
import stat
import subprocess
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
    def test_declared_projection_still_rejects_missing_changed_extra_and_links(self):
        data=b'owned source';row={'mode':'100644','gitBlob':hashlib.sha1(b'blob '+str(len(data)).encode()+b'\0'+data).hexdigest(),
                                  'sha256':c.sha_bytes(data),'size':len(data),'exported':True}
        for kind in ('valid','missing','changed','extra','link'):
            with self.subTest(kind=kind),tempfile.TemporaryDirectory()as temp:
                path=Path(temp)/'sources.tar.gz'
                with tarfile.open(path,'w:gz')as t:
                    if kind!='missing':
                        body=b'changed data'if kind=='changed'else data
                        name='unexpected.java'if kind=='extra'else'App.java'
                        member=tarfile.TarInfo('gridgrind-0.78.0/'+name);member.size=len(body);member.mode=0o644
                        if kind=='link':member.type=tarfile.SYMTYPE;member.linkname='outside';member.size=0;t.addfile(member)
                        else:t.addfile(member,io.BytesIO(body))
                if kind=='valid':self.assertEqual(1,sources.verify_archive(path,'1'*40,'0.78.0',{'App.java':row})['files'])
                else:
                    with self.assertRaises(c.CandidateError):sources.verify_archive(path,'1'*40,'0.78.0',{'App.java':row})


class CommittedExportProjectionTest(unittest.TestCase):
    @contextlib.contextmanager
    def repository(self):
        with tempfile.TemporaryDirectory()as temporary:
            before=Path.cwd();root=Path(temporary)
            environment={k:v for k,v in os.environ.items()if not k.startswith('GIT_')}
            environment.update({'GIT_CONFIG_NOSYSTEM':'1','GIT_CONFIG_GLOBAL':'/dev/null',
                                'GIT_AUTHOR_NAME':'Synthetic export control','GIT_COMMITTER_NAME':'Synthetic export control',
                                'GIT_AUTHOR_EMAIL':'export@example.invalid','GIT_COMMITTER_EMAIL':'export@example.invalid'})
            def command(*args):
                result=subprocess.run(['git','-c','user.name=Synthetic export control','-c','user.email=export@example.invalid',
                                       '-c','commit.gpgsign=false','-c','core.hooksPath=/dev/null',
                                       '-c','core.attributesFile=/dev/null','-c','core.autocrlf=false','-c','core.eol=lf',*args],
                                      cwd=root,capture_output=True,check=True,env=environment)
                return result.stdout.decode().strip()
            command('init','-q')
            (root/'.gitattributes').write_text('* text=auto eol=lf\n*.bat eol=crlf\n/AGENTS.md export-ignore\n/.codex/ export-ignore\n/.codex/keep.txt -export-ignore\n')
            (root/'App.java').write_bytes(b'class App {}\n');(root/'gradlew.bat').write_bytes(b'@echo off\nrem windows\n')
            (root/'AGENTS.md').write_text('Synthetic nonshipping control metadata.\n')
            (root/'.codex').mkdir();(root/'.codex/keep.txt').write_text('Ancestor excludes this despite file unset.\n')
            command('add','.');command('commit','-qm','Synthetic export fixture');commit=command('rev-parse','HEAD')
            (root/'gradlew.bat').unlink()
            command('checkout-index','--all','--force')
            try:
                os.chdir(root)
                with mock.patch.dict(os.environ,environment,clear=True):yield root,commit,command
            finally:os.chdir(before)

    def test_actual_git_file_ancestor_policy_and_crlf_projection(self):
        with self.repository()as(root,commit,command):
            projection=sources.source_projection(commit)
            self.assertEqual(set(sources.tracked(commit)),set(projection))
            self.assertFalse(projection['AGENTS.md']['exported'])
            self.assertFalse(projection['.codex/keep.txt']['exported'])
            self.assertTrue(projection['App.java']['exported'])
            windows=b'@echo off\r\nrem windows\r\n'
            self.assertEqual(c.sha_bytes(windows),projection['gradlew.bat']['sha256'])
            self.assertEqual(len(windows),projection['gradlew.bat']['size'])
            self.assertEqual(windows,(root/'gradlew.bat').read_bytes())
            inputs=sources.working_input_hashes(commit,projection)
            self.assertIn('AGENTS.md',{r['path']for r in inputs})
            self.assertTrue(all(set(r)=={'path','size','sha256'}for r in inputs))
            asset=root/'gridgrind-0.78.0-sources.tar.gz';proof=sources.create_archive(asset,commit,'0.78.0',1)
            self.assertEqual(3,proof['files'])
            with tarfile.open(asset)as t:self.assertEqual(windows,t.extractfile('gridgrind-0.78.0/gradlew.bat').read())
            (root/'App.java').write_text('different source\n')
            with self.assertRaisesRegex(c.CandidateError,'input bytes'):sources.working_input_hashes(commit,projection)

    def test_committed_policy_overrides_worktree_global_and_attribute_source_env(self):
        with self.repository()as(root,commit,command):
            (root/'global-attrs').write_text('/App.java export-ignore\n')
            command('config','core.attributesFile',str(root/'global-attrs'))
            (root/'.gitattributes').write_text('* export-ignore\n')
            command('add','.gitattributes');command('commit','-qm','Different later export policy');later=command('rev-parse','HEAD')
            with mock.patch.dict(os.environ,{'GIT_ATTR_SOURCE':later}):
                projection=sources.source_projection(commit)
                self.assertTrue(projection['App.java']['exported'])
                asset=root/'gridgrind-0.78.0-sources.tar.gz'
                self.assertEqual(3,sources.create_archive(asset,commit,'0.78.0',1)['files'])

    def test_repository_info_overrides_and_custom_filters_refuse(self):
        with self.repository()as(root,commit,command):
            (root/'.git/info/attributes').write_text('/App.java export-ignore\n')
            with self.assertRaisesRegex(c.CandidateError,'info attributes'):sources.source_projection(commit)
        with self.repository()as(root,commit,command):
            with(root/'.gitattributes').open('a')as f:f.write('/App.java filter=unrecorded\n')
            command('add','.gitattributes');command('commit','-qm','Unsupported custom filter')
            with self.assertRaisesRegex(c.CandidateError,'custom filters'):sources.source_projection(command('rev-parse','HEAD'))

    def test_committed_policy_cannot_exclude_required_source_or_near_match_metadata(self):
        for name in ('core/src/main/java/Required.java','native-provider/src/helper/java/Helper.java',
                     'native-provider/patches/poi/Repair.patch','gradle/build-logic/build.gradle.kts',
                     'LICENSE','docs/API.md','AGENTS.other','.codex-other/required.txt','nested/AGENTS.md'):
            with self.subTest(path=name),self.repository()as(root,commit,command):
                path=root/name;path.parent.mkdir(parents=True,exist_ok=True);path.write_text('Required project material.\n')
                with(root/'.gitattributes').open('a')as f:f.write('/'+name+' export-ignore\n')
                command('add','.');command('commit','-qm','Overbroad export policy');selected=command('rev-parse','HEAD')
                projection=sources.source_projection(selected)
                self.assertFalse(projection[name]['exported'])
                self.assertIn(name,{r['path']for r in sources.working_input_hashes(selected,projection)})
                asset=root/'gridgrind-0.78.0-sources.tar.gz'
                with self.assertRaisesRegex(c.CandidateError,'required project source'):sources.create_archive(asset,selected,'0.78.0',1)
                self.assertFalse(asset.exists())

    def test_attribute_triplets_must_be_complete_unique_and_known(self):
        with mock.patch.object(sources,'export_environment',return_value={}),mock.patch.object(sources.subprocess,'run')as run:
            for raw in (b'App.java\0export-ignore\0set',b'App.java\0unknown\0set\0',b'App.java\0export-ignore\0set\0App.java\0export-ignore\0set\0'):
                run.return_value=mock.Mock(stdout=raw)
                with self.subTest(raw=raw),self.assertRaises(c.CandidateError):sources.export_attributes('1'*40,['App.java'])


class BeginSourcePreflightTest(unittest.TestCase):
    def test_begin_refuses_omitted_app_and_runs_all_input_projection_before_expensive_phase(self):
        class MissingTag(Exception):
            status=404
        api=mock.Mock();api.legacy.ApiError=MissingTag;api.legacy.Publication.return_value.release.return_value=None
        def get(path):
            if path=='/branches/main':return {'protected':True,'commit':{'sha':'1'*40}}
            if path=='/actions/workflows/release.yml':return {'id':2,'path':c.WORKFLOW_PATH}
            if path=='/git/ref/tags/v0.78.0':raise MissingTag()
            raise AssertionError('Unexpected API path')
        api.get.side_effect=get
        def git(*args):
            if args[0]=='rev-parse':return ('1'*40+'\n').encode()
            if args[0]=='status':return b''
            if args[0]=='show':return b'1\n'
            raise AssertionError('Unexpected Git source read')
        for exported in (False,True):
            with self.subTest(exported=exported),tempfile.TemporaryDirectory()as temp:
                env={'GITHUB_REPOSITORY':'example/GridGrind','GH_TOKEN':'synthetic-not-a-token','GITHUB_EVENT_NAME':'workflow_dispatch',
                     'GITHUB_REF':'refs/heads/main','RELEASE_WORKFLOW_COMMIT':'1'*40,'GITHUB_RUN_ID':'3','GITHUB_RUN_ATTEMPT':'4','GITHUB_OUTPUT':str(Path(temp)/'outputs')}
                projection={'core/src/main/java/App.java':{'exported':exported}}
                with mock.patch.dict(os.environ,env),mock.patch.object(delivery,'WORK',Path(temp)),                     mock.patch.object(delivery,'Github',return_value=api),mock.patch.object(delivery,'git',side_effect=git),                     mock.patch.object(delivery,'source_projection',return_value=projection),mock.patch.object(delivery,'working_input_hashes')as inputs:
                    if exported:
                        delivery.begin('0.78.0','1'*40);inputs.assert_called_once_with('1'*40,projection)
                    else:
                        with self.assertRaisesRegex(c.CandidateError,'required project source'):delivery.begin('0.78.0','1'*40)
                        inputs.assert_not_called();self.assertFalse((Path(temp)/'outputs').exists())


class PreparationDiagnosticsTest(unittest.TestCase):
    def test_known_guard_and_unknown_failure_keep_safe_context(self):
        with tempfile.TemporaryDirectory()as temp,mock.patch.object(delivery,'WORK',Path(temp)),mock.patch.object(delivery,'ACTIVE_PHASE','assemble'):
            message=delivery.preparation_failure(c.CandidateError('Owned source archive omitted exported committed material'))
            self.assertIn('omitted exported committed material',message)
            self.assertEqual('assemble',c.read_json(Path(temp)/'preparation-error.json')['phase'])
            try:raise RuntimeError('synthetic credential marker must not be logged')
            except RuntimeError as error:message=delivery.preparation_failure(error)
            record=c.read_json(Path(temp)/'preparation-error.json')
            self.assertEqual('RuntimeError',record['failureType']);self.assertTrue(record['frames'])
            self.assertNotIn('synthetic credential marker',message+json.dumps(record))
            with mock.patch.object(Path,'write_bytes',side_effect=OSError('private filesystem detail')):
                message=delivery.preparation_failure(c.CandidateError('Known fixed guard'))
            self.assertIn('Known fixed guard',message);self.assertIn('could not be retained',message)
            self.assertNotIn('private filesystem detail',message)


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


class InspectionReportTest(unittest.TestCase):
    def test_large_statements_remain_available_but_persisted_summary_stays_bounded(self):
        with tempfile.TemporaryDirectory()as temp:
            root=Path(temp);work=root/'work';work.mkdir()
            large='x'*(8*1024*1024+1)
            purl={'referenceType':'purl','referenceLocator':'pkg:maven/dev.erst.gridgrind/native-provider@0.78.0'}
            statements={platform:{'https://spdx.dev/Document':{'predicate':{'packages':[{'name':'native-provider','externalRefs':[purl]}]}},
                                  'https://slsa.dev/provenance/v1':{'predicate':{'large':large}}}
                        for platform in ('linux/amd64','linux/arm64')}
            full={'graph':{'root':{'digest':D},'nodes':{},'edges':[]},'platforms':{},'graphSha256':Q,'statements':statements}
            old=root/'old-full.json';c.write_json(old,full)
            with self.assertRaisesRegex(c.CandidateError,'oversized'):c.read_json(old)
            value={'version':'0.78.0','sourceCommit':'1'*40,'epoch':1,'runId':3,'runAttempt':4}
            with mock.patch.object(delivery,'WORK',work),mock.patch.object(delivery,'DELIVERY',root/'delivery'),\
                 mock.patch.object(delivery,'run_logged'),mock.patch.object(delivery,'unpack',return_value={'digest':D}),\
                 mock.patch.object(delivery,'inspect',return_value=full):
                _,_,returned,_=delivery.build_oci(value)
            compact=c.read_json(work/'runtime-inspection.json')
            self.assertEqual({'graph','platforms','graphSha256'},set(compact))
            self.assertIs(full,returned)
            oci.check_sbom(returned,[{'group':'dev.erst.gridgrind','artifact':'native-provider','version':'0.78.0'}])


class RuntimeImageReferenceTest(unittest.TestCase):
    def test_same_operational_index_is_used_for_both_platforms_not_selected_ids(self):
        from types import SimpleNamespace
        root=D; jar=c.sha_bytes(b'jar'); leaves={'linux/amd64':Q,'linux/arm64':'sha256:'+'4'*64}; operations=[]
        runtime={'graph':{'root':{'digest':root}},'platforms':{p:{'digest':h}for p,h in leaves.items()}}
        def command(args,**kwargs):
            operations.append(args)
            if args[:3]==['docker','image','inspect']:
                self.assertEqual(root,args[-1]);platform=args[args.index('--platform')+1]
                return SimpleNamespace(stdout=json.dumps([{'Id':leaves[platform],'Os':'linux','Architecture':platform.split('/')[1],
                                                          'Descriptor':{'digest':leaves[platform]}}]))
            self.assertEqual(['docker','run'],args[:2]);self.assertEqual(root,args[-2])
            return SimpleNamespace(stdout=jar[7:]+'  /app/gridgrind.jar\n')
        with mock.patch.object(delivery,'command',side_effect=command),mock.patch.object(delivery,'run_logged')as suites,\
             mock.patch.object(delivery,'xml_counts'),mock.patch.object(delivery.shutil,'copytree'),mock.patch.object(delivery,'sha_file',return_value=jar):
            bindings,receipts=delivery.verify_runtime_images({'version':'0.78.0'},runtime,jar)
        self.assertEqual({p:root for p in leaves},bindings)
        self.assertEqual(2,suites.call_count)
        for invocation,platform in zip(suites.call_args_list,leaves):
            self.assertIn('-PcontainerImage='+root,invocation.args[0]);self.assertIn('-PcontainerPlatform='+platform,invocation.args[0])
            self.assertEqual(root,receipts[platform]['operationalImageRoot'])
            self.assertEqual(leaves[platform],receipts[platform]['selectedInspectionId'])
            self.assertEqual(leaves[platform],receipts[platform]['platformManifest'])
        self.assertEqual(4,len(operations))

    def test_selected_platform_or_descriptor_mismatch_remains_a_refusal(self):
        from types import SimpleNamespace
        runtime={'graph':{'root':{'digest':D}},'platforms':{'linux/amd64':{'digest':Q}}}
        for bad in ({'Id':Q,'Os':'linux','Architecture':'arm64','Descriptor':{'digest':Q}},
                    {'Id':Q,'Os':'linux','Architecture':'amd64','Descriptor':{'digest':D}}):
            with self.subTest(selected=bad),mock.patch.object(delivery,'command',return_value=SimpleNamespace(stdout=json.dumps([bad]))):
                with self.assertRaisesRegex(c.CandidateError,'platform/config'):delivery.verify_runtime_images({'version':'0.78.0'},runtime,D)
