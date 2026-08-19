import tempfile
import uuid
import unittest
from io import BytesIO
from pathlib import Path
import xlsxwriter
import app

class EngineTests(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory(); self.db=app.connect(Path(self.tmp.name)/'test.sqlite3'); app.init_db(self.db)
    def tearDown(self): self.db.close(); self.tmp.cleanup()
    def _mk_session(self,sid,name='test',campaign_id=None,group_code=None,expected=None,owner=None):
        t=self.db.execute('select id,version from templates').fetchone()
        self.db.execute("insert into sessions values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",(sid,t['id'],t['version'],name,'','','', 'open',app.now(),None,'',expected,owner,campaign_id,group_code,None,None,None))
        return t['id']
    def _mk_user(self,uid,email=None,role='pilote'):
        h,salt=app.hash_password('motdepasse123')
        self.db.execute('insert into users values(?,?,?,?,?,?,?)',(uid,email or f'{uid}@example.org',h,salt,role,uid,app.now()))
        return uid
    def _mk_campaign(self,cid,owner_user_id,name='Campagne'):
        t=self.db.execute('select id,version from templates').fetchone()
        self.db.execute('insert into campaigns values(?,?,?,?,?,?,?,?,?,?,?)',(cid,owner_user_id,name,'',None,None,t['id'],t['version'],'active',app.now(),app.now()))
        return t['id']
    def _indicator_ids(self,template_id):
        return [r['id'] for r in self.db.execute("select i.id from indicators i join domains d on d.id=i.domain_id where d.template_id=? order by d.display_order,i.display_order",(template_id,))]
    def _add_participant(self,sid,pid,template_id,value=None,status='completed',n=None):
        self.db.execute('insert into participants values(?,?,?,?,?,?,?)',(pid,sid,pid,status,app.now(),app.now() if status=='completed' else None,None))
        if value is not None:
            inds=self._indicator_ids(template_id)
            if n is not None: inds=inds[:n]
            for iid in inds:
                self.db.execute('insert into responses values(?,?,?,?,?,?,?,?)',(str(uuid.uuid4()),sid,pid,iid,str(value),'numeric',app.now(),app.now()))
    def test_epc_seed_has_seven_domains_and_seventy_indicators(self):
        t=self.db.execute('select id from templates').fetchone()['id']; payload=app.template_payload(self.db,t)
        self.assertEqual(len(payload['domains']),7); self.assertEqual(sum(len(d['indicators']) for d in payload['domains']),70)
    def test_grade_and_analysis_keep_raw_responses(self):
        t=self.db.execute('select id,version from templates').fetchone(); sid='session'; self.db.execute("insert into sessions values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",(sid,t['id'],t['version'],'test','','','', 'open',app.now(),None,'',None,None,None,None,None,None,None))
        domain=self.db.execute('select id from domains where display_order=1').fetchone()['id']; inds=self.db.execute('select id from indicators where domain_id=? order by display_order limit 1',(domain,)).fetchone()['id']
        for n,v in [('a',1),('b',5)]:
            pid=n; self.db.execute('insert into participants values(?,?,?,?,?,?,?)',(pid,sid,n,'completed',app.now(),app.now(),None)); self.db.execute('insert into responses values(?,?,?,?,?,?,?,?)',(str(uuid.uuid4()),sid,pid,inds,str(v),'numeric',app.now(),app.now()))
        self.db.commit(); out=app.analysis(self.db,sid); indicator=out['domains'][0]['indicators'][0]
        self.assertEqual(indicator['responses'],2); self.assertEqual(indicator['capacity'],60); self.assertEqual(indicator['consensus'],0); self.assertEqual(app.grade(63,app.GRADING),40)
    def test_reference_questionnaire_fix_never_touches_existing_version(self):
        t=self.db.execute('select id,version from templates').fetchone()
        sid='pinned-session'
        self.db.execute("insert into sessions values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",(sid,t['id'],t['version'],'test','','','', 'open',app.now(),None,'',None,None,None,None,None,None,None))
        domain=self.db.execute('select id from domains where template_id=? order by display_order limit 1',(t['id'],)).fetchone()['id']
        indicator=self.db.execute('select id from indicators where domain_id=? order by display_order limit 1',(domain,)).fetchone()['id']
        self.db.execute('insert into participants values(?,?,?,?,?,?,?)',('p',sid,'p','completed',app.now(),app.now(),None))
        self.db.execute('insert into responses values(?,?,?,?,?,?,?,?)',(str(uuid.uuid4()),sid,'p',indicator,'4','numeric',app.now(),app.now()))
        self.db.commit()
        before_domain_ids={r['id'] for r in self.db.execute('select id from domains where template_id=?',(t['id'],))}
        before_versions={r['version'] for r in self.db.execute("select version from templates where name='EPC / SENEVAL'")}
        # Simulate a stale referential on the latest version (as if EPC_DOMAINS had changed again).
        self.db.execute('update domains set code=? where id=?',('stale-code',domain))
        self.db.commit()
        app.ensure_reference_questionnaire_version(self.db)
        after_versions={r['version'] for r in self.db.execute("select version from templates where name='EPC / SENEVAL'")}
        self.assertGreater(len(after_versions),len(before_versions))
        after_domain_ids={r['id'] for r in self.db.execute('select id from domains where template_id=?',(t['id'],))}
        self.assertEqual(before_domain_ids,after_domain_ids)
        self.assertEqual(self.db.execute('select code from domains where id=?',(domain,)).fetchone()['code'],'stale-code')
        self.assertEqual(self.db.execute('select count(*) from responses where session_id=?',(sid,)).fetchone()[0],1)
        self.assertEqual(self.db.execute('select template_version from sessions where id=?',(sid,)).fetchone()['template_version'],t['version'])
    def test_qualitative_chain_is_persistent_and_exported(self):
        t=self.db.execute('select id,version from templates').fetchone(); sid='qualitative-session'
        self.db.execute("insert into sessions values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",(sid,t['id'],t['version'],'test','','','', 'open',app.now(),None,'',None,None,None,None,None,None,None))
        domain=self.db.execute('select id from domains where display_order=1').fetchone()['id']; indicator=self.db.execute('select id from indicators where domain_id=? limit 1',(domain,)).fetchone()['id']
        self.db.execute('insert into priorities values(?,?,?,?,?,?)',('priority',sid,domain,indicator,1,app.now()))
        self.db.execute('insert into priority_analyses values(?,?,?,?,?,?)',('analysis',sid,'priority','Constat',app.now(),app.now()))
        self.db.execute('insert into analysis_entries values(?,?,?,?,?,?,?,?,?,?,?)',('cause',sid,'priority',None,'cause','Cause','Cause','', 'RETENU',app.now(),app.now()))
        self.db.execute('insert into workshop_recommendations values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)',('recommendation',sid,'priority','cause',None,'Action','Description','Formation','Haute','Responsable','Demain','', 'Retenue',app.now(),app.now()))
        self.db.execute('insert into training_topics values(?,?,?,?,?,?,?,?,?,?,?)',('training',sid,'priority','recommendation','Thème','Besoin','Public','Haute','',app.now(),app.now())); self.db.commit()
        data=app.qualitative_data(self.db,sid)
        self.assertEqual(len(data['priorities']),1); self.assertEqual(len(data['entries']),1); self.assertEqual(data['recommendations'][0]['status'],'Retenue')
        self.assertGreater(len(app.report_xlsx(self.db,sid)),1000)

    # --- Régression : isolement des groupes homonymes entre campagnes (recette 2026-08-17) ---

    def test_homonym_groups_across_campaigns_stay_isolated(self):
        tA=self._mk_session('groupA',name='GROUPE IDENTIQUE',campaign_id='campA',group_code='GRO-01',expected=4)
        self._mk_session('groupB',name='GROUPE IDENTIQUE',campaign_id='campB',group_code='GRO-01',expected=8)
        for i,v in enumerate([5,5,4,4]): self._add_participant('groupA',f'a{i}',tA,value=v)
        for i,v in enumerate([1,1,2,2,3,3,1,2]): self._add_participant('groupB',f'b{i}',tA,value=v)
        self.db.commit()
        outA=app.analysis(self.db,'groupA'); outB=app.analysis(self.db,'groupB')
        self.assertEqual(outA['participantCount'],4); self.assertEqual(outA['completedCount'],4)
        self.assertEqual(outA['global']['capacity'],90)
        self.assertEqual(outB['participantCount'],8); self.assertEqual(outB['completedCount'],8)
        self.assertNotAlmostEqual(outB['global']['capacity'],outA['global']['capacity'])
        ind0=outA['domains'][0]['indicators'][0]
        self.assertEqual(ind0['distribution']['1'],0); self.assertEqual(ind0['distribution']['2'],0); self.assertEqual(ind0['distribution']['3'],0)

    def test_group_code_globally_unique_across_campaigns(self):
        self._mk_session('existing-group',name='Pikine',campaign_id='campX',group_code='PIK-01')
        self.db.commit()
        code=app.generate_group_code(self.db,'Pikine')
        self.assertEqual(code,'PIK-02')

    def test_capacity_excludes_incomplete_participants(self):
        t=self._mk_session('partial-session')
        self._add_participant('partial-session','p1',t,value=5)
        self._add_participant('partial-session','p2',t,value=5)
        self._add_participant('partial-session','p3',t,value=4)
        self._add_participant('partial-session','p4',t,value=4)
        self._add_participant('partial-session','abandoned1',t,value=None,status='in_progress')
        self._add_participant('partial-session','abandoned2',t,value=5,status='in_progress',n=2)
        # Scénario D (consignes finalisation V2) : un 3e incomplet à valeur extrême (1, sur
        # toutes les questions) ne doit toujours strictement rien changer au résultat.
        self._add_participant('partial-session','abandoned3',t,value=1,status='in_progress')
        self.db.commit()
        out=app.analysis(self.db,'partial-session')
        self.assertEqual(out['participantCount'],7)
        self.assertEqual(out['completedCount'],4)
        self.assertEqual(out['global']['capacity'],90.0)
        self.assertAlmostEqual(out['global']['consensus'],71.13248654051871,places=9)
        self.assertEqual(out['global']['gradedCapacity'],85)
        self.assertEqual(out['global']['gradedConsensus'],50)

    def test_recette_scenario_2_consolidated(self):
        t=self._mk_session('scenario2')
        for i,v in enumerate([5]*5+[1]*5): self._add_participant('scenario2',f'p{i}',t,value=v)
        self.db.commit()
        out=app.analysis(self.db,'scenario2')
        self.assertEqual(out['global']['capacity'],60)
        self.assertEqual(out['global']['consensus'],0)
        self.assertEqual(out['global']['gradedCapacity'],40)
        self.assertEqual(out['global']['gradedConsensus'],5)

    def test_recette_scenario_3_groups_and_consolidations(self):
        t=self._mk_session('gr1'); self._mk_session('gr2'); self._mk_session('gr3')
        for i,v in enumerate([5,5,4,4]): self._add_participant('gr1',f'g1p{i}',t,value=v)
        for i,v in enumerate([3,3,3,3]): self._add_participant('gr2',f'g2p{i}',t,value=v)
        for i,v in enumerate([1,2,1,2]): self._add_participant('gr3',f'g3p{i}',t,value=v)
        self.db.commit()
        gr1=app.analysis(self.db,'gr1'); gr2=app.analysis(self.db,'gr2'); gr3=app.analysis(self.db,'gr3')
        self.assertEqual(gr1['global']['capacity'],90); self.assertAlmostEqual(gr1['global']['consensus'],71.1,places=1)
        self.assertEqual(gr2['global']['capacity'],60); self.assertEqual(gr2['global']['consensus'],100)
        self.assertEqual(gr3['global']['capacity'],30); self.assertAlmostEqual(gr3['global']['consensus'],71.1,places=1)
        allg=app.analysis_for(self.db,['gr1','gr2','gr3'])
        self.assertEqual(allg['global']['capacity'],60); self.assertAlmostEqual(allg['global']['consensus'],32.6,places=1)
        gr1_gr3=app.analysis_for(self.db,['gr1','gr3'])
        self.assertEqual(gr1_gr3['global']['capacity'],60); self.assertAlmostEqual(gr1_gr3['global']['consensus'],15.5,places=1)
        gr2_gr3=app.analysis_for(self.db,['gr2','gr3'])
        self.assertEqual(gr2_gr3['global']['capacity'],45); self.assertAlmostEqual(gr2_gr3['global']['consensus'],55.7,places=1)

    def test_massive_isolation_five_campaigns_five_homonym_groups(self):
        t=None; expected={}
        for c in range(5):
            for g in range(5):
                name=f'Groupe {chr(65+g)}'; sid=f'c{c}g{g}'; value=(c*5+g)%5+1
                tid=self._mk_session(sid,name=name,campaign_id=f'camp{c}',group_code=f'COD-{c}{g}')
                t=tid
                for p in range(2): self._add_participant(sid,f'{sid}p{p}',tid,value=value)
                expected[sid]=value
        self.db.commit()
        for sid,value in expected.items():
            out=app.analysis(self.db,sid)
            self.assertEqual(out['participantCount'],2,sid); self.assertEqual(out['completedCount'],2,sid)
            self.assertEqual(out['global']['capacity'],value/5*100,sid)
    # --- Régression : finalisation V2 (objectif, suppression autonome, cloisonnement) ---

    def test_objectif_is_never_a_hard_cap(self):
        t=self._mk_session('grp-cap',expected=20)
        for i in range(25): self._add_participant('grp-cap',f'p{i}',t,value=4)
        self.db.commit()
        out=app.analysis(self.db,'grp-cap')
        self.assertEqual(out['participantCount'],25); self.assertEqual(out['completedCount'],25)
        self.assertEqual(out['global']['capacity'],80.0)

    def test_objectif_30_and_overshoot(self):
        t=self._mk_session('grp-cap30',expected=30)
        for i in range(35): self._add_participant('grp-cap30',f'p{i}',t,value=5)
        self.db.commit()
        out=app.analysis(self.db,'grp-cap30')
        self.assertEqual(out['participantCount'],35); self.assertEqual(out['completedCount'],35)

    def test_group_without_objectif_has_null_expected_participants(self):
        self._mk_session('grp-noobj',expected=None)
        self.db.commit()
        row=self.db.execute('select expected_participants from sessions where id=?',('grp-noobj',)).fetchone()
        self.assertIsNone(row['expected_participants'])

    def test_consolidation_a_b_pools_completed_only(self):
        uid=self._mk_user('u1')
        self._mk_campaign('cmp-ab',uid)
        tA=self._mk_session('grp-a',campaign_id='cmp-ab',group_code='GA-01',expected=20)
        self._mk_session('grp-b',campaign_id='cmp-ab',group_code='GB-01',expected=30)
        for i in range(4): self._add_participant('grp-a',f'a{i}',tA,value=5)
        for i in range(3): self._add_participant('grp-a',f'ax{i}',tA,value=1,status='in_progress')
        for i in range(4): self._add_participant('grp-b',f'b{i}',tA,value=3)
        for i in range(3): self._add_participant('grp-b',f'bx{i}',tA,value=5,status='in_progress')
        self.db.commit()
        out=app.analysis_for(self.db,['grp-a','grp-b'])
        self.assertEqual(out['participantCount'],14); self.assertEqual(out['completedCount'],8)
        self.assertEqual(out['global']['capacity'],80.0)
        self.assertAlmostEqual(out['global']['consensus'],46.5,places=1)
        self.assertEqual(out['global']['gradedCapacity'],65); self.assertEqual(out['global']['gradedConsensus'],25)

    def test_campaign_delete_blocked_without_force(self):
        uid=self._mk_user('u2')
        self._mk_campaign('cmp-block',uid)
        t=self._mk_session('grp-block',campaign_id='cmp-block',expected=5)
        self._add_participant('grp-block','p0',t,value=4)
        self.db.commit()
        deleted,used=app.delete_campaign_cascade(self.db,'cmp-block',force=False)
        self.assertFalse(deleted); self.assertEqual(used,70)  # 1 participant x 70 indicateurs répondus
        self.assertIsNotNone(self.db.execute('select id from campaigns where id=?',('cmp-block',)).fetchone())

    def test_campaign_delete_forced_removes_all_campaign_rows(self):
        uid=self._mk_user('u3')
        self._mk_campaign('cmp-force',uid)
        t=self._mk_session('grp-force',campaign_id='cmp-force',expected=5)
        self._add_participant('grp-force','p0',t,value=4)
        self._add_participant('grp-force','p1',t,value=4)
        self.db.commit()
        deleted,used=app.delete_campaign_cascade(self.db,'cmp-force',force=True)
        self.assertTrue(deleted); self.assertEqual(used,140)  # 2 participants x 70 indicateurs répondus
        self.assertIsNone(self.db.execute('select id from campaigns where id=?',('cmp-force',)).fetchone())
        self.assertIsNone(self.db.execute('select id from sessions where id=?',('grp-force',)).fetchone())
        self.assertEqual(self.db.execute('select count(*) from participants where session_id=?',('grp-force',)).fetchone()[0],0)
        self.assertEqual(self.db.execute('select count(*) from responses where session_id=?',('grp-force',)).fetchone()[0],0)

    def test_campaign_delete_never_touches_shared_questionnaire(self):
        uid=self._mk_user('u4')
        self._mk_campaign('cmp-shared',uid)
        t=self._mk_session('grp-shared',campaign_id='cmp-shared',expected=5)
        self._add_participant('grp-shared','p0',t,value=4)
        self.db.commit()
        before=(self.db.execute('select count(*) from templates').fetchone()[0],
                self.db.execute('select count(*) from domains').fetchone()[0],
                self.db.execute('select count(*) from indicators').fetchone()[0])
        app.delete_campaign_cascade(self.db,'cmp-shared',force=True)
        after=(self.db.execute('select count(*) from templates').fetchone()[0],
               self.db.execute('select count(*) from domains').fetchone()[0],
               self.db.execute('select count(*) from indicators').fetchone()[0])
        self.assertEqual(before,after)

    def test_campaign_delete_leaves_other_pilot_data_intact(self):
        uidA=self._mk_user('pilote-a'); uidB=self._mk_user('pilote-b')
        self._mk_campaign('cmp-a',uidA,name='TEST-CLAUDE-DELETE-A')
        self._mk_campaign('cmp-b',uidB,name='TEST-CLAUDE-DELETE-B')
        tA=self._mk_session('grp-a2',campaign_id='cmp-a',expected=5)
        tB=self._mk_session('grp-b2',campaign_id='cmp-b',expected=5,owner=uidB)
        self._mk_session('standalone-b',owner=uidB)
        self._add_participant('grp-a2','pa0',tA,value=4)
        self._add_participant('grp-b2','pb0',tB,value=3)
        self.db.commit()
        deleted,_=app.delete_campaign_cascade(self.db,'cmp-a',force=True)
        self.assertTrue(deleted)
        self.assertIsNotNone(self.db.execute('select id from campaigns where id=?',('cmp-b',)).fetchone())
        self.assertIsNotNone(self.db.execute('select id from sessions where id=?',('grp-b2',)).fetchone())
        self.assertEqual(self.db.execute('select count(*) from participants where session_id=?',('grp-b2',)).fetchone()[0],1)
        self.assertIsNotNone(self.db.execute('select id from sessions where id=?',('standalone-b',)).fetchone())

    def test_cross_pilot_campaign_access_denied(self):
        uidA=self._mk_user('own-a'); uidB=self._mk_user('own-b')
        self._mk_campaign('cmp-owned-a',uidA)
        self.db.commit()
        userA={'id':uidA,'role':'pilote'}; userB={'id':uidB,'role':'pilote'}; admin={'id':'adm','role':'admin'}
        for path in ('/api/campaigns/cmp-owned-a','/api/campaigns/cmp-owned-a/groups','/api/campaigns/cmp-owned-a/consolidate'):
            with self.assertRaises(app.PermissionDeniedError):
                app.Handler.check_ownership(None,path,self.db,userB)
            app.Handler.check_ownership(None,path,self.db,userA)
            app.Handler.check_ownership(None,path,self.db,admin)

    def test_group_delete_force(self):
        uid=self._mk_user('u5')
        self._mk_campaign('cmp-grp',uid)
        t=self._mk_session('grp-del',campaign_id='cmp-grp',expected=5)
        self._add_participant('grp-del','p0',t,value=4)
        self.db.commit()
        deleted,used=app.delete_group_cascade(self.db,'cmp-grp','grp-del',force=False)
        self.assertFalse(deleted); self.assertEqual(used,70)  # 1 participant x 70 indicateurs répondus
        deleted,used=app.delete_group_cascade(self.db,'cmp-grp','grp-del',force=True)
        self.assertTrue(deleted); self.assertEqual(used,70)
        self.assertIsNone(self.db.execute('select id from sessions where id=?',('grp-del',)).fetchone())

    # --- Lot 1b (modularisation) : resolution d'utilisateur et d'ownership extraites vers epc/auth.py ---

    def test_resolve_current_user_from_cookie_header(self):
        uid=self._mk_user('u-auth',role='pilote')
        self.db.commit()
        token=app.create_auth_token(self.db,uid)
        user=app.resolve_current_user(self.db,f'epc_session={token}')
        self.assertIsNotNone(user); self.assertEqual(user['id'],uid)
        self.assertIsNone(app.resolve_current_user(self.db,None))
        self.assertIsNone(app.resolve_current_user(self.db,'epc_session=not-a-real-token'))

    def test_resolve_auth_public_route_without_cookie_returns_none(self):
        self.assertIsNone(app.resolve_auth('/api/auth/setup-status','GET',self.db,None))

    def test_resolve_auth_private_route_requires_cookie(self):
        with self.assertRaises(app.AuthRequiredError):
            app.resolve_auth('/api/sessions','GET',self.db,None)

    def test_resolve_auth_private_route_enforces_ownership(self):
        uidA=self._mk_user('own-c'); uidB=self._mk_user('own-d')
        self._mk_campaign('cmp-owned-c',uidA)
        self.db.commit()
        tokenA=app.create_auth_token(self.db,uidA); tokenB=app.create_auth_token(self.db,uidB)
        user=app.resolve_auth('/api/campaigns/cmp-owned-c','GET',self.db,f'epc_session={tokenA}')
        self.assertEqual(user['id'],uidA)
        with self.assertRaises(app.PermissionDeniedError):
            app.resolve_auth('/api/campaigns/cmp-owned-c','GET',self.db,f'epc_session={tokenB}')

    # --- Lot 1c (modularisation) : clonage/creation/import de questionnaires extraits vers epc/templates.py ---

    def test_clone_template_creates_new_version_with_same_content(self):
        t=self.db.execute('select id from templates').fetchone()['id']
        new_id=app.clone_template(self.db,t,name='EPC / SENEVAL')
        self.assertNotEqual(new_id,t)
        old=app.template_payload(self.db,t); new=app.template_payload(self.db,new_id)
        self.assertEqual(new['version'],old['version']+1)
        self.assertEqual(len(new['domains']),len(old['domains']))
        self.assertEqual(sum(len(d['indicators']) for d in new['domains']),sum(len(d['indicators']) for d in old['domains']))

    def test_create_blank_template_and_next_order_increments(self):
        tid=app.create_blank_template(self.db,{'name':'Nouveau'},owner_user_id='owner1')
        self.db.commit()
        self.assertEqual(app.next_order(self.db,'domains','template_id',tid),1)
        did=str(uuid.uuid4())
        self.db.execute("insert into domains values(?,?,?,?,?,?,?)",(did,tid,'d1','Domaine 1','',1,1))
        self.db.commit()
        self.assertEqual(app.next_order(self.db,'domains','template_id',tid),2)

    def test_matrix_xlsx_produces_a_sizeable_workbook(self):
        t=self.db.execute('select id from templates').fetchone()['id']
        xlsx_bytes=app.matrix_xlsx(app.template_payload(self.db,t))
        self.assertGreater(len(xlsx_bytes),1000)

    def test_real_epc_questionnaire_matrix_roundtrips_cleanly(self):
        # The scenario that surfaced both fixed bugs: exporting the real EPC/SENEVAL
        # questionnaire (whose indicators have a non-empty code/label and an always-
        # empty description, see seed_epc) and re-importing it unedited must now work
        # with zero errors and all 70 indicators recovered.
        t=self.db.execute('select id from templates').fetchone()['id']
        template=app.template_payload(self.db,t)
        preview=app.import_preview(app.matrix_xlsx(template))
        self.assertEqual(preview['errors'],[])
        self.assertEqual(preview['rows'],70)
        first=preview['template']['domains'][0]['indicators'][0]
        self.assertEqual(first['code'],'Formation au personnel')
        self.assertEqual(first['label'],'Nous offrons régulièrement la formation au personnel')

    def test_matrix_xlsx_roundtrips_through_import_preview(self):
        # Fixed bugs (see epc/templates.py matrix_xlsx/import_preview):
        # 1) matrix_xlsx() labels the PARAMETRES name cell "Nom du questionnaire
        #    (à remplacer par le vôtre)" (an in-sheet instruction) - import_preview()
        #    now accepts that key too, so a workbook downloaded via matrix_xlsx() can
        #    be re-imported unedited.
        # 2) matrix_xlsx()'s QUESTIONNAIRE sheet writes code -> "Référence" and
        #    label (the actual question statement) -> "Indicateur qualitatif ou
        #    Capacité", matching the EXEMPLE row's own semantics and the EPC
        #    code/label split (indicators.code/indicators.label). import_preview()
        #    reads them back the same way, so the participant-facing question ends
        #    up in "label" again (not the short reference).
        template={'name':'Modele export-import','description':'','version':1,
            'scale':{'type':'numeric','min':1,'max':5,'labels':{}},'priority':{'maxPerDomain':3},
            'domains':[{'label':'Domaine A','indicators':[
                {'code':'Q1','label':'Premiere question complete'},
                {'code':'Q2','label':'Deuxieme question complete'},
            ]}]}
        preview=app.import_preview(app.matrix_xlsx(template))
        self.assertEqual(preview['errors'],[])
        self.assertEqual(preview['template']['name'],'Modele export-import')
        self.assertEqual(preview['rows'],2)
        indicators=preview['template']['domains'][0]['indicators']
        self.assertEqual(indicators[0]['code'],'Q1')
        self.assertEqual(indicators[0]['label'],'Premiere question complete')

    def _questionnaire_xlsx(self,name='Modele test import',rows=(('Domaine A','Q1','Premiere question'),('Domaine A','Q2','Deuxieme question'))):
        # Built directly with xlsxwriter (not via app.matrix_xlsx) so this test targets
        # import_preview()/save_import() in isolation with a plain "Nom du questionnaire"
        # PARAMETRES key (a hand-built workbook, as opposed to one downloaded via
        # app.matrix_xlsx() - see test_matrix_xlsx_roundtrips_through_import_preview).
        # rows are (Domaine, Référence=code, Indicateur qualitatif ou Capacité=label).
        out=BytesIO(); wb=xlsxwriter.Workbook(out,{'in_memory':True})
        ps=wb.add_worksheet('PARAMETRES'); ps.write_row(0,0,['Nom du questionnaire',name]); ps.write_row(1,0,['Description',''])
        qs=wb.add_worksheet('QUESTIONNAIRE'); qs.write_row(0,0,['Domaine','Référence','Indicateur qualitatif ou Capacité'])
        for n,r in enumerate(rows,1): qs.write_row(n,0,list(r))
        wb.close(); return out.getvalue()

    def test_import_preview_reads_valid_questionnaire_workbook(self):
        preview=app.import_preview(self._questionnaire_xlsx())
        self.assertEqual(preview['errors'],[])
        self.assertEqual(preview['template']['name'],'Modele test import')
        self.assertEqual(preview['rows'],2)
        self.assertEqual(len(preview['template']['domains']),1)
        indicators=preview['template']['domains'][0]['indicators']
        self.assertEqual(indicators[0]['code'],'Q1')
        self.assertEqual(indicators[0]['label'],'Premiere question')

    def test_save_import_persists_domains_and_indicators(self):
        preview=app.import_preview(self._questionnaire_xlsx())
        new_id=app.save_import(self.db,preview,owner_user_id='owner2')
        self.db.commit()
        payload=app.template_payload(self.db,new_id)
        self.assertEqual(payload['name'],'Modele test import')
        self.assertEqual(sum(len(d['indicators']) for d in payload['domains']),preview['rows'])
        indicator=payload['domains'][0]['indicators'][0]
        self.assertEqual(indicator['code'],'Q1'); self.assertEqual(indicator['label'],'Premiere question')

    # --- Lot 1e (modularisation) : collecte participant extraite vers epc/collecte.py ---

    def test_create_participant_rejects_closed_session(self):
        sid=self._mk_session('sess-closed')
        self.db.execute("update sessions set status='closed' where id=?",(sid,)); self.db.commit()
        with self.assertRaises(app.CollecteClosedError):
            app.create_participant(self.db,sid,{})

    def test_create_participant_defaults_anonymous_id(self):
        t=self._mk_session('sess-open')
        out=app.create_participant(self.db,'sess-open',{'displayName':'Awa'})
        self.assertTrue(out['anonymousId'].startswith('P-'))
        row=self.db.execute('select * from participants where id=?',(out['id'],)).fetchone()
        self.assertEqual(row['status'],'in_progress'); self.assertEqual(row['display_name'],'Awa')

    def test_submit_response_and_complete_and_resume(self):
        t=self._mk_session('sess-flow')
        pid=app.create_participant(self.db,'sess-flow',{})['id']
        inds=self._indicator_ids(t)
        app.submit_response(self.db,'sess-flow',{'participantId':pid,'indicatorId':inds[0],'value':4})
        resume=app.participant_resume(self.db,'sess-flow',pid)
        self.assertEqual(resume['participant']['status'],'in_progress')
        self.assertEqual(resume['responses'][inds[0]],4)
        self.assertEqual(resume['template']['id'],t)
        app.complete_participant(self.db,pid)
        self.assertEqual(self.db.execute('select status from participants where id=?',(pid,)).fetchone()['status'],'completed')
        # upsert: resubmitting the same indicator updates rather than duplicates
        app.submit_response(self.db,'sess-flow',{'participantId':pid,'indicatorId':inds[0],'value':5})
        self.assertEqual(self.db.execute('select count(*) from responses where participant_id=?',(pid,)).fetchone()[0],1)
        self.assertEqual(app.participant_resume(self.db,'sess-flow',pid)['responses'][inds[0]],5)

    def test_update_participant_display_name(self):
        self._mk_session('sess-name')
        pid=app.create_participant(self.db,'sess-name',{})['id']
        app.update_participant_display_name(self.db,pid,'Nouveau nom')
        self.assertEqual(self.db.execute('select display_name from participants where id=?',(pid,)).fetchone()['display_name'],'Nouveau nom')
        app.update_participant_display_name(self.db,pid,'')
        self.assertIsNone(self.db.execute('select display_name from participants where id=?',(pid,)).fetchone()['display_name'])

    # --- Lot 1f (modularisation) : chaine qualitative extraite vers epc/qualitatif.py ---

    def _mk_priority(self,sid,template_id):
        domain=self.db.execute('select id from domains where template_id=? order by display_order limit 1',(template_id,)).fetchone()['id']
        indicator=self.db.execute('select id from indicators where domain_id=? order by display_order limit 1',(domain,)).fetchone()['id']
        app.toggle_priority(self.db,sid,{'domainId':domain,'indicatorId':indicator,'votes':1})
        return self.db.execute('select id from priorities where session_id=? and indicator_id=?',(sid,indicator)).fetchone()['id'], domain, indicator

    def test_toggle_priority_upserts_votes_then_delete_removes_it(self):
        t=self._mk_session('sess-prio')
        pid,domain,indicator=self._mk_priority('sess-prio',t)
        self.assertEqual(self.db.execute('select votes from priorities where id=?',(pid,)).fetchone()['votes'],1)
        app.toggle_priority(self.db,'sess-prio',{'domainId':domain,'indicatorId':indicator,'votes':3})
        self.assertEqual(self.db.execute('select count(*) from priorities where session_id=?',('sess-prio',)).fetchone()[0],1)
        self.assertEqual(self.db.execute('select votes from priorities where id=?',(pid,)).fetchone()['votes'],3)
        app.delete_priority(self.db,'sess-prio',indicator)
        self.assertIsNone(self.db.execute('select id from priorities where id=?',(pid,)).fetchone())

    def test_priority_analysis_upsert_and_update(self):
        t=self._mk_session('sess-pa')
        pid,_,_=self._mk_priority('sess-pa',t)
        app.upsert_priority_analysis(self.db,'sess-pa',{'priorityId':pid,'problem':'Constat initial'})
        aid=self.db.execute('select id from priority_analyses where session_id=? and priority_id=?',('sess-pa',pid)).fetchone()['id']
        self.assertEqual(self.db.execute('select problem from priority_analyses where id=?',(aid,)).fetchone()['problem'],'Constat initial')
        app.upsert_priority_analysis(self.db,'sess-pa',{'priorityId':pid,'problem':'Constat revise'})
        self.assertEqual(self.db.execute('select count(*) from priority_analyses where session_id=?',('sess-pa',)).fetchone()[0],1)
        self.assertEqual(self.db.execute('select problem from priority_analyses where id=?',(aid,)).fetchone()['problem'],'Constat revise')
        app.update_priority_analysis(self.db,aid,'Constat final')
        self.assertEqual(self.db.execute('select problem from priority_analyses where id=?',(aid,)).fetchone()['problem'],'Constat final')

    def test_analysis_entry_create_update_and_dependency_blocked_delete(self):
        t=self._mk_session('sess-ae')
        pid,_,_=self._mk_priority('sess-ae',t)
        cause_id=app.create_analysis_entry(self.db,'sess-ae',{'priorityId':pid,'kind':'cause','content':'Cause A'})
        app.update_analysis_entry(self.db,cause_id,{'content':'Cause A revisee','validationStatus':'RETENU'})
        self.assertEqual(self.db.execute('select content,validation_status from analysis_entries where id=?',(cause_id,)).fetchone()[:],('Cause A revisee','RETENU'))
        rec_id=app.create_workshop_recommendation(self.db,'sess-ae',{'priorityId':pid,'causeId':cause_id,'title':'Action X','description':'Desc'})
        deleted,dependent=app.delete_analysis_entry(self.db,cause_id,force=False)
        self.assertFalse(deleted); self.assertEqual(dependent,1)
        self.assertIsNotNone(self.db.execute('select id from analysis_entries where id=?',(cause_id,)).fetchone())
        deleted,dependent=app.delete_analysis_entry(self.db,cause_id,force=True)
        self.assertTrue(deleted)
        self.assertIsNone(self.db.execute('select id from analysis_entries where id=?',(cause_id,)).fetchone())
        self.assertIsNone(self.db.execute('select cause_id from workshop_recommendations where id=?',(rec_id,)).fetchone()['cause_id'])

    def test_workshop_recommendation_update_and_dependency_blocked_delete(self):
        t=self._mk_session('sess-rec')
        pid,_,_=self._mk_priority('sess-rec',t)
        rec_id=app.create_workshop_recommendation(self.db,'sess-rec',{'priorityId':pid,'title':'Action Y','description':'Desc'})
        app.update_workshop_recommendation(self.db,rec_id,{'priorityId':pid,'title':'Action Y modifiee','description':'Desc 2','status':'Retenue'})
        self.assertEqual(self.db.execute('select title,status from workshop_recommendations where id=?',(rec_id,)).fetchone()[:],('Action Y modifiee','Retenue'))
        topic_id=app.create_training_topic(self.db,'sess-rec',{'recommendationId':rec_id,'title':'Formation Z'})
        deleted,dependent=app.delete_workshop_recommendation(self.db,rec_id,force=False)
        self.assertFalse(deleted); self.assertEqual(dependent,1)
        deleted,dependent=app.delete_workshop_recommendation(self.db,rec_id,force=True)
        self.assertTrue(deleted)
        self.assertIsNone(self.db.execute('select id from workshop_recommendations where id=?',(rec_id,)).fetchone())
        self.assertIsNone(self.db.execute('select recommendation_id from training_topics where id=?',(topic_id,)).fetchone()['recommendation_id'])

    def test_training_topic_update_and_delete(self):
        t=self._mk_session('sess-tt')
        pid,_,_=self._mk_priority('sess-tt',t)
        topic_id=app.create_training_topic(self.db,'sess-tt',{'priorityId':pid,'title':'Formation initiale'})
        app.update_training_topic(self.db,topic_id,{'priorityId':pid,'title':'Formation renommee'})
        self.assertEqual(self.db.execute('select title from training_topics where id=?',(topic_id,)).fetchone()['title'],'Formation renommee')
        app.delete_training_topic(self.db,topic_id)
        self.assertIsNone(self.db.execute('select id from training_topics where id=?',(topic_id,)).fetchone())

    def test_report_meta_upsert(self):
        self._mk_session('sess-meta')
        app.upsert_report_meta(self.db,'sess-meta',{'facilitator':'Awa','audience':'Equipe','context':'Ctx','conclusion':'Concl'})
        row=self.db.execute('select * from session_report_meta where session_id=?',('sess-meta',)).fetchone()
        self.assertEqual(row['facilitator'],'Awa')
        app.upsert_report_meta(self.db,'sess-meta',{'facilitator':'Awa D.','audience':'Equipe','context':'Ctx','conclusion':'Concl'})
        self.assertEqual(self.db.execute('select count(*) from session_report_meta where session_id=?',('sess-meta',)).fetchone()[0],1)
        self.assertEqual(self.db.execute('select facilitator from session_report_meta where session_id=?',('sess-meta',)).fetchone()['facilitator'],'Awa D.')

    def test_legacy_v1_recommendation_still_works(self):
        t=self._mk_session('sess-legacy')
        indicator=self.db.execute('select id from indicators limit 1').fetchone()['id']
        app.create_legacy_recommendation(self.db,'sess-legacy',{'indicatorId':indicator,'title':'Reco V1'})
        self.assertEqual(self.db.execute('select count(*) from recommendations where session_id=?',('sess-legacy',)).fetchone()[0],1)

    def test_legacy_v1_analysis_note_column_count_bug_is_fixed(self):
        # Was broken since before this modularisation (9 "?" placeholders for the
        # 8-column analysis_notes table, present verbatim on master @909236c) -
        # fixed as a separate, documented change (see epc/qualitatif.py). This
        # legacy V1 route now behaves like create_legacy_recommendation.
        self._mk_session('sess-legacy-fixed')
        indicator=self.db.execute('select id from indicators limit 1').fetchone()['id']
        app.create_analysis_note(self.db,'sess-legacy-fixed',{'indicatorId':indicator,'kind':'cause','content':'Note V1'})
        row=self.db.execute('select * from analysis_notes where session_id=?',('sess-legacy-fixed',)).fetchone()
        self.assertEqual(row['kind'],'cause'); self.assertEqual(row['content'],'Note V1'); self.assertEqual(row['validation_status'],'HYPOTHESE')

    # --- Lot 1g (modularisation) : CRUD questionnaires/domaines/indicateurs extrait vers epc/templates.py ---

    def test_create_update_delete_domain(self):
        tid=app.create_blank_template(self.db,{'name':'Modele domaines'})
        self.db.commit()
        did=app.create_domain(self.db,tid,{'label':'Domaine 1'})
        self.assertEqual(app.template_payload(self.db,tid)['domains'][0]['label'],'Domaine 1')
        app.update_domain(self.db,did,{'label':'Domaine renomme','displayOrder':1,'active':True})
        self.assertEqual(app.template_payload(self.db,tid)['domains'][0]['label'],'Domaine renomme')
        deleted,affected=app.delete_domain(self.db,did)
        self.assertTrue(deleted); self.assertEqual(affected,[])
        self.assertEqual(app.template_payload(self.db,tid)['domains'],[])

    def test_delete_domain_blocked_by_responses(self):
        t=self.db.execute('select id from templates').fetchone()['id']
        sid=self._mk_session('sess-dom-block')
        inds=self._indicator_ids(t)
        domain=self.db.execute('select domain_id from indicators where id=?',(inds[0],)).fetchone()['domain_id']
        self._add_participant('sess-dom-block','p0',t,value=4,n=1)
        deleted,affected=app.delete_domain(self.db,domain)
        self.assertFalse(deleted); self.assertEqual(len(affected),1)

    def test_create_update_delete_indicator(self):
        tid=app.create_blank_template(self.db,{'name':'Modele indicateurs'})
        self.db.commit()
        did=app.create_domain(self.db,tid,{'label':'Domaine 1'})
        iid=app.create_indicator(self.db,did,{'label':'Q1'})
        app.update_indicator(self.db,iid,{'domainId':did,'code':'q1','label':'Q1 modifiee'})
        row=self.db.execute('select code,label from indicators where id=?',(iid,)).fetchone()
        self.assertEqual((row['code'],row['label']),('q1','Q1 modifiee'))
        deleted,used=app.delete_indicator(self.db,iid)
        self.assertTrue(deleted); self.assertEqual(used,0)

    def test_delete_indicator_blocked_by_responses(self):
        t=self.db.execute('select id from templates').fetchone()['id']
        self._mk_session('sess-ind-block')
        self._add_participant('sess-ind-block','p0',t,value=4,n=1)
        indicator=self._indicator_ids(t)[0]
        deleted,used=app.delete_indicator(self.db,indicator)
        self.assertFalse(deleted); self.assertEqual(used,1)

    def test_update_template_forks_when_used_by_a_session(self):
        t=self.db.execute('select id from templates').fetchone()['id']
        self._mk_session('sess-tpl-fork')
        new_id,version_created=app.update_template(self.db,t,{'name':'EPC / SENEVAL modifie'})
        self.assertTrue(version_created); self.assertNotEqual(new_id,t)
        self.assertEqual(self.db.execute('select name from templates where id=?',(t,)).fetchone()['name'],'EPC / SENEVAL')
        self.assertEqual(self.db.execute('select name from templates where id=?',(new_id,)).fetchone()['name'],'EPC / SENEVAL modifie')

    def test_update_template_edits_in_place_when_unused(self):
        tid=app.create_blank_template(self.db,{'name':'Modele libre'})
        self.db.commit()
        new_id,version_created=app.update_template(self.db,tid,{'name':'Modele renomme'})
        self.assertFalse(version_created); self.assertEqual(new_id,tid)
        self.assertEqual(self.db.execute('select name from templates where id=?',(tid,)).fetchone()['name'],'Modele renomme')

    def test_delete_template_outcomes(self):
        epc_id=self.db.execute("select id from templates where name='EPC / SENEVAL'").fetchone()['id']
        self.assertEqual(app.delete_template(self.db,epc_id),'protected')
        tid=app.create_blank_template(self.db,{'name':'Modele suppression'})
        self.db.commit()
        self._mk_session('sess-tpl-del',owner=None)
        self.db.execute('update sessions set template_id=? where id=?',(tid,'sess-tpl-del')); self.db.commit()
        self.assertEqual(app.delete_template(self.db,tid,force=False),'in_use')
        self.assertEqual(app.delete_template(self.db,tid,force=True),'archived')
        tid2=app.create_blank_template(self.db,{'name':'Modele suppression 2'})
        self.db.commit()
        self.assertEqual(app.delete_template(self.db,tid2),'deleted')
        self.assertIsNone(self.db.execute('select id from templates where id=?',(tid2,)).fetchone())

    # --- Lot 1h (modularisation) : CRUD campagnes/groupes/sessions extrait vers epc/campaigns.py ---

    def test_create_campaign_and_update(self):
        uid=self._mk_user('u-camp')
        t=self.db.execute('select id,version from templates').fetchone()
        cid=app.create_campaign(self.db,uid,t['id'],t['version'],{'name':'Campagne creee'})
        camp=self.db.execute('select * from campaigns where id=?',(cid,)).fetchone()
        self.assertEqual(camp['name'],'Campagne creee'); self.assertEqual(camp['owner_user_id'],uid)
        app.update_campaign(self.db,camp,{'name':'Campagne renommee','status':'closed'})
        camp2=self.db.execute('select * from campaigns where id=?',(cid,)).fetchone()
        self.assertEqual(camp2['name'],'Campagne renommee'); self.assertEqual(camp2['status'],'closed')

    def test_create_group_generates_code_and_color(self):
        uid=self._mk_user('u-grp')
        t=self.db.execute('select id,version from templates').fetchone()
        cid=app.create_campaign(self.db,uid,t['id'],t['version'],{'name':'Campagne groupes'})
        camp=self.db.execute('select * from campaigns where id=?',(cid,)).fetchone()
        g1=app.create_group(self.db,camp,uid,{'name':'Groupe A'})
        g2=app.create_group(self.db,camp,uid,{'name':'Groupe B'})
        self.assertNotEqual(g1['groupCode'],g2['groupCode'])
        self.assertNotEqual(g1['groupColor'],g2['groupColor'])
        self.assertTrue(g1['relayToken'])

    def test_regenerate_group_relay_changes_token_hash(self):
        sid=self._mk_session('sess-relay')
        before=app.relay_token_hash('token-a')
        self.db.execute('update sessions set relay_token_hash=? where id=?',(before,'sess-relay')); self.db.commit()
        new_token=app.regenerate_group_relay(self.db,'sess-relay')
        after=self.db.execute('select relay_token_hash from sessions where id=?',('sess-relay',)).fetchone()['relay_token_hash']
        self.assertNotEqual(before,after)
        self.assertEqual(app.relay_token_hash(new_token),after)

    def test_create_session_rejects_questionnaire_without_active_question(self):
        empty_tid=app.create_blank_template(self.db,{'name':'Modele vide'})
        self.db.commit()
        self.assertIsNone(app.create_session(self.db,'owner-x',{'templateId':empty_tid,'name':'Atelier'}))
        t=self.db.execute("select id from templates where name='EPC / SENEVAL'").fetchone()['id']
        sid=app.create_session(self.db,'owner-x',{'templateId':t,'name':'Atelier valide'})
        self.assertIsNotNone(sid)
        self.assertEqual(self.db.execute('select name from sessions where id=?',(sid,)).fetchone()['name'],'Atelier valide')

    def test_update_session_rejects_unknown_template(self):
        self._mk_session('sess-upd')
        self.assertFalse(app.update_session(self.db,'sess-upd',{'name':'x','templateId':'does-not-exist'}))
        self.assertTrue(app.update_session(self.db,'sess-upd',{'name':'Atelier renomme'}))
        self.assertEqual(self.db.execute('select name from sessions where id=?',('sess-upd',)).fetchone()['name'],'Atelier renomme')

if __name__=='__main__': unittest.main()
