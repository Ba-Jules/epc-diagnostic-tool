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

    def _questionnaire_xlsx(self,name='Modele test import',rows=(('Domaine A','Q1','Premiere question'),('Domaine A','Q2','Deuxieme question'))):
        # Built directly with xlsxwriter (not via app.matrix_xlsx) so this test targets
        # import_preview()/save_import() in isolation, using exactly the PARAMETRES keys
        # ("Nom du questionnaire", "Description") that import_preview's QUESTIONNAIRE-sheet
        # path reads. NB: app.matrix_xlsx() itself writes a longer PARAMETRES key ("Nom du
        # questionnaire (à remplacer par le vôtre)") that import_preview does not recognise,
        # so a workbook downloaded via matrix_xlsx() cannot be re-imported unedited as-is -
        # a pre-existing mismatch between the two, unrelated to this refactor (copied
        # verbatim from app.py into epc/templates.py).
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

    def test_save_import_persists_domains_and_indicators(self):
        preview=app.import_preview(self._questionnaire_xlsx())
        new_id=app.save_import(self.db,preview,owner_user_id='owner2')
        self.db.commit()
        payload=app.template_payload(self.db,new_id)
        self.assertEqual(payload['name'],'Modele test import')
        self.assertEqual(sum(len(d['indicators']) for d in payload['domains']),preview['rows'])

if __name__=='__main__': unittest.main()
