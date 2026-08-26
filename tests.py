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
    def _mk_session(self,sid,name='test',campaign_id=None,group_code=None,expected=None,owner=None,profile_schema_id=None):
        t=self.db.execute('select id,version from templates').fetchone()
        self.db.execute("insert into sessions values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",(sid,t['id'],t['version'],name,'','','', 'open',app.now(),None,'',expected,owner,campaign_id,group_code,None,None,None,profile_schema_id))
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
        t=self.db.execute('select id,version from templates').fetchone(); sid='session'; self.db.execute("insert into sessions values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",(sid,t['id'],t['version'],'test','','','', 'open',app.now(),None,'',None,None,None,None,None,None,None,None))
        domain=self.db.execute('select id from domains where display_order=1').fetchone()['id']; inds=self.db.execute('select id from indicators where domain_id=? order by display_order limit 1',(domain,)).fetchone()['id']
        for n,v in [('a',1),('b',5)]:
            pid=n; self.db.execute('insert into participants values(?,?,?,?,?,?,?)',(pid,sid,n,'completed',app.now(),app.now(),None)); self.db.execute('insert into responses values(?,?,?,?,?,?,?,?)',(str(uuid.uuid4()),sid,pid,inds,str(v),'numeric',app.now(),app.now()))
        self.db.commit(); out=app.analysis(self.db,sid); indicator=out['domains'][0]['indicators'][0]
        self.assertEqual(indicator['responses'],2); self.assertEqual(indicator['capacity'],60); self.assertEqual(indicator['consensus'],0); self.assertEqual(app.grade(63,app.GRADING),40)
    def test_reference_questionnaire_fix_never_touches_existing_version(self):
        t=self.db.execute('select id,version from templates').fetchone()
        sid='pinned-session'
        self.db.execute("insert into sessions values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",(sid,t['id'],t['version'],'test','','','', 'open',app.now(),None,'',None,None,None,None,None,None,None,None))
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
        self.db.execute("insert into sessions values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",(sid,t['id'],t['version'],'test','','','', 'open',app.now(),None,'',None,None,None,None,None,None,None,None))
        domain=self.db.execute('select id from domains where display_order=1').fetchone()['id']; indicator=self.db.execute('select id from indicators where domain_id=? limit 1',(domain,)).fetchone()['id']
        self.db.execute('insert into priorities values(?,?,?,?,?,?)',('priority',sid,domain,indicator,1,app.now()))
        self.db.execute('insert into priority_analyses values(?,?,?,?,?,?)',('analysis',sid,'priority','Constat',app.now(),app.now()))
        self.db.execute('insert into analysis_entries values(?,?,?,?,?,?,?,?,?,?,?)',('cause',sid,'priority',None,'cause','Cause','Cause','', 'RETENU',app.now(),app.now()))
        self.db.execute('insert into workshop_recommendations values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)',('recommendation',sid,'priority','cause',None,'Action','Description','Formation','Haute','Responsable','Demain','', 'Retenue',app.now(),app.now()))
        self.db.execute('insert into training_topics values(?,?,?,?,?,?,?,?,?,?,?)',('training',sid,'priority','recommendation','Thème','Besoin','Public','Haute','',app.now(),app.now())); self.db.commit()
        data=app.qualitative_data(self.db,sid)
        self.assertEqual(len(data['priorities']),1); self.assertEqual(len(data['entries']),1); self.assertEqual(data['recommendations'][0]['status'],'Retenue')
        self.assertGreater(len(app.report_xlsx(self.db,sid)),1000)

    def test_report_xlsx_and_docx_raise_clean_error_for_an_unknown_session(self):
        # Regression for an ultrareview finding: report_rows()/report_xlsx() had
        # no guard for analysis() returning None, so an invalid/deleted session
        # id crashed with a raw TypeError instead of a catchable error like
        # report_docx() already raised for the same case.
        with self.assertRaises(ValueError):
            app.report_xlsx(self.db,'no-such-session')
        with self.assertRaises(ValueError):
            app.report_docx(self.db,'no-such-session')

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

    # --- Lot 3 (modularisation, AUDIT_MODULARISATION_8800.md) : golden tests decimaux
    # figeant le comportement de grade()/analysis()/analysis_for() AVANT extraction vers
    # epc/scoring.py - bornes de graduation, N=0/N=1/N>1, "manquants", ponderation entre
    # domaines. Doivent rester vrais a l'identique apres le deplacement du code. ---

    def test_grade_boundaries_exact(self):
        norm=app.GRADING
        self.assertEqual(app.grade(0,norm),5)
        self.assertEqual(app.grade(22,norm),5)
        self.assertEqual(app.grade(23,norm),10)
        self.assertEqual(app.grade(59,norm),35)
        self.assertEqual(app.grade(60,norm),40)
        self.assertEqual(app.grade(98,norm),95)
        self.assertEqual(app.grade(99,norm),100)
        self.assertEqual(app.grade(100,norm),100)
        self.assertEqual(app.grade(-5,norm),5)
        self.assertEqual(app.grade(150,norm),100)
        self.assertEqual(app.grade(22.4,norm),5)
        self.assertEqual(app.grade(22.6,norm),10)
        self.assertIsNone(app.grade(None,norm))

    def test_analysis_single_respondent_consensus_not_calculable(self):
        t=self._mk_session('sess-n1')
        self._add_participant('sess-n1','solo',t,value=3,n=1)
        self.db.commit()
        out=app.analysis(self.db,'sess-n1')
        ind=out['domains'][0]['indicators'][0]
        self.assertEqual(ind['responses'],1); self.assertEqual(ind['capacity'],60.0)
        self.assertIsNone(ind['consensus']); self.assertEqual(ind['consensusNote'],'single_respondent')
        dom=out['domains'][0]
        self.assertEqual(dom['responses'],1); self.assertIsNone(dom['consensus'])
        self.assertEqual(dom['consensusNote'],'single_respondent')
        self.assertEqual(out['global']['consensusNote'],'single_respondent')
        self.assertIsNone(out['global']['consensus'])

    def test_analysis_zero_participants_returns_none_scores(self):
        self._mk_session('sess-n0')
        self.db.commit()
        out=app.analysis(self.db,'sess-n0')
        self.assertEqual(out['participantCount'],0); self.assertEqual(out['completedCount'],0)
        ind=out['domains'][0]['indicators'][0]
        self.assertEqual(ind['responses'],0); self.assertEqual(ind['missing'],0)
        self.assertIsNone(ind['capacity']); self.assertIsNone(ind['consensus'])
        self.assertIsNone(out['global']['capacity']); self.assertIsNone(out['global']['consensus'])
        self.assertEqual(out['global']['responses'],0)

    def test_analysis_missing_count_reflects_total_participants(self):
        t=self._mk_session('sess-missing')
        self._add_participant('sess-missing','p1',t,value=4,status='completed')
        self._add_participant('sess-missing','p2',t,value=None,status='in_progress')
        self.db.commit()
        out=app.analysis(self.db,'sess-missing')
        self.assertEqual(out['participantCount'],2)
        ind=out['domains'][0]['indicators'][0]
        self.assertEqual(ind['responses'],1); self.assertEqual(ind['missing'],1)

    def test_global_capacity_is_unweighted_mean_across_domains(self):
        # Proves global capacity averages domain capacities directly (the reference
        # KOICA tool's "Moyenne" row), never a response-weighted pool - domain 0 has
        # 1 respondent, domain 1 has 4, and a weighted average would be far from 60.
        t=self._mk_session('sess-domain-weight')
        domains=[r['id'] for r in self.db.execute('select id from domains where template_id=? order by display_order',(t,))]
        dom0_inds=[r['id'] for r in self.db.execute('select id from indicators where domain_id=? order by display_order',(domains[0],))]
        dom1_inds=[r['id'] for r in self.db.execute('select id from indicators where domain_id=? order by display_order',(domains[1],))]
        def add(pid,inds,value):
            self.db.execute('insert into participants values(?,?,?,?,?,?,?)',(pid,'sess-domain-weight',pid,'completed',app.now(),app.now(),None))
            for iid in inds:
                self.db.execute('insert into responses values(?,?,?,?,?,?,?,?)',(str(uuid.uuid4()),'sess-domain-weight',pid,iid,str(value),'numeric',app.now(),app.now()))
        add('p1',dom0_inds,5)
        for pid in ('p2','p3','p4','p5'): add(pid,dom1_inds,1)
        self.db.commit()
        out=app.analysis(self.db,'sess-domain-weight')
        dom_by_id={d['id']:d for d in out['domains']}
        self.assertEqual(dom_by_id[domains[0]]['capacity'],100.0)
        self.assertEqual(dom_by_id[domains[1]]['capacity'],20.0)
        self.assertEqual(out['global']['capacity'],60.0)

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

    def test_cross_pilot_profile_field_access_denied(self):
        # Regression for an ultrareview finding: enforce_ownership() had branches
        # for /api/profile-schemas/ but none for /api/profile-fields/, so a PUT/
        # DELETE on another pilot's field silently skipped the ownership check.
        uidA=self._mk_user('own-field-a'); uidB=self._mk_user('own-field-b')
        schema_id,fields=self._mk_schema_with_fields(owner=uidA)
        self.db.commit()
        userA={'id':uidA,'role':'pilote'}; userB={'id':uidB,'role':'pilote'}; admin={'id':'adm-field','role':'admin'}
        field_path='/api/profile-fields/'+fields['organisation']
        with self.assertRaises(app.PermissionDeniedError):
            app.enforce_ownership(field_path,self.db,userB)
        app.enforce_ownership(field_path,self.db,userA)
        app.enforce_ownership(field_path,self.db,admin)

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

    # --- Mission parite :8810->:8820 (consignes_claude.txt) : profil EPC/SENEVAL
    # par defaut - PAS des colonnes en dur, seulement les valeurs par defaut du
    # modele EPC/SENEVAL creees via le moteur generique profile_schema/
    # profile_fields - epc/profile.py ensure_default_profile_schema() ---

    def test_create_session_attaches_default_epc_profile_with_five_dimensions(self):
        t=self.db.execute("select id from templates where name='EPC / SENEVAL'").fetchone()['id']
        sid=app.create_session(self.db,'owner-epc',{'templateId':t,'name':'Atelier EPC'})
        schema_id=self.db.execute('select profile_schema_id from sessions where id=?',(sid,)).fetchone()['profile_schema_id']
        self.assertIsNotNone(schema_id)
        dims=app.available_dimensions(self.db,sid)
        self.assertEqual(len(dims),5)
        self.assertEqual({d['fieldKey'] for d in dims},{'type-de-participant','profil','sexe','tranche-dage','niveau-de-scolarisation'})

    def test_default_epc_profile_fields_are_modifiable_and_deletable(self):
        # "modifiable/desactivable par le pilote" (consignes_claude.txt) : une
        # fois crees, ce sont des profile_fields ordinaires - meme API que
        # n'importe quel champ saisi a la main.
        t=self.db.execute("select id from templates where name='EPC / SENEVAL'").fetchone()['id']
        sid=app.create_session(self.db,'owner-epc2',{'templateId':t,'name':'Atelier EPC 2'})
        schema_id=self.db.execute('select profile_schema_id from sessions where id=?',(sid,)).fetchone()['profile_schema_id']
        payload=app.profile_schema_payload(self.db,schema_id)
        sexe_field=next(f for f in payload['fields'] if f['field_key']=='sexe')
        app.update_profile_field(self.db,sexe_field['id'],{'active':False})
        deleted,used=app.delete_profile_field(self.db,sexe_field['id'])
        self.assertTrue(deleted)

    def test_create_session_respects_an_explicit_profile_schema_id(self):
        schema_id,_=self._mk_schema_with_fields(owner='owner-custom')
        t=self.db.execute("select id from templates where name='EPC / SENEVAL'").fetchone()['id']
        sid=app.create_session(self.db,'owner-epc3',{'templateId':t,'name':'Atelier profil custom','profileSchemaId':schema_id})
        self.assertEqual(self.db.execute('select profile_schema_id from sessions where id=?',(sid,)).fetchone()['profile_schema_id'],schema_id)

    def test_create_session_on_a_blank_template_gets_no_default_profile(self):
        # resolve_model_key() (epc/restitution.py) fait retomber tout modele
        # custom/vide sur "epc_seneval" pour les BESOINS DU RAPPORT uniquement -
        # ensure_default_profile_schema() ne doit jamais s'appuyer sur ce
        # fallback, sinon un questionnaire personnalise heriterait a tort des
        # 5 dimensions EPC (Type/Profil/Sexe/Age/Scolarisation n'ont aucun sens
        # pour un modele qui n'est pas reellement EPC/SENEVAL).
        blank_tid=app.create_blank_template(self.db,{'name':'Modele sur mesure'})
        app.create_domain(self.db,blank_tid,{'label':'Domaine'})
        domain_id=self.db.execute('select id from domains where template_id=?',(blank_tid,)).fetchone()['id']
        app.create_indicator(self.db,domain_id,{'label':'Indicateur'})
        self.db.commit()
        sid=app.create_session(self.db,'owner-blank',{'templateId':blank_tid,'name':'Atelier sur mesure'})
        self.assertIsNone(self.db.execute('select profile_schema_id from sessions where id=?',(sid,)).fetchone()['profile_schema_id'])

    def test_create_group_also_attaches_default_epc_profile(self):
        uid=self._mk_user('u-grp-profile')
        t=self.db.execute("select id,version from templates where name='EPC / SENEVAL'").fetchone()
        cid=app.create_campaign(self.db,uid,t['id'],t['version'],{'name':'Campagne EPC'})
        camp=self.db.execute('select * from campaigns where id=?',(cid,)).fetchone()
        g=app.create_group(self.db,camp,uid,{'name':'Groupe A'})
        schema_id=self.db.execute('select profile_schema_id from sessions where id=?',(g['id'],)).fetchone()['profile_schema_id']
        self.assertIsNotNone(schema_id)
        self.assertEqual(len(app.available_dimensions(self.db,g['id'])),5)

    def test_update_session_rejects_unknown_template(self):
        self._mk_session('sess-upd')
        self.assertFalse(app.update_session(self.db,'sess-upd',{'name':'x','templateId':'does-not-exist'}))
        self.assertTrue(app.update_session(self.db,'sess-upd',{'name':'Atelier renomme'}))
        self.assertEqual(self.db.execute('select name from sessions where id=?',('sess-upd',)).fetchone()['name'],'Atelier renomme')

    # --- Lot 2a (modularisation, AUDIT_MODULARISATION_8800.md) : identite canonique stable ---
    # Colonnes additives (model_key/is_canonical) sur templates, encore inertes : rien ne
    # les lit ailleurs dans le moteur a ce stade, seule leur coherence est testee ici.

    def test_epc_seneval_is_tagged_canonical_after_init(self):
        row=self.db.execute("select model_key,is_canonical,version from templates where name='EPC / SENEVAL'").fetchone()
        self.assertEqual(row['model_key'],app.MODEL_KEY_EPC_SENEVAL)
        self.assertEqual(row['is_canonical'],1)

    def test_exactly_one_canonical_row_per_model_key(self):
        rows_=self.db.execute("select id from templates where model_key=? and is_canonical=1",(app.MODEL_KEY_EPC_SENEVAL,)).fetchall()
        self.assertEqual(len(rows_),1)

    def test_custom_templates_are_not_tagged_canonical(self):
        tid=app.create_blank_template(self.db,{'name':'Modele perso'})
        self.db.commit()
        app.ensure_model_identity(self.db)
        row=self.db.execute('select model_key,is_canonical from templates where id=?',(tid,)).fetchone()
        self.assertIsNone(row['model_key']); self.assertEqual(row['is_canonical'],0)

    def test_ensure_model_identity_is_idempotent(self):
        app.ensure_model_identity(self.db)
        before=[dict(r) for r in self.db.execute('select id,model_key,is_canonical from templates order by id')]
        app.ensure_model_identity(self.db)
        after=[dict(r) for r in self.db.execute('select id,model_key,is_canonical from templates order by id')]
        self.assertEqual(before,after)

    def test_new_reference_version_becomes_the_only_canonical_one(self):
        # Mirrors test_reference_questionnaire_fix_never_touches_existing_version's setup:
        # force a stale referential so ensure_reference_questionnaire_version() inserts a
        # new version, then check ensure_model_identity() moves is_canonical onto it and
        # off the old one - without touching the old version's rows/history.
        old_tid=self.db.execute("select id from templates where name='EPC / SENEVAL'").fetchone()['id']
        domain=self.db.execute('select id from domains where template_id=? order by display_order limit 1',(old_tid,)).fetchone()['id']
        self.db.execute('update domains set code=? where id=?',('stale-code',domain)); self.db.commit()
        app.ensure_reference_questionnaire_version(self.db)
        app.ensure_model_identity(self.db)
        new_tid=self.db.execute("select id from templates where name='EPC / SENEVAL' order by version desc limit 1").fetchone()['id']
        self.assertNotEqual(new_tid,old_tid)
        self.assertEqual(self.db.execute('select is_canonical from templates where id=?',(old_tid,)).fetchone()['is_canonical'],0)
        self.assertEqual(self.db.execute('select is_canonical from templates where id=?',(new_tid,)).fetchone()['is_canonical'],1)
        self.assertEqual(self.db.execute('select code from domains where id=?',(domain,)).fetchone()['code'],'stale-code')

    # --- Lot 2b (modularisation, AUDIT_MODULARISATION_8800.md) : branchement de la
    # detection canonique sur model_key/is_canonical a la place du nom ---

    def test_renamed_reference_template_stays_protected_from_deletion(self):
        tid=self.db.execute("select id from templates where name='EPC / SENEVAL'").fetchone()['id']
        self.db.execute("update templates set name=? where id=?",('EPC / SENEVAL (renomme)',tid)); self.db.commit()
        self.assertEqual(app.delete_template(self.db,tid),'protected')

    def test_renamed_reference_template_stays_ownerless_after_migration(self):
        tid=self.db.execute("select id from templates where name='EPC / SENEVAL'").fetchone()['id']
        self.db.execute("update templates set name=? where id=?",('EPC / SENEVAL (renomme)',tid)); self.db.commit()
        uid=self._mk_user('u-first-account'); self.db.commit()
        app.migrate_v2_ownership(self.db)
        self.assertIsNone(self.db.execute('select owner_user_id from templates where id=?',(tid,)).fetchone()['owner_user_id'])

    def test_migrate_v2_ownership_still_assigns_custom_templates(self):
        tid=app.create_blank_template(self.db,{'name':'Modele perso ownership'})
        self.db.commit()
        uid=self._mk_user('u-second-account'); self.db.commit()
        app.migrate_v2_ownership(self.db)
        self.assertEqual(self.db.execute('select owner_user_id from templates where id=?',(tid,)).fetchone()['owner_user_id'],uid)

    # --- Lot 2c (modularisation, AUDIT_MODULARISATION_8800.md) : la reference EPC
    # se retrouve par model_key, plus seulement par nom, dans ensure_reference_
    # questionnaire_version() elle-meme (pas seulement dans les lecteurs du lot 2b) ---

    def test_renaming_the_only_reference_version_does_not_spawn_a_duplicate(self):
        # Was broken before this fix: renaming the sole EPC/SENEVAL version made
        # ensure_reference_questionnaire_version() find no row by name, so it called
        # seed_epc() again and silently created a brand new default "EPC / SENEVAL" v1,
        # orphaning the renamed one.
        tid=self.db.execute("select id from templates where name='EPC / SENEVAL'").fetchone()['id']
        self.db.execute("update templates set name=? where id=?",('EPC / SENEVAL (renomme)',tid)); self.db.commit()
        app.ensure_reference_questionnaire_version(self.db)
        app.ensure_model_identity(self.db)
        rows_=self.db.execute("select id,name from templates where model_key=?",(app.MODEL_KEY_EPC_SENEVAL,)).fetchall()
        self.assertEqual(len(rows_),1)
        self.assertEqual(rows_[0]['id'],tid)
        self.assertEqual(rows_[0]['name'],'EPC / SENEVAL (renomme)')

    def test_renaming_the_latest_of_several_versions_does_not_spawn_a_duplicate(self):
        # Same bug, narrower trigger: with two versions, renaming only the latest one
        # used to make the name-based lookup fall back to the older (stale) version,
        # which then looked out-of-date and triggered a spurious next-version insert.
        old_tid=self.db.execute("select id from templates where name='EPC / SENEVAL'").fetchone()['id']
        domain=self.db.execute('select id from domains where template_id=? order by display_order limit 1',(old_tid,)).fetchone()['id']
        self.db.execute('update domains set code=? where id=?',('stale-code',domain)); self.db.commit()
        app.ensure_reference_questionnaire_version(self.db); app.ensure_model_identity(self.db)
        new_tid=self.db.execute("select id from templates where model_key=? order by version desc limit 1",(app.MODEL_KEY_EPC_SENEVAL,)).fetchone()['id']
        self.assertNotEqual(new_tid,old_tid)
        self.db.execute("update templates set name=? where id=?",('EPC / SENEVAL (renomme)',new_tid)); self.db.commit()
        app.ensure_reference_questionnaire_version(self.db)
        app.ensure_model_identity(self.db)
        rows_=self.db.execute("select id,name,version from templates where model_key=? order by version",(app.MODEL_KEY_EPC_SENEVAL,)).fetchall()
        self.assertEqual(len(rows_),2)
        self.assertEqual(rows_[-1]['id'],new_tid)
        self.assertEqual(rows_[-1]['name'],'EPC / SENEVAL (renomme)')

    # --- Lot 4a (modularisation, AUDIT_MODULARISATION_8800.md) : profil participant
    # composable, backend seul, desactive par defaut (sessions.profile_schema_id NULL) ---

    def _mk_schema_with_fields(self,owner='owner-profile'):
        schema_id=app.create_profile_schema(self.db,owner,{'name':'Profil standard'})
        text_id=app.create_profile_field(self.db,schema_id,{'fieldType':'text','label':'Organisation'})
        num_id=app.create_profile_field(self.db,schema_id,{'fieldType':'number','label':'Age','required':True})
        single_id=app.create_profile_field(self.db,schema_id,{'fieldType':'single_choice','label':'Genre','options':['F','M','Autre']})
        multi_id=app.create_profile_field(self.db,schema_id,{'fieldType':'multi_choice','label':'Langues','options':['fr','wo','en']})
        self.db.commit()
        return schema_id,{'organisation':text_id,'age':num_id,'genre':single_id,'langues':multi_id}

    def test_create_profile_schema_requires_a_name(self):
        with self.assertRaises(ValueError):
            app.create_profile_schema(self.db,'owner-x',{'name':''})

    def test_set_participant_profile_values_matches_choice_options_across_json_types(self):
        # Regression for an ultrareview finding: choice validation compared
        # raw equality, so an option stored as a number (possible via direct
        # API use, not the textarea UI which only ever sends strings)
        # rejected an equal-looking submitted string as invalid.
        schema_id=app.create_profile_schema(self.db,'owner-types',{'name':'Profil types'})
        app.create_profile_field(self.db,schema_id,{'fieldType':'single_choice','label':'Niveau','options':[1,2,3]})
        sid='sess-choice-types'; self._mk_session(sid,profile_schema_id=schema_id)
        pid=app.create_participant(self.db,sid,{})['id']
        app.set_participant_profile_values(self.db,pid,{'niveau':'2'})
        self.assertEqual(app.get_participant_profile_values(self.db,pid),{'niveau':'2'})
        with self.assertRaises(ValueError):
            app.set_participant_profile_values(self.db,pid,{'niveau':'9'})

    def test_create_profile_field_respects_an_explicit_display_order_of_zero(self):
        # Regression for an ultrareview finding: `data.get("displayOrder") or
        # next_order(...)` treated an explicit 0 as falsy and silently replaced
        # it with the auto-computed order.
        schema_id=app.create_profile_schema(self.db,'owner-order',{'name':'Profil ordre'})
        fid=app.create_profile_field(self.db,schema_id,{'fieldType':'text','label':'Premier','displayOrder':0})
        self.assertEqual(self.db.execute('select display_order from profile_fields where id=?',(fid,)).fetchone()['display_order'],0)

    def test_profile_field_types_require_options_for_choice_fields(self):
        schema_id=app.create_profile_schema(self.db,'owner-x',{'name':'Profil'}); self.db.commit()
        with self.assertRaises(ValueError):
            app.create_profile_field(self.db,schema_id,{'fieldType':'single_choice','label':'Genre','options':[]})
        with self.assertRaises(ValueError):
            app.create_profile_field(self.db,schema_id,{'fieldType':'bogus_type','label':'X'})

    def test_profile_schema_payload_lists_typed_fields_in_order(self):
        schema_id,fields=self._mk_schema_with_fields()
        payload=app.profile_schema_payload(self.db,schema_id)
        self.assertEqual(payload['name'],'Profil standard')
        self.assertEqual([f['field_type'] for f in payload['fields']],['text','number','single_choice','multi_choice'])
        self.assertEqual(payload['fields'][2]['options'],['F','M','Autre'])

    def test_update_and_delete_profile_field(self):
        schema_id,fields=self._mk_schema_with_fields()
        app.update_profile_field(self.db,fields['organisation'],{'label':'Organisation renommee'})
        self.assertEqual(self.db.execute('select label from profile_fields where id=?',(fields['organisation'],)).fetchone()['label'],'Organisation renommee')
        deleted,used=app.delete_profile_field(self.db,fields['organisation'])
        self.assertTrue(deleted); self.assertEqual(used,0)

    def test_delete_profile_field_blocked_once_a_value_exists(self):
        schema_id,fields=self._mk_schema_with_fields()
        sid='sess-profile-del-field'; self._mk_session(sid,profile_schema_id=schema_id)
        pid=app.create_participant(self.db,sid,{})['id']
        app.set_participant_profile_values(self.db,pid,{'organisation':'ONG Test'})
        deleted,used=app.delete_profile_field(self.db,fields['organisation'])
        self.assertFalse(deleted); self.assertEqual(used,1)

    def test_delete_profile_schema_blocked_when_session_uses_it(self):
        schema_id,fields=self._mk_schema_with_fields()
        self._mk_session('sess-profile-del-schema',profile_schema_id=schema_id)
        self.assertEqual(app.delete_profile_schema(self.db,schema_id),'in_use')

    def test_delete_profile_schema_succeeds_when_unused(self):
        schema_id,fields=self._mk_schema_with_fields()
        self.assertEqual(app.delete_profile_schema(self.db,schema_id),'deleted')
        self.assertIsNone(self.db.execute('select id from profile_schemas where id=?',(schema_id,)).fetchone())

    def test_set_participant_profile_values_validates_each_type(self):
        schema_id,fields=self._mk_schema_with_fields()
        sid='sess-profile-values'; self._mk_session(sid,profile_schema_id=schema_id)
        pid=app.create_participant(self.db,sid,{})['id']
        app.set_participant_profile_values(self.db,pid,{'organisation':'ONG Test','age':34,'genre':'F','langues':['fr','wo']})
        values=app.get_participant_profile_values(self.db,pid)
        self.assertEqual(values,{'organisation':'ONG Test','age':34,'genre':'F','langues':['fr','wo']})
        # upsert: resubmitting a field updates rather than duplicates
        app.set_participant_profile_values(self.db,pid,{'age':35})
        self.assertEqual(app.get_participant_profile_values(self.db,pid)['age'],35)
        self.assertEqual(self.db.execute('select count(*) from participant_profile_values where participant_id=?',(pid,)).fetchone()[0],4)

    def test_set_participant_profile_values_rejects_invalid_type_and_choice(self):
        schema_id,fields=self._mk_schema_with_fields()
        sid='sess-profile-invalid'; self._mk_session(sid,profile_schema_id=schema_id)
        pid=app.create_participant(self.db,sid,{})['id']
        with self.assertRaises(ValueError): app.set_participant_profile_values(self.db,pid,{'age':'not-a-number'})
        with self.assertRaises(ValueError): app.set_participant_profile_values(self.db,pid,{'genre':'Inconnu'})
        with self.assertRaises(ValueError): app.set_participant_profile_values(self.db,pid,{'langues':['fr','klingon']})
        with self.assertRaises(ValueError): app.set_participant_profile_values(self.db,pid,{'unknown_field':'x'})

    def test_set_participant_profile_values_enforces_required(self):
        schema_id,fields=self._mk_schema_with_fields()
        sid='sess-profile-required'; self._mk_session(sid,profile_schema_id=schema_id)
        pid=app.create_participant(self.db,sid,{})['id']
        with self.assertRaises(ValueError): app.set_participant_profile_values(self.db,pid,{'age':None})
        app.set_participant_profile_values(self.db,pid,{'organisation':''})  # optional field: empty is fine
        self.assertEqual(app.get_participant_profile_values(self.db,pid),{})

    def test_participant_resume_includes_profile_when_session_has_a_schema(self):
        schema_id,fields=self._mk_schema_with_fields()
        sid='sess-profile-resume'; self._mk_session(sid,profile_schema_id=schema_id)
        pid=app.create_participant(self.db,sid,{})['id']
        app.set_participant_profile_values(self.db,pid,{'organisation':'ONG Test'})
        resume=app.participant_resume(self.db,sid,pid)
        self.assertEqual(resume['profile']['name'],'Profil standard')
        self.assertEqual(resume['profileValues'],{'organisation':'ONG Test'})

    def test_participant_resume_omits_profile_for_pre_existing_sessions(self):
        # Migration safety: a session created before this feature (profile_schema_id
        # NULL, the only possible value before lot 4a) must resume exactly as before.
        sid='sess-no-profile'; self._mk_session(sid)
        pid=app.create_participant(self.db,sid,{})['id']
        resume=app.participant_resume(self.db,sid,pid)
        self.assertIsNone(resume['profile'])
        self.assertEqual(resume['profileValues'],{})

    def test_set_participant_profile_values_refused_without_a_schema(self):
        sid='sess-no-profile-2'; self._mk_session(sid)
        pid=app.create_participant(self.db,sid,{})['id']
        with self.assertRaises(ValueError):
            app.set_participant_profile_values(self.db,pid,{'organisation':'x'})

    def test_profile_submission_route_is_public(self):
        # Privacy/access check (audit: "aucune fuite relais/export") from the other
        # direction - a participant must be able to submit their OWN profile without
        # a pilot cookie, exactly like responses/complete already are.
        self.assertTrue(app.is_public_api('/api/participants/some-id/profile','POST'))
        self.assertFalse(app.is_public_api('/api/participants/some-id/profile','GET'))
        self.assertFalse(app.is_public_api('/api/profile-schemas','GET'))
        self.assertFalse(app.is_public_api('/api/profile-schemas','POST'))

    def test_relay_payload_never_includes_participant_profile_data(self):
        # Privacy check called out by the audit ("aucune fuite relais/export"): the
        # public relay dashboard only ever exposes aggregate counts, never
        # per-participant details - confirmed by inspecting its actual key set.
        schema_id,fields=self._mk_schema_with_fields()
        cid='camp-relay-privacy'; uid=self._mk_user('u-relay-privacy')
        self._mk_campaign(cid,uid)
        sid='sess-relay-privacy'; self._mk_session(sid,campaign_id=cid,profile_schema_id=schema_id)
        self.db.execute("update sessions set relay_token_hash=? where id=?",(app.relay_token_hash('tok-privacy'),sid)); self.db.commit()
        pid=app.create_participant(self.db,sid,{})['id']
        app.set_participant_profile_values(self.db,pid,{'organisation':'Donnee privee'})
        g=self.db.execute("SELECT s.*, c.name AS campaign_name FROM sessions s LEFT JOIN campaigns c ON c.id=s.campaign_id WHERE s.relay_token_hash=?",(app.relay_token_hash('tok-privacy'),)).fetchone()
        self.assertIsNotNone(g)
        relay_keys={'campaignName','groupName','relayName','groupCode','groupColor','expectedParticipants','participantCount','completedCount','participantLink'}
        self.assertNotIn('profile',relay_keys); self.assertNotIn('profileValues',relay_keys)

    def test_session_delete_cascade_removes_participant_profile_values(self):
        # Reproduces a real bug caught manually while testing lot 4a: deleting a
        # session/group/campaign cascades through SESSION_CHILD_TABLES, and
        # participant_profile_values wasn't in that list - deleting participants
        # while their profile values still referenced them raised a FK
        # IntegrityError instead of cleanly cascading, for any session that had
        # ever collected profile data. Also checks the FK-order-sensitive delete
        # actually succeeds (not just that the table is listed).
        schema_id,fields=self._mk_schema_with_fields()
        sid='sess-cascade-profile'; self._mk_session(sid,profile_schema_id=schema_id)
        pid=app.create_participant(self.db,sid,{})['id']
        app.set_participant_profile_values(self.db,pid,{'organisation':'ONG Test'})
        self.db.commit()
        for table in app.SESSION_CHILD_TABLES:
            self.db.execute(f"DELETE FROM {table} WHERE session_id=?",(sid,))
        self.db.execute("DELETE FROM sessions WHERE id=?",(sid,))
        self.db.commit()
        self.assertEqual(self.db.execute('select count(*) from participant_profile_values where session_id=?',(sid,)).fetchone()[0],0)
        self.assertIsNone(self.db.execute('select id from sessions where id=?',(sid,)).fetchone())

    # --- Lot 4c (modularisation, AUDIT_MODULARISATION_8800.md) : liste des participants
    # (pilote) avec valeurs de profil - epc/collecte.py list_session_participants() ---

    def test_list_session_participants_includes_profile_values(self):
        schema_id,fields=self._mk_schema_with_fields()
        sid='sess-roster'; self._mk_session(sid,profile_schema_id=schema_id)
        p1=app.create_participant(self.db,sid,{'displayName':'Awa'})['id']
        app.set_participant_profile_values(self.db,p1,{'organisation':'ONG A','age':40})
        p2=app.create_participant(self.db,sid,{})['id']
        roster=app.list_session_participants(self.db,sid)
        self.assertEqual(len(roster),2)
        byId={p['id']:p for p in roster}
        self.assertEqual(byId[p1]['profileValues'],{'organisation':'ONG A','age':40})
        self.assertEqual(byId[p2]['profileValues'],{})
        self.assertEqual(byId[p1]['display_name'],'Awa')

    def test_list_session_participants_empty_profile_values_without_schema(self):
        sid='sess-roster-no-profile'; self._mk_session(sid)
        pid=app.create_participant(self.db,sid,{})['id']
        roster=app.list_session_participants(self.db,sid)
        self.assertEqual(len(roster),1)
        self.assertEqual(roster[0]['profileValues'],{})

    def test_list_session_participants_ordered_by_start_time(self):
        sid='sess-roster-order'; self._mk_session(sid)
        first=app.create_participant(self.db,sid,{})['id']
        second=app.create_participant(self.db,sid,{})['id']
        roster=app.list_session_participants(self.db,sid)
        self.assertEqual([p['id'] for p in roster],[first,second])

    # --- Lot 5 (modularisation, AUDIT_MODULARISATION_8800.md) : dimensions
    # d'analyse - tout champ de profil categoriel explicitement marque par le
    # pilote devient filtrable. epc/profile.py available_dimensions() /
    # resolve_dimension_field() (seule porte d'entree du filtre - garde-fou
    # vie privee) ; epc/scoring.py analysis_for(participant_ids=...) /
    # dimension_analysis() (garde-fou petits N) ---

    def test_create_profile_field_rejects_is_dimension_on_non_categorical_type(self):
        schema_id,_=self._mk_schema_with_fields()
        with self.assertRaises(ValueError):
            app.create_profile_field(self.db,schema_id,{'fieldType':'text','label':'Notes','isDimension':True})

    def test_update_profile_field_rejects_is_dimension_on_non_categorical_type(self):
        schema_id,fields=self._mk_schema_with_fields()
        with self.assertRaises(ValueError):
            app.update_profile_field(self.db,fields['age'],{'isDimension':True})

    def test_available_dimensions_lists_only_flagged_categorical_fields(self):
        schema_id,fields=self._mk_schema_with_fields()
        app.update_profile_field(self.db,fields['genre'],{'isDimension':True})
        sid='sess-dims'; self._mk_session(sid,profile_schema_id=schema_id)
        dims=app.available_dimensions(self.db,sid)
        self.assertEqual([d['fieldKey'] for d in dims],['genre'])
        self.assertEqual(dims[0]['options'],['F','M','Autre'])

    def test_available_dimensions_empty_without_a_profile_schema(self):
        sid='sess-dims-none'; self._mk_session(sid)
        self.assertEqual(app.available_dimensions(self.db,sid),[])

    def test_dimension_analysis_refuses_unflagged_field(self):
        # The privacy gate: a categorical field the pilot never explicitly
        # flagged must stay unfilterable, even though it exists on the schema.
        schema_id,fields=self._mk_schema_with_fields()
        sid='sess-dim-refuse'; self._mk_session(sid,profile_schema_id=schema_id)
        with self.assertRaises(ValueError):
            app.dimension_analysis(self.db,sid,'genre','F')

    def test_dimension_analysis_refuses_when_session_has_no_profile(self):
        sid='sess-dim-no-profile'; self._mk_session(sid)
        with self.assertRaises(ValueError):
            app.dimension_analysis(self.db,sid,'genre','F')

    def _mk_dimension_session(self,sid,n_f=5,n_m=5,value_f=5,value_m=1):
        schema_id,fields=self._mk_schema_with_fields()
        app.update_profile_field(self.db,fields['genre'],{'isDimension':True})
        tid=self._mk_session(sid,profile_schema_id=schema_id)
        for i in range(n_f):
            pid=f'{sid}-f{i}'; self._add_participant(sid,pid,tid,value=value_f,n=1)
            app.set_participant_profile_values(self.db,pid,{'genre':'F'})
        for i in range(n_m):
            pid=f'{sid}-m{i}'; self._add_participant(sid,pid,tid,value=value_m,n=1)
            app.set_participant_profile_values(self.db,pid,{'genre':'M'})
        self.db.commit()
        return schema_id

    def test_dimension_analysis_restricts_capacity_to_matching_participants(self):
        # 5 "F" answer 5/5 (capacity 100), 5 "M" answer 1/5 (capacity 20): the
        # whole-session mean (3/5 => 60) must NOT leak into either filtered
        # cohort - each is recomputed strictly from its own responses.
        sid='sess-dim-single'; self._mk_dimension_session(sid)
        whole=app.analysis(self.db,sid)
        self.assertEqual(whole['completedCount'],10)
        self.assertEqual(whole['domains'][0]['capacity'],60)
        filtered=app.dimension_analysis(self.db,sid,'genre','F')
        self.assertEqual(filtered['completedCount'],5)
        self.assertEqual(filtered['domains'][0]['capacity'],100)
        self.assertEqual(filtered['dimension'],{'fieldKey':'genre','fieldLabel':'Genre','value':'F','minRequired':app.MIN_COHORT_N,'suppressed':False})
        other=app.dimension_analysis(self.db,sid,'genre','M')
        self.assertEqual(other['domains'][0]['capacity'],20)

    def test_dimension_analysis_multi_returns_one_result_per_value_in_order(self):
        # Regression for an ultrareview efficiency finding: comparing several
        # values used to re-run the identical unfiltered participant-matching
        # query once per value; dimension_analysis_multi() batches it into a
        # single query and must still produce the exact same per-value numbers
        # as calling dimension_analysis() once per value.
        sid='sess-dim-multi-batch'; self._mk_dimension_session(sid)
        results=app.dimension_analysis_multi(self.db,sid,'genre',['F','M'])
        self.assertEqual([r['dimension']['value'] for r in results],['F','M'])
        self.assertEqual(results[0]['completedCount'],5); self.assertEqual(results[0]['domains'][0]['capacity'],100)
        self.assertEqual(results[1]['completedCount'],5); self.assertEqual(results[1]['domains'][0]['capacity'],20)

    def test_dimension_analysis_multi_choice_matches_array_membership(self):
        schema_id,fields=self._mk_schema_with_fields()
        app.update_profile_field(self.db,fields['langues'],{'isDimension':True})
        sid='sess-dim-multi'; tid=self._mk_session(sid,profile_schema_id=schema_id)
        for i in range(5):
            pid=f'{sid}-fr{i}'; self._add_participant(sid,pid,tid,value=5,n=1)
            app.set_participant_profile_values(self.db,pid,{'langues':['fr','wo']})
        for i in range(5):
            pid=f'{sid}-en{i}'; self._add_participant(sid,pid,tid,value=1,n=1)
            app.set_participant_profile_values(self.db,pid,{'langues':['en']})
        self.db.commit()
        result=app.dimension_analysis(self.db,sid,'langues','fr')
        self.assertEqual(result['completedCount'],5)
        self.assertEqual(result['domains'][0]['capacity'],100)

    def test_dimension_analysis_suppresses_results_below_min_cohort_n(self):
        sid='sess-dim-small'; self._mk_dimension_session(sid,n_f=2,n_m=8)
        result=app.dimension_analysis(self.db,sid,'genre','F')
        self.assertEqual(result['completedCount'],2)
        self.assertTrue(result['dimension']['suppressed'])
        self.assertIsNone(result['domains'][0]['capacity'])
        self.assertIsNone(result['domains'][0]['consensus'])
        self.assertIsNone(result['global']['capacity'])
        self.assertEqual(result['domains'][0]['indicators'][0]['distribution'],{})
        # Counts stay visible even when the numbers are suppressed - needed by
        # the UI to explain *why* it's hiding the results.
        self.assertEqual(result['participantCount'],2)

    def test_dimension_analysis_no_matching_participant_returns_empty_cohort_without_crashing(self):
        sid='sess-dim-empty'; self._mk_dimension_session(sid,n_f=0,n_m=5)
        result=app.dimension_analysis(self.db,sid,'genre','F')
        self.assertEqual(result['completedCount'],0)
        self.assertEqual(result['participantCount'],0)
        self.assertTrue(result['dimension']['suppressed'])

    def test_analysis_for_participant_ids_none_is_unchanged_from_pre_lot5_behaviour(self):
        sid='sess-dim-regression'; self._mk_dimension_session(sid)
        self.assertEqual(app.analysis_for(self.db,[sid]),app.analysis_for(self.db,[sid],participant_ids=None))

    # --- Mission parite :8810->:8820 (consignes_claude.txt) : filtres
    # combinables entre PLUSIEURS dimensions differentes a la fois (par
    # opposition a dimension_analysis_multi, qui compare plusieurs valeurs
    # d'UNE seule dimension) - epc/profile.py participants_matching_filters()
    # / epc/scoring.py filtered_analysis() ---

    def _mk_combined_filter_session(self,sid):
        schema_id,fields=self._mk_schema_with_fields()
        app.update_profile_field(self.db,fields['genre'],{'isDimension':True})
        app.update_profile_field(self.db,fields['langues'],{'isDimension':True})
        tid=self._mk_session(sid,profile_schema_id=schema_id)
        # Matches BOTH filters (genre=F and langues contains fr): value 5 -> capacity 100
        for i in range(5):
            pid=f'{sid}-f-fr{i}'; self._add_participant(sid,pid,tid,value=5,n=1)
            app.set_participant_profile_values(self.db,pid,{'genre':'F','langues':['fr']})
        # Matches genre=F only (langues=en): must NOT leak into the combined cohort
        for i in range(5):
            pid=f'{sid}-f-en{i}'; self._add_participant(sid,pid,tid,value=3,n=1)
            app.set_participant_profile_values(self.db,pid,{'genre':'F','langues':['en']})
        # Matches langues=fr only (genre=M): must NOT leak into the combined cohort
        for i in range(5):
            pid=f'{sid}-m-fr{i}'; self._add_participant(sid,pid,tid,value=1,n=1)
            app.set_participant_profile_values(self.db,pid,{'genre':'M','langues':['fr']})
        self.db.commit()
        return schema_id

    def test_filtered_analysis_combines_multiple_dimensions_with_and(self):
        sid='sess-filters-and'; self._mk_combined_filter_session(sid)
        result=app.filtered_analysis(self.db,sid,{'genre':['F'],'langues':['fr']})
        self.assertEqual(result['completedCount'],5)
        self.assertEqual(result['domains'][0]['capacity'],100)
        self.assertEqual(result['filters']['applied'],[{'fieldKey':'genre','fieldLabel':'Genre','values':['F']},{'fieldKey':'langues','fieldLabel':'Langues','values':['fr']}])

    def test_filtered_analysis_or_within_one_filter_and_across_filters(self):
        # genre in [F,M] (matches everyone) AND langues=fr (only fr speakers):
        # OR is per-filter, AND is between filters.
        sid='sess-filters-or'; self._mk_combined_filter_session(sid)
        result=app.filtered_analysis(self.db,sid,{'genre':['F','M'],'langues':['fr']})
        self.assertEqual(result['completedCount'],10)

    def test_filtered_analysis_empty_filters_is_whole_session(self):
        sid='sess-filters-empty'; self._mk_combined_filter_session(sid)
        empty_filter=app.filtered_analysis(self.db,sid,{})
        whole=app.analysis(self.db,sid)
        self.assertEqual(empty_filter['completedCount'],whole['completedCount'])
        self.assertEqual(empty_filter['domains'],whole['domains'])
        self.assertEqual(empty_filter['global'],whole['global'])

    def test_filtered_analysis_no_matching_participant_returns_empty_cohort(self):
        sid='sess-filters-none'; self._mk_combined_filter_session(sid)
        result=app.filtered_analysis(self.db,sid,{'genre':['Autre']})
        self.assertEqual(result['completedCount'],0)
        self.assertIsNone(result['global']['capacity'])

    def test_filtered_analysis_n1_yields_real_capacity_with_non_calculable_consensus(self):
        # Regle explicite de consignes_claude.txt (section 5) pour ce moteur de
        # filtres combines : N=1 => capacite reelle, consensus "Non calculable"
        # (pas de suppression avant N=5 comme sur l'ecran de comparaison a une
        # seule dimension, qui garde son propre MIN_COHORT_N).
        sid='sess-filters-n1'
        schema_id,fields=self._mk_schema_with_fields()
        app.update_profile_field(self.db,fields['genre'],{'isDimension':True})
        t=self._mk_session(sid,profile_schema_id=schema_id)
        self._add_participant(sid,'p-solo',t,value=4,n=1)
        app.set_participant_profile_values(self.db,'p-solo',{'genre':'F'})
        self.db.commit()
        result=app.filtered_analysis(self.db,sid,{'genre':['F']})
        self.assertEqual(result['completedCount'],1)
        self.assertEqual(result['domains'][0]['capacity'],80)
        self.assertEqual(result['domains'][0]['consensusNote'],'single_respondent')

    def test_filtered_analysis_refuses_unflagged_dimension(self):
        sid='sess-filters-refuse'; self._mk_session(sid,profile_schema_id=self._mk_schema_with_fields()[0])
        with self.assertRaises(ValueError):
            app.filtered_analysis(self.db,sid,{'organisation':['ONG A']})

    # --- Mission parite :8810->:8820 (consignes_claude.txt) : constats
    # automatiques deterministes - epc/scoring.py objective_findings() ---

    def _add_domain_responses(self,sid,domain_index,values):
        """Inserts one completed participant per value in `values`, each
        answering every indicator of the domain at display_order `domain_index`
        (0-based) with that single value - lets a test control a domain's
        capacity/consensus precisely without touching the other domains."""
        t=self.db.execute('select template_id from sessions where id=?',(sid,)).fetchone()['template_id']
        domain=self.db.execute('select id from domains where template_id=? order by display_order limit 1 offset ?',(t,domain_index)).fetchone()['id']
        inds=[r['id'] for r in self.db.execute('select id from indicators where domain_id=? order by display_order',(domain,))]
        for i,value in enumerate(values):
            pid=f'{sid}-d{domain_index}-{i}'
            self.db.execute('insert into participants values(?,?,?,?,?,?,?)',(pid,sid,pid,'completed',app.now(),app.now(),None))
            for iid in inds:
                self.db.execute('insert into responses values(?,?,?,?,?,?,?,?)',(str(uuid.uuid4()),sid,pid,iid,str(value),'numeric',app.now(),app.now()))
        return domain

    def _mk_findings_session(self,sid):
        self._mk_session(sid)
        domains={}
        domains['force']=self._add_domain_responses(sid,0,[5,5,5,5,5])  # capacity 100, consensus 100 (sd=0)
        domains['vigilance_hi_cap_lo_cons']=self._add_domain_responses(sid,1,[5,5,5,5,1])  # capacity 84, consensus ~10.6
        domains['vigilance_lo_cap_hi_cons']=self._add_domain_responses(sid,2,[1,1,1,1,1])  # capacity 20, consensus 100
        domains['fragile_only']=self._add_domain_responses(sid,3,[1,1,1,1,3])  # capacity 28, consensus ~55.3
        self.db.commit()
        return domains

    def test_objective_findings_flags_force_with_high_consensus(self):
        sid='sess-find-forces'; domains=self._mk_findings_session(sid)
        findings=app.objective_findings(app.analysis_for(self.db,[sid]))
        force_ids=[d['id'] for d in findings['forces']['domains']]
        self.assertIn(domains['force'],force_ids)
        force=next(d for d in findings['forces']['domains'] if d['id']==domains['force'])
        self.assertEqual(force['capacity'],100)
        self.assertTrue(force['alsoHighConsensus'])
        self.assertEqual(force['responses'],5)

    def test_objective_findings_flags_fragile_domains(self):
        sid='sess-find-fragile'; domains=self._mk_findings_session(sid)
        findings=app.objective_findings(app.analysis_for(self.db,[sid]))
        fragile_ids=[d['id'] for d in findings['fragilites']['domains']]
        self.assertIn(domains['fragile_only'],fragile_ids)
        self.assertIn(domains['vigilance_lo_cap_hi_cons'],fragile_ids)

    def test_objective_findings_flags_vigilance_high_capacity_low_consensus(self):
        sid='sess-find-vig1'; domains=self._mk_findings_session(sid)
        findings=app.objective_findings(app.analysis_for(self.db,[sid]))
        entry=next(v for v in findings['vigilance'] if v['id']==domains['vigilance_hi_cap_lo_cons'])
        self.assertEqual(entry['reason'],'capacite_elevee_consensus_faible')
        self.assertAlmostEqual(entry['capacity'],84)

    def test_objective_findings_flags_vigilance_low_capacity_high_consensus(self):
        sid='sess-find-vig2'; domains=self._mk_findings_session(sid)
        findings=app.objective_findings(app.analysis_for(self.db,[sid]))
        entry=next(v for v in findings['vigilance'] if v['id']==domains['vigilance_lo_cap_hi_cons'])
        self.assertEqual(entry['reason'],'capacite_faible_consensus_eleve')
        self.assertEqual(entry['consensus'],100)

    def test_objective_findings_never_invents_a_cause_for_plain_fragile_domain(self):
        # capacite faible mais consensus moyen (< FORCE_THRESHOLD) : fragile
        # seul, ne doit PAS apparaitre en vigilance (pas de desaccord capacite/consensus).
        sid='sess-find-nofalsepositive'; domains=self._mk_findings_session(sid)
        findings=app.objective_findings(app.analysis_for(self.db,[sid]))
        vigilance_ids=[v['id'] for v in findings['vigilance']]
        self.assertNotIn(domains['fragile_only'],vigilance_ids)

    def test_objective_findings_comparison_flags_gap_between_sub_populations(self):
        sid='sess-find-gap'; self._mk_dimension_session(sid)  # F capacity 100, M capacity 20 on domain 0
        results=app.dimension_analysis_multi(self.db,sid,'genre',['F','M'])
        findings=app.objective_findings(app.analysis(self.db,sid),comparison=results)
        gap_entries=[v for v in findings['vigilance'] if v['reason']=='ecart_sous_populations']
        self.assertEqual(len(gap_entries),1)
        self.assertEqual(gap_entries[0]['gap'],80)

    def test_objective_findings_does_not_mutate_the_source_result(self):
        # Regression guard: annotating a "force" item must never leak an extra
        # key back into the plain analysis payload (result["domains"]) shown
        # elsewhere - objective_findings() must copy, never alias, domain dicts.
        sid='sess-find-nomutate'; self._mk_findings_session(sid)
        result=app.analysis_for(self.db,[sid])
        app.objective_findings(result)
        self.assertNotIn('alsoHighConsensus',result['domains'][0])

    def test_analysis_embeds_findings(self):
        sid='sess-find-embed'; domains=self._mk_findings_session(sid)
        out=app.analysis(self.db,sid)
        self.assertIn('findings',out)
        self.assertIn(domains['force'],[d['id'] for d in out['findings']['forces']['domains']])

    def test_filtered_analysis_embeds_findings(self):
        sid='sess-find-filtered'; self._mk_combined_filter_session(sid)
        result=app.filtered_analysis(self.db,sid,{'genre':['F'],'langues':['fr']})
        self.assertIn('findings',result)

    def test_indicator_level_output_carries_graded_capacity_and_consensus(self):
        # Mission de parite :8810->:8820 (consignes_claude.txt), "priorites
        # enrichies" : stable-simple graduait deja capacite/consensus au
        # niveau indicateur (pas seulement domaine) - domainDiagnostic() en
        # depend pour sa colonne "Graduation".
        sid='sess-indicator-grade'
        t=self._mk_session(sid)
        for i,v in enumerate([5]*5): self._add_participant(sid,f'p{i}',t,value=v,n=1)
        self.db.commit()
        indicator=app.analysis(self.db,sid)['domains'][0]['indicators'][0]
        self.assertEqual(indicator['capacity'],100)
        self.assertEqual(indicator['gradedCapacity'],100)
        self.assertIsNotNone(indicator['gradedConsensus'])

    def test_suppressed_cohort_nulls_indicator_graded_values_too(self):
        sid='sess-suppress-grade'; self._mk_dimension_session(sid,n_f=2,n_m=8)
        result=app.dimension_analysis(self.db,sid,'genre','F')
        self.assertIsNone(result['domains'][0]['indicators'][0]['gradedCapacity'])
        self.assertIsNone(result['domains'][0]['indicators'][0]['gradedConsensus'])

    def test_filtered_analysis_small_cohort_is_not_suppressed_from_n2_upward(self):
        # Contrairement a l'ecran de comparaison a une seule dimension (qui
        # garde sa propre suppression MIN_COHORT_N), le moteur de filtres
        # combines calcule normalement des N>=2 (regle explicite de
        # consignes_claude.txt, section 5).
        sid='sess-find-small-cohort'; self._mk_dimension_session(sid,n_f=2,n_m=8)
        result=app.filtered_analysis(self.db,sid,{'genre':['F']})
        self.assertEqual(result['completedCount'],2)
        self.assertIsNotNone(result['domains'][0]['capacity'])

    # --- Mission parite :8810->:8820 (consignes_claude.txt section 14) : test
    # de parite EPC prescrit - profil EPC par defaut + scenario Hommes/Femmes/
    # Tous avec valeurs et resultats attendus exacts ---

    def test_parity_scenario_hommes_femmes_tous(self):
        t=self.db.execute("select id from templates where name='EPC / SENEVAL'").fetchone()['id']
        sid=app.create_session(self.db,'owner-parity',{'templateId':t,'name':'Atelier parite EPC'})
        schema_id=self.db.execute('select profile_schema_id from sessions where id=?',(sid,)).fetchone()['profile_schema_id']
        self.assertIsNotNone(schema_id)
        dims=app.available_dimensions(self.db,sid)
        self.assertEqual(len(dims),5)
        # "modifiables" : un champ par defaut se modifie comme n'importe quel
        # champ de profil ordinaire (deja verifie en detail au point 5/13 ;
        # ici on verifie juste que le scenario de recette ne heurte rien).
        tid=self.db.execute('select template_id from sessions where id=?',(sid,)).fetchone()['template_id']
        def add(pid,value,sexe):
            self.db.execute('insert into participants values(?,?,?,?,?,?,?)',(pid,sid,pid,'completed',app.now(),app.now(),None))
            for iid in self._indicator_ids(tid):
                self.db.execute('insert into responses values(?,?,?,?,?,?,?,?)',(str(uuid.uuid4()),sid,pid,iid,str(value),'numeric',app.now(),app.now()))
            app.set_participant_profile_values(self.db,pid,{'sexe':sexe})
        for i,v in enumerate([5,5,4,4]): add(f'h{i}',v,'Homme')
        for i,v in enumerate([1,1,2,2]): add(f'f{i}',v,'Femme')
        self.db.commit()

        whole=app.analysis(self.db,sid)
        self.assertEqual(whole['completedCount'],8)
        self.assertEqual(whole['global']['capacity'],60)

        hommes=app.filtered_analysis(self.db,sid,{'sexe':['Homme']})
        self.assertEqual(hommes['completedCount'],4)
        self.assertEqual(hommes['global']['capacity'],90)
        self.assertAlmostEqual(hommes['global']['consensus'],71.13,places=1)
        self.assertEqual(hommes['global']['gradedCapacity'],85)
        self.assertEqual(hommes['global']['gradedConsensus'],50)

        femmes=app.filtered_analysis(self.db,sid,{'sexe':['Femme']})
        self.assertEqual(femmes['completedCount'],4)
        self.assertEqual(femmes['global']['capacity'],30)
        self.assertAlmostEqual(femmes['global']['consensus'],71.13,places=1)
        self.assertEqual(femmes['global']['gradedCapacity'],10)
        self.assertEqual(femmes['global']['gradedConsensus'],50)

        # "consensus recalcule sur les reponses individuelles" : jamais une
        # moyenne des deux consensus de sous-groupes (qui vaudrait ~71.1 aussi
        # et masquerait une vraie recomputation).
        self.assertNotAlmostEqual(whole['global']['consensus'],71.13,places=1)

        # Filtre combine : Sexe + Profil + Age (les 3 dimensions du modele par
        # defaut) - aucun participant ne porte encore Profil/Age ici, donc le
        # filtre combine doit legitimement retourner N=0, pas une erreur.
        combined=app.filtered_analysis(self.db,sid,{'sexe':['Homme'],'profil':['ONG'],'tranche-dage':['18–24']})
        self.assertEqual(combined['completedCount'],0)
        self.assertIsNone(combined['global']['capacity'])

    # --- Mission parite :8810->:8820 (consignes_claude.txt section 15) : test
    # de modularite - un champ Region cree via le moteur generique doit se
    # comporter EXACTEMENT comme un champ EPC par defaut, sans aucun code
    # specifique a "Region" nulle part dans le moteur ---

    def test_modularity_custom_region_field_behaves_like_any_dimension(self):
        t=self.db.execute("select id from templates where name='EPC / SENEVAL'").fetchone()['id']
        sid=app.create_session(self.db,'owner-region',{'templateId':t,'name':'Atelier Région'})
        schema_id=self.db.execute('select profile_schema_id from sessions where id=?',(sid,)).fetchone()['profile_schema_id']
        app.create_profile_field(self.db,schema_id,{'fieldType':'single_choice','label':'Région','options':['Dakar','Thiès','Saint-Louis'],'isDimension':True})
        self.db.commit()

        # Filtre : disponible au meme titre que les 5 dimensions EPC par defaut.
        dims=app.available_dimensions(self.db,sid)
        self.assertEqual(len(dims),6)
        region_key=next(d['fieldKey'] for d in dims if d['label']=='Région')

        tid=self.db.execute('select template_id from sessions where id=?',(sid,)).fetchone()['template_id']
        def add(pid,value,region):
            self.db.execute('insert into participants values(?,?,?,?,?,?,?)',(pid,sid,pid,'completed',app.now(),app.now(),None))
            for iid in self._indicator_ids(tid):
                self.db.execute('insert into responses values(?,?,?,?,?,?,?,?)',(str(uuid.uuid4()),sid,pid,iid,str(value),'numeric',app.now(),app.now()))
            app.set_participant_profile_values(self.db,pid,{region_key:region})
        for i,v in enumerate([5,5,5]): add(f'dk{i}',v,'Dakar')
        for i,v in enumerate([1,1,1]): add(f'th{i}',v,'Thiès')
        self.db.commit()

        # Filtre (une seule valeur, via le moteur combinable).
        dakar=app.filtered_analysis(self.db,sid,{region_key:['Dakar']})
        self.assertEqual(dakar['completedCount'],3)
        self.assertEqual(dakar['global']['capacity'],100)

        # Comparaison (plusieurs valeurs de la meme dimension, ecran existant) :
        # cet ecran garde son propre seuil MIN_COHORT_N (3 < 5 ici), donc les
        # nombres sont masques - c'est le comportement generique attendu pour
        # N'IMPORTE QUELLE dimension, pas une specificite de Région.
        comparison=app.dimension_analysis_multi(self.db,sid,region_key,['Dakar','Thiès'])
        self.assertEqual([r['dimension']['value'] for r in comparison],['Dakar','Thiès'])
        self.assertEqual(comparison[0]['completedCount'],3)
        self.assertTrue(comparison[0]['dimension']['suppressed'])

        # Export filtre : ne plante pas et reflete le filtre applique.
        header,domain_rows,_=app.filtered_analysis_rows(self.db,sid,{region_key:['Dakar']})
        filt_desc=[h[1] for h in header if h[0]=='Filtres actifs'][0]
        self.assertIn('Région',filt_desc)
        self.assertEqual(domain_rows[0][1],100)

        # Rapport : la repartition par region apparait dans le profil agrege.
        breakdown=app.report_data(self.db,sid)['profile']
        self.assertIn(['Région','Dakar',3],breakdown)

    # --- Lot 7 (modularisation, AUDIT_MODULARISATION_8800.md) : migration V1
    # (analysis_notes/recommendations) -> V2 (analysis_entries/
    # workshop_recommendations) - epc/qualitatif.py
    # migrate_legacy_qualitative_data(), jamais appelee automatiquement ---

    def test_migrate_creates_analysis_entry_and_auto_priority_when_none_exists(self):
        sid='sess-migrate-1'; self._mk_session(sid)
        domain=self.db.execute('select id from domains where display_order=1').fetchone()['id']
        indicator=self.db.execute('select id from indicators where domain_id=? order by display_order limit 1',(domain,)).fetchone()['id']
        app.create_analysis_note(self.db,sid,{'indicatorId':indicator,'kind':'Cause racine potentielle','content':'Constat V1','validationStatus':'FAIT_VALIDE'})
        self.assertIsNone(self.db.execute('select id from priorities where session_id=? and indicator_id=?',(sid,indicator)).fetchone())
        result=app.migrate_legacy_qualitative_data(self.db,session_id=sid)
        self.assertEqual(result['migratedEntries'],1)
        self.assertEqual(result['skippedEntries'],0)
        priority=self.db.execute('select id from priorities where session_id=? and indicator_id=?',(sid,indicator)).fetchone()
        self.assertIsNotNone(priority)
        entry=self.db.execute('select * from analysis_entries where session_id=?',(sid,)).fetchone()
        self.assertEqual(entry['priority_id'],priority['id'])
        self.assertEqual(entry['kind'],'cause')
        self.assertEqual(entry['item_type'],'Cause racine potentielle')
        self.assertEqual(entry['content'],'Constat V1')
        self.assertEqual(entry['validation_status'],'RETENU')

    def test_migrate_reuses_existing_priority_for_the_same_indicator(self):
        sid='sess-migrate-2'; self._mk_session(sid)
        domain=self.db.execute('select id from domains where display_order=1').fetchone()['id']
        indicator=self.db.execute('select id from indicators where domain_id=? order by display_order limit 1',(domain,)).fetchone()['id']
        app.toggle_priority(self.db,sid,{'domainId':domain,'indicatorId':indicator,'votes':1})
        existing_priority=self.db.execute('select id from priorities where session_id=? and indicator_id=?',(sid,indicator)).fetchone()['id']
        app.create_analysis_note(self.db,sid,{'indicatorId':indicator,'kind':'Symptôme','content':'Deuxieme constat','validationStatus':'HYPOTHESE'})
        app.migrate_legacy_qualitative_data(self.db,session_id=sid)
        self.assertEqual(self.db.execute('select count(*) from priorities where session_id=?',(sid,)).fetchone()[0],1)
        entry=self.db.execute('select * from analysis_entries where session_id=?',(sid,)).fetchone()
        self.assertEqual(entry['priority_id'],existing_priority)
        self.assertEqual(entry['validation_status'],'A_DISCUTER')

    def test_migrate_recommendation_maps_category_and_folds_lever_into_description(self):
        sid='sess-migrate-3'; self._mk_session(sid)
        domain=self.db.execute('select id from domains where display_order=1').fetchone()['id']
        indicator=self.db.execute('select id from indicators where domain_id=? order by display_order limit 1',(domain,)).fetchone()['id']
        app.create_legacy_recommendation(self.db,sid,{'indicatorId':indicator,'title':'Reco V1','description':'Description V1','lever':'Levier libre V1','kind':'formation','owner':'Awa','horizon':'2026'})
        result=app.migrate_legacy_qualitative_data(self.db,session_id=sid)
        self.assertEqual(result['migratedRecommendations'],1)
        rec=self.db.execute('select * from workshop_recommendations where session_id=?',(sid,)).fetchone()
        self.assertEqual(rec['title'],'Reco V1')
        self.assertEqual(rec['category'],'Formation')
        self.assertEqual(rec['owner'],'Awa')
        self.assertEqual(rec['status'],'Proposée')
        self.assertIn('Description V1',rec['description'])
        self.assertIn('Levier (V1) : Levier libre V1',rec['description'])
        self.assertIsNotNone(self.db.execute('select id from priorities where session_id=? and indicator_id=?',(sid,indicator)).fetchone())

    def test_migrate_is_idempotent(self):
        sid='sess-migrate-4'; self._mk_session(sid)
        domain=self.db.execute('select id from domains where display_order=1').fetchone()['id']
        indicator=self.db.execute('select id from indicators where domain_id=? order by display_order limit 1',(domain,)).fetchone()['id']
        app.create_analysis_note(self.db,sid,{'indicatorId':indicator,'kind':'Cause','content':'Constat unique','validationStatus':'FAIT_VALIDE'})
        app.create_legacy_recommendation(self.db,sid,{'indicatorId':indicator,'title':'Reco unique','description':'Desc unique','kind':'organisation','lever':'Levier test'})
        first=app.migrate_legacy_qualitative_data(self.db,session_id=sid)
        second=app.migrate_legacy_qualitative_data(self.db,session_id=sid)
        self.assertEqual(first['migratedEntries'],1); self.assertEqual(first['migratedRecommendations'],1)
        self.assertEqual(second['migratedEntries'],0); self.assertEqual(second['migratedRecommendations'],0)
        self.assertEqual(self.db.execute('select count(*) from analysis_entries where session_id=?',(sid,)).fetchone()[0],1)
        self.assertEqual(self.db.execute('select count(*) from workshop_recommendations where session_id=?',(sid,)).fetchone()[0],1)

    def test_migrate_skips_notes_without_an_indicator(self):
        sid='sess-migrate-5'; self._mk_session(sid)
        app.create_analysis_note(self.db,sid,{'kind':'Cause','content':'Constat orphelin'})
        result=app.migrate_legacy_qualitative_data(self.db,session_id=sid)
        self.assertEqual(result['migratedEntries'],0); self.assertEqual(result['skippedEntries'],1)

    def test_migrate_never_touches_v1_rows(self):
        sid='sess-migrate-6'; self._mk_session(sid)
        domain=self.db.execute('select id from domains where display_order=1').fetchone()['id']
        indicator=self.db.execute('select id from indicators where domain_id=? order by display_order limit 1',(domain,)).fetchone()['id']
        app.create_analysis_note(self.db,sid,{'indicatorId':indicator,'kind':'Cause','content':'Constat preserve'})
        app.migrate_legacy_qualitative_data(self.db,session_id=sid)
        self.assertEqual(self.db.execute('select count(*) from analysis_notes where session_id=?',(sid,)).fetchone()[0],1)

    def test_migrate_session_scoping_leaves_other_sessions_untouched(self):
        sid1='sess-migrate-scope-1'; sid2='sess-migrate-scope-2'
        self._mk_session(sid1); self._mk_session(sid2)
        domain=self.db.execute('select id from domains where display_order=1').fetchone()['id']
        indicator=self.db.execute('select id from indicators where domain_id=? order by display_order limit 1',(domain,)).fetchone()['id']
        app.create_analysis_note(self.db,sid1,{'indicatorId':indicator,'kind':'Cause','content':'Session 1'})
        app.create_analysis_note(self.db,sid2,{'indicatorId':indicator,'kind':'Cause','content':'Session 2'})
        result=app.migrate_legacy_qualitative_data(self.db,session_id=sid1)
        self.assertEqual(result['migratedEntries'],1)
        self.assertEqual(self.db.execute('select count(*) from analysis_entries where session_id=?',(sid1,)).fetchone()[0],1)
        self.assertEqual(self.db.execute('select count(*) from analysis_entries where session_id=?',(sid2,)).fetchone()[0],0)

    # --- Lot 6 (modularisation, AUDIT_MODULARISATION_8800.md) : manifeste de
    # restitution - un seul modele (epc_seneval) existe aujourd'hui, son
    # manifeste doit lister exactement ce qui est deja rendu (aucun
    # changement de comportement), tout en etant le point que les generateurs
    # XLSX/DOCX et les routes IA consultent reellement - epc/restitution.py ---

    def test_resolve_model_key_defaults_untagged_template_to_epc_seneval(self):
        # A custom/blank questionnaire (model_key NULL, cf. lot 2) still runs
        # on the single EPC scoring engine that exists today - it must get
        # the same restitution manifest as the canonical model, not a crash
        # or an empty one.
        self.assertEqual(app.resolve_model_key({'model_key':None}),'epc_seneval')
        self.assertEqual(app.resolve_model_key({'model_key':'epc_seneval'}),'epc_seneval')

    def test_restitution_manifest_lists_every_current_report_section_for_epc_seneval(self):
        manifest=app.restitution_manifest({'model_key':'epc_seneval'})
        self.assertEqual(manifest['modelKey'],'epc_seneval')
        self.assertEqual(manifest['reportSections'],['synthese','profil_participants','domaines','indicateurs','constats','priorites','analyses','causes','consequences','leviers','recommandations','formations','plan_action','questionnaire'])
        self.assertEqual(len(manifest['aiReportSections']),9)
        self.assertEqual(set(manifest['aiReportSections']),set(manifest['aiSectionLabels']))
        self.assertIn('EPC/SENEVAL',manifest['aiSystemPrompt'])
        # Mission de parite :8810->:8820 (consignes_claude.txt) : restaure
        # l'action IA "Proposer une rédaction de synthèse" sur la Synthèse finale.
        self.assertIn('synthese_finale',manifest['aiReportSections'])
        self.assertEqual(manifest['aiSectionLabels']['synthese_finale'],'Synthèse finale')

    def test_ai_report_context_includes_aggregate_profile_and_findings(self):
        sid='sess-ai-context'
        schema_id,fields=self._mk_schema_with_fields()
        t=self._mk_session(sid,profile_schema_id=schema_id)
        for i,v in enumerate([5]*5): self._add_participant(sid,f'p{i}',t,value=v)
        app.set_participant_profile_values(self.db,'p0',{'genre':'F'})
        self.db.commit()
        context=app.ai_report_context(self.db,sid)
        self.assertIn('Profil agrégé des participants',context)
        self.assertIn('Genre — F',context)
        self.assertIn('Constats automatiques',context)
        self.assertIn('Force :',context)

    def test_restitution_manifest_falls_back_to_epc_seneval_for_an_unknown_model_key(self):
        # Defensive default (no second model exists yet to actually hit this
        # branch) - documented in restitution_manifest()'s own docstring.
        self.assertEqual(app.restitution_manifest({'model_key':'some_future_model'}),app.restitution_manifest({'model_key':'epc_seneval'}))

    def test_report_data_includes_the_session_restitution_manifest(self):
        sid='sess-restitution-manifest'; self._mk_session(sid)
        pid='p'; self._add_participant(sid,pid,self.db.execute('select id from templates').fetchone()['id'],value=4,n=1)
        self.db.commit()
        data=app.report_data(self.db,sid)
        self.assertEqual(data['manifest']['modelKey'],'epc_seneval')
        self.assertIn('domaines',data['manifest']['reportSections'])

    # --- Mission parite :8810->:8820 (consignes_claude.txt) : profil des
    # participants + constats automatiques dans le rapport final (JSON/XLSX/DOCX)
    # - epc/profile.py participant_profile_breakdown(), epc/restitution.py
    # findings_rows() ---

    def test_report_data_includes_participant_profile_breakdown(self):
        sid='sess-report-profile'
        schema_id,fields=self._mk_schema_with_fields()
        self._mk_session(sid,profile_schema_id=schema_id)
        t=self.db.execute('select template_id from sessions where id=?',(sid,)).fetchone()['template_id']
        self._add_participant(sid,'p1',t,value=5,n=1)
        app.set_participant_profile_values(self.db,'p1',{'genre':'F'})
        self.db.commit()
        data=app.report_data(self.db,sid)
        self.assertIn(['Genre','F',1],data['profile'])

    def test_report_data_profile_empty_without_a_schema(self):
        sid='sess-report-profile-none'; self._mk_session(sid)
        self.assertEqual(app.report_data(self.db,sid)['profile'],[])

    def test_individual_responses_rows_has_one_row_per_completed_participant_with_dynamic_profile(self):
        sid='sess-individual-responses'
        schema_id,fields=self._mk_schema_with_fields()
        t=self._mk_session(sid,profile_schema_id=schema_id)
        self._add_participant(sid,'p-done',t,value=5,n=2,status='completed')
        app.set_participant_profile_values(self.db,'p-done',{'genre':'F','age':30})
        self._add_participant(sid,'p-progress',t,value=3,n=1,status='in_progress')
        self.db.commit()
        data=app.individual_responses_rows(self.db,sid)
        self.assertEqual(len(data['rows']),1)
        row=data['rows'][0]
        self.assertEqual(row['status'],'Anonyme')
        self.assertEqual(row[fields['genre']],'F')
        self.assertEqual(row[fields['age']],30)
        self.assertEqual(len(data['profileFields']),4)

    def test_individual_responses_rows_marks_nominative_when_display_name_set(self):
        sid='sess-individual-nominative'; t=self._mk_session(sid)
        self._add_participant(sid,'p-named',t,value=5,n=1)
        self.db.execute("UPDATE participants SET display_name='Awa Diop' WHERE id='p-named'"); self.db.commit()
        row=app.individual_responses_rows(self.db,sid)['rows'][0]
        self.assertEqual(row['status'],'Nominatif')
        self.assertEqual(row['name'],'Awa Diop')

    def test_individual_responses_xlsx_and_csv_do_not_crash(self):
        sid='sess-individual-export'; t=self._mk_session(sid)
        self._add_participant(sid,'p1',t,value=5,n=1)
        self.db.commit()
        self.assertGreater(len(app.individual_responses_xlsx(self.db,sid)),500)
        self.assertGreater(len(app.individual_responses_csv(self.db,sid)),20)

    def test_filtered_analysis_export_matches_the_filtered_result_exactly(self):
        sid='sess-filtered-export'; self._mk_combined_filter_session(sid)
        header,domain_rows,indicator_rows=app.filtered_analysis_rows(self.db,sid,{'genre':['F'],'langues':['fr']})
        self.assertIn(['N (validés)',5],header)
        self.assertEqual(domain_rows[0][1],100)

    def test_filtered_analysis_export_includes_campaign_and_group_when_present(self):
        uid=self._mk_user('u-filtered-export')
        t=self.db.execute("select id,version from templates where name='EPC / SENEVAL'").fetchone()
        cid=app.create_campaign(self.db,uid,t['id'],t['version'],{'name':'Campagne export'})
        camp=self.db.execute('select * from campaigns where id=?',(cid,)).fetchone()
        g=app.create_group(self.db,camp,uid,{'name':'Groupe export'})
        header,_,_=app.filtered_analysis_rows(self.db,g['id'],{})
        self.assertIn(['Campagne',cid],header)

    def test_filtered_analysis_xlsx_and_csv_do_not_crash(self):
        sid='sess-filtered-export-bytes'; self._mk_combined_filter_session(sid)
        self.assertGreater(len(app.filtered_analysis_xlsx(self.db,sid,{'genre':['F']})),500)
        self.assertGreater(len(app.filtered_analysis_csv(self.db,sid,{'genre':['F']})),20)

    def test_filtered_analysis_export_raises_for_unknown_session(self):
        with self.assertRaises(ValueError):
            app.filtered_analysis_xlsx(self.db,'no-such-session',{})
        with self.assertRaises(ValueError):
            app.filtered_analysis_csv(self.db,'no-such-session',{})

    def test_report_xlsx_and_docx_include_profile_and_findings_sections(self):
        sid='sess-report-findings'
        schema_id,fields=self._mk_schema_with_fields()
        self._mk_session(sid,profile_schema_id=schema_id)
        self._add_domain_responses(sid,0,[5,5,5,5,5])
        self.db.commit()
        app.set_participant_profile_values(self.db,f'{sid}-d0-0',{'genre':'F'})
        xlsx=app.report_xlsx(self.db,sid)
        docx=app.report_docx(self.db,sid)
        self.assertGreater(len(xlsx),1000)
        self.assertGreater(len(docx),1000)

    def test_session_restitution_manifest_matches_restitution_manifest_for_the_sessions_template(self):
        sid='sess-restitution-manifest-2'; self._mk_session(sid)
        self.assertEqual(app.session_restitution_manifest(self.db,sid),app.restitution_manifest({'model_key':'epc_seneval'}))

if __name__=='__main__': unittest.main()
