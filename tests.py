import tempfile
import uuid
import unittest
from pathlib import Path
import app

class EngineTests(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory(); self.db=app.connect(Path(self.tmp.name)/'test.sqlite3'); app.init_db(self.db)
    def tearDown(self): self.db.close(); self.tmp.cleanup()
    def _mk_session(self,sid,name='test',campaign_id=None,group_code=None,expected=None):
        t=self.db.execute('select id,version from templates').fetchone()
        self.db.execute("insert into sessions values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",(sid,t['id'],t['version'],name,'','','', 'open',app.now(),None,'',expected,None,campaign_id,group_code,None,None,None))
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
        self.db.commit()
        out=app.analysis(self.db,'partial-session')
        self.assertEqual(out['participantCount'],6)
        self.assertEqual(out['completedCount'],4)
        self.assertEqual(out['global']['capacity'],90)
        self.assertAlmostEqual(out['global']['consensus'],71.13,places=1)
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
if __name__=='__main__': unittest.main()
