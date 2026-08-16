import tempfile
import uuid
import unittest
from pathlib import Path
import app

class EngineTests(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory(); self.db=app.connect(Path(self.tmp.name)/'test.sqlite3'); app.init_db(self.db)
    def tearDown(self): self.db.close(); self.tmp.cleanup()
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
if __name__=='__main__': unittest.main()
