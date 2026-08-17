import json
import tempfile
import uuid
import unittest
from pathlib import Path
import app

class EngineTests(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory(); self.db=app.connect(Path(self.tmp.name)/'test.sqlite3'); app.init_db(self.db)
    def tearDown(self): self.db.close(); self.tmp.cleanup()
    def _mk_session(self,sid,name='test',template=None):
        t=template or self.db.execute('select id,version from templates').fetchone()
        self.db.execute("insert into sessions values(?,?,?,?,?,?,?,?,?,?,?,?,?)",(sid,t['id'],t['version'],name,'','','', 'open',app.now(),None,'',None,None))
        self.db.commit()
    def _mk_participant(self,sid,pid,status='in_progress'):
        self.db.execute('insert into participants values(?,?,?,?,?,?,?)',(pid,sid,pid,status,app.now(),app.now() if status=='completed' else None,None))
        self.db.commit()
    def test_epc_seed_has_seven_domains_and_seventy_indicators(self):
        t=self.db.execute('select id from templates').fetchone()['id']; payload=app.template_payload(self.db,t)
        self.assertEqual(len(payload['domains']),7); self.assertEqual(sum(len(d['indicators']) for d in payload['domains']),70)
    def test_grade_and_analysis_keep_raw_responses(self):
        t=self.db.execute('select id,version from templates').fetchone(); sid='session'; self.db.execute("insert into sessions values(?,?,?,?,?,?,?,?,?,?,?,?,?)",(sid,t['id'],t['version'],'test','','','', 'open',app.now(),None,'',None,None))
        domain=self.db.execute('select id from domains where display_order=1').fetchone()['id']; inds=self.db.execute('select id from indicators where domain_id=? order by display_order limit 1',(domain,)).fetchone()['id']
        for n,v in [('a',1),('b',5)]:
            pid=n; self.db.execute('insert into participants values(?,?,?,?,?,?,?)',(pid,sid,n,'completed',app.now(),app.now(),None)); self.db.execute('insert into responses values(?,?,?,?,?,?,?,?)',(str(uuid.uuid4()),sid,pid,inds,str(v),'numeric',app.now(),app.now()))
        self.db.commit(); out=app.analysis(self.db,sid); indicator=out['domains'][0]['indicators'][0]
        self.assertEqual(indicator['responses'],2); self.assertEqual(indicator['capacity'],60); self.assertEqual(indicator['consensus'],0); self.assertEqual(app.grade(63,app.GRADING),40)
    def test_single_respondent_consensus_is_not_calculable(self):
        t=self.db.execute('select id,version from templates').fetchone(); sid='solo-session'; self.db.execute("insert into sessions values(?,?,?,?,?,?,?,?,?,?,?,?,?)",(sid,t['id'],t['version'],'test','','','', 'open',app.now(),None,'',None,None))
        domain=self.db.execute('select id from domains where display_order=1').fetchone()['id']; inds=self.db.execute('select id from indicators where domain_id=? order by display_order limit 1',(domain,)).fetchone()['id']
        self.db.execute('insert into participants values(?,?,?,?,?,?,?)',('solo',sid,'solo','completed',app.now(),app.now(),None)); self.db.execute('insert into responses values(?,?,?,?,?,?,?,?)',(str(uuid.uuid4()),sid,'solo',inds,'4','numeric',app.now(),app.now()))
        self.db.commit(); out=app.analysis(self.db,sid); indicator=out['domains'][0]['indicators'][0]; domain_out=out['domains'][0]
        self.assertIsNotNone(indicator['capacity']); self.assertIsNone(indicator['consensus']); self.assertEqual(indicator['consensusNote'],'single_respondent')
        self.assertIsNone(domain_out['consensus']); self.assertEqual(domain_out['consensusNote'],'single_respondent')
        self.assertIsNone(out['global']['consensus']); self.assertEqual(out['global']['consensusNote'],'single_respondent')
    def test_qualitative_chain_is_persistent_and_exported(self):
        t=self.db.execute('select id,version from templates').fetchone(); sid='qualitative-session'
        self.db.execute("insert into sessions values(?,?,?,?,?,?,?,?,?,?,?,?,?)",(sid,t['id'],t['version'],'test','','','', 'open',app.now(),None,'',None,None))
        domain=self.db.execute('select id from domains where display_order=1').fetchone()['id']; indicator=self.db.execute('select id from indicators where domain_id=? limit 1',(domain,)).fetchone()['id']
        self.db.execute('insert into priorities values(?,?,?,?,?,?)',('priority',sid,domain,indicator,1,app.now()))
        self.db.execute('insert into priority_analyses values(?,?,?,?,?,?)',('analysis',sid,'priority','Constat',app.now(),app.now()))
        self.db.execute('insert into analysis_entries values(?,?,?,?,?,?,?,?,?,?,?)',('cause',sid,'priority',None,'cause','Cause','Cause','', 'RETENU',app.now(),app.now()))
        self.db.execute('insert into workshop_recommendations values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)',('recommendation',sid,'priority','cause',None,'Action','Description','Formation','Haute','Responsable','Demain','', 'Retenue',app.now(),app.now()))
        self.db.execute('insert into training_topics values(?,?,?,?,?,?,?,?,?,?,?)',('training',sid,'priority','recommendation','Thème','Besoin','Public','Haute','',app.now(),app.now())); self.db.commit()
        data=app.qualitative_data(self.db,sid)
        self.assertEqual(len(data['priorities']),1); self.assertEqual(len(data['entries']),1); self.assertEqual(data['recommendations'][0]['status'],'Retenue')
        self.assertGreater(len(app.report_xlsx(self.db,sid)),1000)

    # --- Régression : versionnement du questionnaire par mission (recette 2026-08-17) ---

    def test_is_canonical_flag_identifies_only_the_true_reference(self):
        t = self.db.execute('select id from templates').fetchone()['id']
        self.assertTrue(app.is_canonical_template(self.db, t))
        fork = app.clone_template(self.db, t, status="session_locked")
        self.assertFalse(app.is_canonical_template(self.db, fork))

    def test_canonical_reference_structural_edit_is_refused(self):
        t = self.db.execute('select id from templates').fetchone()['id']
        with self.assertRaises(app.StructuralEditForbiddenError):
            app.guard_structural_edit(self.db, t, "add")
        with self.assertRaises(app.StructuralEditForbiddenError):
            app.guard_structural_edit(self.db, t, "delete")
        payload = app.template_payload(self.db, t)
        self.assertEqual(len(payload['domains']), 7)
        self.assertEqual(sum(len(d['indicators']) for d in payload['domains']), 70)

    def test_private_fork_is_freely_editable(self):
        t = self.db.execute('select id from templates').fetchone()['id']
        fork = app.clone_template(self.db, t, status="session_locked")
        app.guard_structural_edit(self.db, fork, "add")  # must not raise
        did = str(uuid.uuid4())
        self.db.execute('insert into domains values(?,?,?,?,?,?,?)', (did, fork, 'nouveau', 'Nouveau domaine', '', 8, 1))
        self.db.commit()
        self.assertEqual(self.db.execute('select count(*) from domains where template_id=?', (fork,)).fetchone()[0], 8)

    def test_ensure_private_template_forks_and_remaps_historical_responses(self):
        db = self.db
        t = db.execute('select id,version from templates').fetchone()
        sid_a, sid_b = 'mission-a', 'mission-b'
        self._mk_session(sid_a, template=t); self._mk_session(sid_b, template=t)
        domain = db.execute('select id from domains where display_order=1').fetchone()['id']
        indicator = db.execute('select id from indicators where domain_id=? order by display_order limit 1', (domain,)).fetchone()['id']
        pid_a = 'participant-a'
        self._mk_participant(sid_a, pid_a, status='completed')
        db.execute('insert into responses values(?,?,?,?,?,?,?,?)', (str(uuid.uuid4()), sid_a, pid_a, indicator, '4', 'numeric', app.now(), app.now()))
        db.commit()
        # both missions still share the canonical template before forking
        self.assertEqual(db.execute('select count(*) from sessions where template_id=?', (t['id'],)).fetchone()[0], 2)

        new_tid = app.ensure_private_template(db, sid_a)
        self.assertNotEqual(new_tid, t['id'])
        self.assertFalse(app.is_canonical_template(db, new_tid))
        # mission B is untouched: still on the canonical template
        self.assertEqual(db.execute('select template_id from sessions where id=?', (sid_b,)).fetchone()['template_id'], t['id'])
        # mission A's historical response now points at an indicator that exists in its NEW template
        new_indicator_id = db.execute('select indicator_id from responses where session_id=?', (sid_a,)).fetchone()['indicator_id']
        self.assertNotEqual(new_indicator_id, indicator)
        payload = app.template_payload(db, new_tid)
        all_new_ids = {i['id'] for d in payload['domains'] for i in d['indicators']}
        self.assertIn(new_indicator_id, all_new_ids)
        # analysis() still finds the response (capacity is computable, not blank)
        out = app.analysis(db, sid_a)
        self.assertEqual(out['domains'][0]['responses'], 1)
        self.assertIsNotNone(out['domains'][0]['indicators'][0]['capacity'])
        # idempotent
        self.assertEqual(app.ensure_private_template(db, sid_a), new_tid)

    def test_responses_survive_structural_edit_from_another_mission(self):
        db = self.db
        t = db.execute('select id,version from templates').fetchone()
        sid_a, sid_b = 'mission-a', 'mission-b'
        self._mk_session(sid_a, template=t); self._mk_session(sid_b, template=t)
        pid_a = 'participant-a'
        self._mk_participant(sid_a, pid_a, status='completed')
        app.ensure_private_template(db, sid_a)
        tid_a = db.execute('select template_id from sessions where id=?', (sid_a,)).fetchone()['template_id']
        domain_a = db.execute('select id from domains where template_id=? order by display_order limit 1', (tid_a,)).fetchone()['id']
        indicator_a = db.execute('select id from indicators where domain_id=? order by display_order limit 1', (domain_a,)).fetchone()['id']
        db.execute('insert into responses values(?,?,?,?,?,?,?,?)', (str(uuid.uuid4()), sid_a, pid_a, indicator_a, '5', 'numeric', app.now(), app.now()))
        db.commit()

        pid_b = 'participant-b'
        self._mk_participant(sid_b, pid_b, status='in_progress')
        app.ensure_private_template(db, sid_b)
        tid_b = db.execute('select template_id from sessions where id=?', (sid_b,)).fetchone()['template_id']
        new_did = str(uuid.uuid4())
        db.execute('insert into domains values(?,?,?,?,?,?,?)', (new_did, tid_b, 'extra', 'Domaine ajouté par B', '', 99, 1))
        db.commit()

        payload_a = app.template_payload(db, tid_a)
        self.assertEqual(len(payload_a['domains']), 7)
        self.assertEqual(sum(len(d['indicators']) for d in payload_a['domains']), 70)
        self.assertIsNotNone(db.execute('select 1 from responses where session_id=?', (sid_a,)).fetchone())
        out = app.analysis(db, sid_a)
        self.assertEqual(out['domains'][0]['responses'], 1)

        canonical = db.execute('select id from templates where is_canonical=1').fetchone()['id']
        canonical_payload = app.template_payload(db, canonical)
        self.assertEqual(len(canonical_payload['domains']), 7)
        self.assertEqual(sum(len(d['indicators']) for d in canonical_payload['domains']), 70)

    def test_five_missions_keep_isolated_questionnaire_versions_after_independent_edits(self):
        db = self.db
        t = db.execute('select id,version from templates').fetchone()
        missions = []
        for n in range(5):
            sid = f'mission-{n}'
            self._mk_session(sid, name=f'Mission {n}', template=t)
            pid = f'participant-{n}'
            self._mk_participant(sid, pid, status='in_progress')
            app.ensure_private_template(db, sid)
            tid = db.execute('select template_id from sessions where id=?', (sid,)).fetchone()['template_id']
            missions.append((sid, pid, tid))

        tids = [m[2] for m in missions]
        self.assertEqual(len(set(tids)), 5)
        for tid in tids:
            self.assertNotEqual(tid, t['id'])
            self.assertFalse(app.is_canonical_template(db, tid))

        for n, (sid, pid, tid) in enumerate(missions):
            domain = db.execute('select id from domains where template_id=? order by display_order limit 1', (tid,)).fetchone()['id']
            new_did = str(uuid.uuid4())
            db.execute('insert into domains values(?,?,?,?,?,?,?)', (new_did, tid, f'extra-{n}', f'Domaine ajouté {n}', '', 99, 1))
            indicator = db.execute('select id from indicators where domain_id=? limit 1', (domain,)).fetchone()['id']
            db.execute('insert into responses values(?,?,?,?,?,?,?,?)', (str(uuid.uuid4()), sid, pid, indicator, str(n + 1), 'numeric', app.now(), app.now()))
            db.commit()

        for n, (sid, pid, tid) in enumerate(missions):
            payload = app.template_payload(db, tid)
            codes = [d['code'] for d in payload['domains']]
            self.assertIn(f'extra-{n}', codes)
            for other in range(5):
                if other != n:
                    self.assertNotIn(f'extra-{other}', codes)
            val = db.execute('select value_json from responses where session_id=?', (sid,)).fetchone()['value_json']
            self.assertEqual(json.loads(val), n + 1)

        canonical_payload = app.template_payload(db, t['id'])
        self.assertEqual(len(canonical_payload['domains']), 7)
        self.assertEqual(sum(len(d['indicators']) for d in canonical_payload['domains']), 70)

    def test_init_schema_does_not_rerun_destructive_migration_per_request(self):
        calls = {"n": 0}
        original = app.migrate_reference_questionnaire
        app.migrate_reference_questionnaire = lambda db: calls.__setitem__("n", calls["n"] + 1)
        try:
            app.init_schema(self.db)
            app.init_schema(self.db)
        finally:
            app.migrate_reference_questionnaire = original
        self.assertEqual(calls["n"], 0)

    def test_capacity_excludes_incomplete_participants(self):
        db = self.db
        sid = 'partial-session'
        self._mk_session(sid)
        domain = db.execute('select id from domains where display_order=1').fetchone()['id']
        inds = [r['id'] for r in db.execute('select id from indicators where domain_id=? order by display_order', (domain,))]
        for n, v in [('p1', 5), ('p2', 5), ('p3', 4), ('p4', 4)]:
            self._mk_participant(sid, n, status='completed')
            db.execute('insert into responses values(?,?,?,?,?,?,?,?)', (str(uuid.uuid4()), sid, n, inds[0], str(v), 'numeric', app.now(), app.now()))
        self._mk_participant(sid, 'abandoned1', status='in_progress')
        self._mk_participant(sid, 'abandoned2', status='in_progress')
        db.execute('insert into responses values(?,?,?,?,?,?,?,?)', (str(uuid.uuid4()), sid, 'abandoned2', inds[0], '5', 'numeric', app.now(), app.now()))
        db.commit()
        out = app.analysis(db, sid)
        indicator = out['domains'][0]['indicators'][0]
        self.assertEqual(out['participantCount'], 6)
        self.assertEqual(out['completedCount'], 4)
        self.assertEqual(indicator['responses'], 4)
        self.assertEqual(indicator['capacity'], 90)

if __name__=='__main__': unittest.main()
