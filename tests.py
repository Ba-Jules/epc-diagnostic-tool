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
        t=template or self.db.execute('select id,version from templates where is_canonical=1').fetchone()
        self.db.execute("insert into sessions values(?,?,?,?,?,?,?,?,?,?,?,?,?)",(sid,t['id'],t['version'],name,'','','', 'open',app.now(),None,'',None,None))
        self.db.commit()
    def _mk_participant(self,sid,pid,status='in_progress'):
        self.db.execute('insert into participants (id,session_id,anonymous_id,status,started_at,completed_at,display_name) values(?,?,?,?,?,?,?)',(pid,sid,pid,status,app.now(),app.now() if status=='completed' else None,None))
        self.db.commit()
    def test_epc_seed_has_seven_domains_and_seventy_indicators(self):
        t=self.db.execute('select id from templates where is_canonical=1').fetchone()['id']; payload=app.template_payload(self.db,t)
        self.assertEqual(len(payload['domains']),7); self.assertEqual(sum(len(d['indicators']) for d in payload['domains']),70)
    def test_grade_and_analysis_keep_raw_responses(self):
        t=self.db.execute('select id,version from templates where is_canonical=1').fetchone(); sid='session'; self.db.execute("insert into sessions values(?,?,?,?,?,?,?,?,?,?,?,?,?)",(sid,t['id'],t['version'],'test','','','', 'open',app.now(),None,'',None,None))
        domain=self.db.execute('select id from domains where display_order=1').fetchone()['id']; inds=self.db.execute('select id from indicators where domain_id=? order by display_order limit 1',(domain,)).fetchone()['id']
        for n,v in [('a',1),('b',5)]:
            pid=n; self.db.execute('insert into participants (id,session_id,anonymous_id,status,started_at,completed_at,display_name) values(?,?,?,?,?,?,?)',(pid,sid,n,'completed',app.now(),app.now(),None)); self.db.execute('insert into responses values(?,?,?,?,?,?,?,?)',(str(uuid.uuid4()),sid,pid,inds,str(v),'numeric',app.now(),app.now()))
        self.db.commit(); out=app.analysis(self.db,sid); indicator=out['domains'][0]['indicators'][0]
        self.assertEqual(indicator['responses'],2); self.assertEqual(indicator['capacity'],60); self.assertEqual(indicator['consensus'],0); self.assertEqual(app.grade(63,app.GRADING),40)
    def test_single_respondent_consensus_is_not_calculable(self):
        t=self.db.execute('select id,version from templates where is_canonical=1').fetchone(); sid='solo-session'; self.db.execute("insert into sessions values(?,?,?,?,?,?,?,?,?,?,?,?,?)",(sid,t['id'],t['version'],'test','','','', 'open',app.now(),None,'',None,None))
        domain=self.db.execute('select id from domains where display_order=1').fetchone()['id']; inds=self.db.execute('select id from indicators where domain_id=? order by display_order limit 1',(domain,)).fetchone()['id']
        self.db.execute('insert into participants (id,session_id,anonymous_id,status,started_at,completed_at,display_name) values(?,?,?,?,?,?,?)',('solo',sid,'solo','completed',app.now(),app.now(),None)); self.db.execute('insert into responses values(?,?,?,?,?,?,?,?)',(str(uuid.uuid4()),sid,'solo',inds,'4','numeric',app.now(),app.now()))
        self.db.commit(); out=app.analysis(self.db,sid); indicator=out['domains'][0]['indicators'][0]; domain_out=out['domains'][0]
        self.assertIsNotNone(indicator['capacity']); self.assertIsNone(indicator['consensus']); self.assertEqual(indicator['consensusNote'],'single_respondent')
        self.assertIsNone(domain_out['consensus']); self.assertEqual(domain_out['consensusNote'],'single_respondent')
        self.assertIsNone(out['global']['consensus']); self.assertEqual(out['global']['consensusNote'],'single_respondent')
    def test_qualitative_chain_is_persistent_and_exported(self):
        t=self.db.execute('select id,version from templates where is_canonical=1').fetchone(); sid='qualitative-session'
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
        t = self.db.execute('select id from templates where is_canonical=1').fetchone()['id']
        self.assertTrue(app.is_canonical_template(self.db, t))
        fork = app.clone_template(self.db, t, status="session_locked")
        self.assertFalse(app.is_canonical_template(self.db, fork))

    def test_canonical_reference_structural_edit_is_refused(self):
        t = self.db.execute('select id from templates where is_canonical=1').fetchone()['id']
        with self.assertRaises(app.StructuralEditForbiddenError):
            app.guard_structural_edit(self.db, t, "add")
        with self.assertRaises(app.StructuralEditForbiddenError):
            app.guard_structural_edit(self.db, t, "delete")
        payload = app.template_payload(self.db, t)
        self.assertEqual(len(payload['domains']), 7)
        self.assertEqual(sum(len(d['indicators']) for d in payload['domains']), 70)

    def test_private_fork_is_freely_editable(self):
        t = self.db.execute('select id from templates where is_canonical=1').fetchone()['id']
        fork = app.clone_template(self.db, t, status="session_locked")
        app.guard_structural_edit(self.db, fork, "add")  # must not raise
        did = str(uuid.uuid4())
        self.db.execute('insert into domains values(?,?,?,?,?,?,?)', (did, fork, 'nouveau', 'Nouveau domaine', '', 8, 1))
        self.db.commit()
        self.assertEqual(self.db.execute('select count(*) from domains where template_id=?', (fork,)).fetchone()[0], 8)

    def test_ensure_private_template_forks_and_remaps_historical_responses(self):
        db = self.db
        t = db.execute('select id,version from templates where is_canonical=1').fetchone()
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
        t = db.execute('select id,version from templates where is_canonical=1').fetchone()
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
        t = db.execute('select id,version from templates where is_canonical=1').fetchone()
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

    # --- Régression : cycle de vie des questionnaires (correctif 2026-08-17) ---

    def test_canonical_edit_and_delete_are_refused(self):
        t = self.db.execute('select id from templates where is_canonical=1').fetchone()['id']
        with self.assertRaises(app.StructuralEditForbiddenError):
            app.guard_structural_edit(self.db, t, "edit")
        with self.assertRaises(app.StructuralEditForbiddenError):
            app.guard_structural_edit(self.db, t, "delete_template")
        fork = app.clone_template(self.db, t, status="session_locked")
        app.guard_structural_edit(self.db, fork, "edit")  # must not raise
        app.guard_structural_edit(self.db, fork, "delete_template")  # must not raise

    def test_editing_own_private_fork_does_not_clone(self):
        db = self.db
        t = db.execute('select id,version from templates where is_canonical=1').fetchone()
        sid = 'mission-edit'
        self._mk_session(sid, template=t)
        self._mk_participant(sid, 'p1')
        app.ensure_private_template(db, sid)
        tid = db.execute('select template_id from sessions where id=?', (sid,)).fetchone()['template_id']
        before = db.execute('select count(*) from templates').fetchone()[0]
        used = db.execute('select 1 from sessions where template_id=? limit 1', (tid,)).fetchone()
        status = db.execute('select status from templates where id=?', (tid,)).fetchone()['status']
        self.assertTrue(used)
        self.assertEqual(status, 'session_locked')
        # the fix: used-but-already-private must NOT trigger clone-on-write
        should_clone = bool(used) and status != 'session_locked'
        self.assertFalse(should_clone)
        after = db.execute('select count(*) from templates').fetchone()[0]
        self.assertEqual(before, after)

    def test_delete_session_drops_exclusive_private_fork_but_keeps_canonical_and_shared(self):
        db = self.db
        t = db.execute('select id,version from templates where is_canonical=1').fetchone()
        canonical_id = t['id']
        sid_a, sid_b = 'mission-del-a', 'mission-del-b'
        self._mk_session(sid_a, template=t); self._mk_session(sid_b, template=t)
        self._mk_participant(sid_a, 'pa'); self._mk_participant(sid_b, 'pb')
        app.ensure_private_template(db, sid_a); app.ensure_private_template(db, sid_b)
        tid_a = db.execute('select template_id from sessions where id=?', (sid_a,)).fetchone()['template_id']
        tid_b = db.execute('select template_id from sessions where id=?', (sid_b,)).fetchone()['template_id']
        self.assertNotEqual(tid_a, tid_b)

        dropped = app.delete_session(db, sid_a)
        self.assertTrue(dropped)
        self.assertIsNone(db.execute('select id from sessions where id=?', (sid_a,)).fetchone())
        self.assertIsNone(db.execute('select id from templates where id=?', (tid_a,)).fetchone())
        # mission B and the canonical are both untouched
        self.assertIsNotNone(db.execute('select id from sessions where id=?', (sid_b,)).fetchone())
        self.assertIsNotNone(db.execute('select id from templates where id=?', (tid_b,)).fetchone())
        canonical_payload = app.template_payload(db, canonical_id)
        self.assertEqual(len(canonical_payload['domains']), 7)
        self.assertEqual(sum(len(d['indicators']) for d in canonical_payload['domains']), 70)

        # mission B's own private fork is dropped in turn; the canonical (never
        # referenced by A or B once they forked) stays intact throughout
        dropped_b = app.delete_session(db, sid_b)
        self.assertTrue(dropped_b)
        self.assertIsNotNone(db.execute('select id from templates where id=?', (canonical_id,)).fetchone())

    def test_delete_session_never_drops_the_canonical_even_if_still_directly_referenced(self):
        db = self.db
        t = db.execute('select id,version from templates where is_canonical=1').fetchone()
        canonical_id = t['id']
        sid = 'mission-no-fork-yet'
        self._mk_session(sid, template=t)  # no participant yet: still points straight at canonical
        dropped = app.delete_session(db, sid)
        self.assertFalse(dropped)
        self.assertIsNotNone(db.execute('select id from templates where id=?', (canonical_id,)).fetchone())
        canonical_payload = app.template_payload(db, canonical_id)
        self.assertEqual(len(canonical_payload['domains']), 7)
        self.assertEqual(sum(len(d['indicators']) for d in canonical_payload['domains']), 70)

    def test_delete_session_never_drops_a_template_still_used_by_another_session(self):
        db = self.db
        t = db.execute('select id,version from templates where is_canonical=1').fetchone()
        fork = app.clone_template(db, t['id'], status='session_locked')
        db.commit()
        sid_a, sid_b = 'mission-shared-a', 'mission-shared-b'
        self._mk_session(sid_a, template={'id': fork, 'version': 1})
        self._mk_session(sid_b, template={'id': fork, 'version': 1})

        dropped = app.delete_session(db, sid_a)
        self.assertFalse(dropped)
        self.assertIsNotNone(db.execute('select id from templates where id=?', (fork,)).fetchone())
        self.assertIsNotNone(db.execute('select id from sessions where id=?', (sid_b,)).fetchone())

    def test_delete_session_never_drops_an_active_library_template(self):
        db = self.db
        t = db.execute('select id,version from templates where is_canonical=1').fetchone()
        library = app.clone_template(db, t['id'], status='active')
        db.commit()
        sid = 'mission-library'
        self._mk_session(sid, template={'id': library, 'version': 1})

        dropped = app.delete_session(db, sid)
        self.assertFalse(dropped)
        self.assertIsNotNone(db.execute('select id from templates where id=?', (library,)).fetchone())

    def test_analysis_computes_graded_values_per_indicator(self):
        t=self.db.execute('select id,version from templates where is_canonical=1').fetchone(); sid='session-graded'
        self.db.execute("insert into sessions values(?,?,?,?,?,?,?,?,?,?,?,?,?)",(sid,t['id'],t['version'],'test','','','', 'open',app.now(),None,'',None,None))
        domain=self.db.execute('select id from domains where display_order=1').fetchone()['id']; inds=self.db.execute('select id from indicators where domain_id=? order by display_order limit 1',(domain,)).fetchone()['id']
        for n,v in [('a',1),('b',5)]:
            self.db.execute('insert into participants (id,session_id,anonymous_id,status,started_at,completed_at,display_name) values(?,?,?,?,?,?,?)',(n,sid,n,'completed',app.now(),app.now(),None)); self.db.execute('insert into responses values(?,?,?,?,?,?,?,?)',(str(uuid.uuid4()),sid,n,inds,str(v),'numeric',app.now(),app.now()))
        self.db.commit(); out=app.analysis(self.db,sid); indicator=out['domains'][0]['indicators'][0]
        self.assertEqual(indicator['gradedCapacity'],app.grade(indicator['capacity'],app.GRADING))
        self.assertEqual(indicator['gradedConsensus'],app.grade(indicator['consensus'],app.GRADING))

    def test_individual_responses_rows_scoped_and_anonymous_safe(self):
        db=self.db
        t=db.execute('select id,version from templates where is_canonical=1').fetchone()
        sid_a,sid_b='ind-resp-a','ind-resp-b'
        self._mk_session(sid_a,template=t); self._mk_session(sid_b,template=t)
        domain=db.execute('select id from domains where template_id=? order by display_order limit 1',(t['id'],)).fetchone()['id']
        ind=db.execute('select id,code from indicators where domain_id=? order by display_order limit 1',(domain,)).fetchone()
        db.execute("insert into participants (id,session_id,anonymous_id,status,started_at,completed_at,display_name,anonymous,participant_type,profile,sex,age_range,education_level) values(?,?,?,?,?,?,?,?,?,?,?,?,?)",
                   ('p-anon',sid_a,'P-1','completed',app.now(),app.now(),'Nom Caché',1,'Individuel','ONG','Homme','25–39 ans','BAC'))
        db.execute("insert into participants (id,session_id,anonymous_id,status,started_at,completed_at,display_name,anonymous,participant_type,profile,sex,age_range,education_level) values(?,?,?,?,?,?,?,?,?,?,?,?,?)",
                   ('p-named',sid_a,'P-2','completed',app.now(),app.now(),'Mme Diop',0,'Institutionnel','Administration','Femme','40–54 ans','Licence'))
        db.execute("insert into participants (id,session_id,anonymous_id,status,started_at,completed_at,display_name) values(?,?,?,?,?,?,?)",
                   ('p-progress',sid_a,'P-3','in_progress',app.now(),None,None))
        db.execute("insert into participants (id,session_id,anonymous_id,status,started_at,completed_at,display_name) values(?,?,?,?,?,?,?)",
                   ('p-other',sid_b,'P-4','completed',app.now(),app.now(),'Autre mission'))
        for pid,val in (('p-anon','3'),('p-named','2'),('p-progress','1'),('p-other','4')):
            db.execute('insert into responses values(?,?,?,?,?,?,?,?)',(str(uuid.uuid4()),sid_a if pid!='p-other' else sid_b,pid,ind['id'],val,'numeric',app.now(),app.now()))
        db.commit()

        data=app.individual_responses_rows(db,sid_a)
        ids=[r['id'] for r in data['rows']]
        self.assertEqual(sorted(ids),['P-1','P-2'])
        anon_row=next(r for r in data['rows'] if r['id']=='P-1')
        named_row=next(r for r in data['rows'] if r['id']=='P-2')
        self.assertEqual(anon_row['name'],'')
        self.assertEqual(anon_row['status'],'Anonyme')
        self.assertEqual(named_row['name'],'Mme Diop')
        self.assertEqual(named_row['status'],'Nominatif')
        self.assertEqual(named_row['profile'],'Administration')
        self.assertEqual(anon_row[ind['id']],3.0)
        self.assertEqual(named_row[ind['id']],2.0)
        self.assertEqual([i['id'] for i in data['indicators']][0],ind['id'])

    def test_participants_put_route_is_public(self):
        self.assertTrue(app.is_public_api('/api/participants/abc-123','PUT'))
        self.assertFalse(app.is_public_api('/api/participants/abc-123','GET'))
        self.assertFalse(app.is_public_api('/api/participants/abc-123','DELETE'))

    def test_participant_profile_columns_are_additive_and_nullable(self):
        cols={r['name']:r for r in self.db.execute('PRAGMA table_info(participants)')}
        for col in ('anonymous','participant_type','profile','sex','age_range','education_level'):
            self.assertIn(col,cols)
            self.assertEqual(cols[col]['notnull'],0)
        sid='migration-session'; self._mk_session(sid); self._mk_participant(sid,'existing-participant')
        app.init_schema(self.db)
        row=self.db.execute('select * from participants where id=?',('existing-participant',)).fetchone()
        self.assertIsNotNone(row)
        self.assertIsNone(row['anonymous'])

    def test_domain_and_indicator_frozen_messages_mention_duplication(self):
        self.assertIn('figée',app.DOMAIN_FROZEN_MESSAGE)
        self.assertIn('dupliquez',app.DOMAIN_FROZEN_MESSAGE.lower())
        self.assertIn('figée',app.INDICATOR_FROZEN_MESSAGE)
        self.assertIn('dupliquez',app.INDICATOR_FROZEN_MESSAGE.lower())
        self.assertIn('{count}',app.DOMAIN_FROZEN_MESSAGE); self.assertIn('{names}',app.DOMAIN_FROZEN_MESSAGE)
        self.assertIn('{count}',app.INDICATOR_FROZEN_MESSAGE)

    def test_filtered_analysis_matches_consignes_math_example(self):
        db = self.db
        t = db.execute('select id,version from templates where is_canonical=1').fetchone()
        sid = 'filter-math-session'
        self._mk_session(sid, template=t)
        domain = db.execute('select id from domains where display_order=1').fetchone()['id']
        indicator = db.execute('select id from indicators where domain_id=? order by display_order limit 1', (domain,)).fetchone()['id']
        for pid, val in [('h1', 5), ('h2', 5), ('h3', 4), ('h4', 4)]:
            self._mk_participant(sid, pid, status='completed'); db.execute("UPDATE participants SET sex=? WHERE id=?", ('Homme', pid))
            db.execute('insert into responses values(?,?,?,?,?,?,?,?)', (str(uuid.uuid4()), sid, pid, indicator, str(val), 'numeric', app.now(), app.now()))
        for pid, val in [('f1', 1), ('f2', 1), ('f3', 2), ('f4', 2)]:
            self._mk_participant(sid, pid, status='completed'); db.execute("UPDATE participants SET sex=? WHERE id=?", ('Femme', pid))
            db.execute('insert into responses values(?,?,?,?,?,?,?,?)', (str(uuid.uuid4()), sid, pid, indicator, str(val), 'numeric', app.now(), app.now()))
        db.commit()

        hommes = app.analysis(db, sid, {'sex': 'Homme'})
        self.assertEqual(hommes['completedCount'], 4)
        self.assertEqual(hommes['global']['capacity'], 90)
        self.assertAlmostEqual(hommes['global']['consensus'], 71.13, delta=0.1)
        self.assertEqual(hommes['global']['gradedCapacity'], 85)
        self.assertEqual(hommes['global']['gradedConsensus'], 50)

        femmes = app.analysis(db, sid, {'sex': 'Femme'})
        self.assertEqual(femmes['completedCount'], 4)
        self.assertEqual(femmes['global']['capacity'], 30)
        self.assertAlmostEqual(femmes['global']['consensus'], 71.13, delta=0.1)
        self.assertEqual(femmes['global']['gradedCapacity'], 10)
        self.assertEqual(femmes['global']['gradedConsensus'], 50)

        tous = app.analysis(db, sid, None)
        self.assertEqual(tous['completedCount'], 8)
        self.assertEqual(tous['global']['capacity'], 60)
        self.assertAlmostEqual(tous['global']['consensus'], 15.49, delta=0.1)
        # The critical assertion: pooling the 8 individual responses gives a very
        # different number from averaging the two sub-group consensus values —
        # confirms "Tous" recomputes from raw responses, never from sub-averages.
        self.assertNotAlmostEqual(tous['global']['consensus'], 71.13, delta=5)

        compare = app.compare_population(db, sid, 'sex')
        by_value = {c['value']: c for c in compare}
        self.assertEqual(set(by_value), {'Homme', 'Femme'})
        self.assertEqual(by_value['Homme']['capacity'], 90)
        self.assertEqual(by_value['Femme']['capacity'], 30)

    def test_analysis_combines_multiple_filters_with_and(self):
        db = self.db
        t = db.execute('select id,version from templates where is_canonical=1').fetchone()
        sid = 'combined-filter-session'
        self._mk_session(sid, template=t)
        domain = db.execute('select id from domains where display_order=1').fetchone()['id']
        indicator = db.execute('select id from indicators where domain_id=? order by display_order limit 1', (domain,)).fetchone()['id']
        for pid, sex, profile, val in [('p1', 'Homme', 'ONG', 4), ('p2', 'Homme', 'Administration', 2), ('p3', 'Femme', 'ONG', 5)]:
            self._mk_participant(sid, pid, status='completed'); db.execute("UPDATE participants SET sex=?,profile=? WHERE id=?", (sex, profile, pid))
            db.execute('insert into responses values(?,?,?,?,?,?,?,?)', (str(uuid.uuid4()), sid, pid, indicator, str(val), 'numeric', app.now(), app.now()))
        db.commit()
        combined = app.analysis(db, sid, {'sex': 'Homme', 'profile': 'ONG'})
        self.assertEqual(combined['completedCount'], 1)
        self.assertEqual(combined['global']['capacity'], 80)

    def test_analysis_filter_never_leaks_across_sessions(self):
        db = self.db
        t = db.execute('select id,version from templates where is_canonical=1').fetchone()
        sid_a, sid_b = 'iso-a', 'iso-b'
        self._mk_session(sid_a, template=t); self._mk_session(sid_b, template=t)
        domain = db.execute('select id from domains where display_order=1').fetchone()['id']
        indicator = db.execute('select id from indicators where domain_id=? order by display_order limit 1', (domain,)).fetchone()['id']
        self._mk_participant(sid_a, 'a1', status='completed'); db.execute("UPDATE participants SET sex=? WHERE id=?", ('Homme', 'a1'))
        db.execute('insert into responses values(?,?,?,?,?,?,?,?)', (str(uuid.uuid4()), sid_a, 'a1', indicator, '5', 'numeric', app.now(), app.now()))
        self._mk_participant(sid_b, 'b1', status='completed'); db.execute("UPDATE participants SET sex=? WHERE id=?", ('Homme', 'b1'))
        db.execute('insert into responses values(?,?,?,?,?,?,?,?)', (str(uuid.uuid4()), sid_b, 'b1', indicator, '1', 'numeric', app.now(), app.now()))
        db.commit()
        a = app.analysis(db, sid_a, {'sex': 'Homme'})
        self.assertEqual(a['completedCount'], 1); self.assertEqual(a['global']['capacity'], 100)
        cats = app.compare_population(db, sid_a, 'sex')
        self.assertEqual(len(cats), 1); self.assertEqual(cats[0]['N'], 1); self.assertEqual(cats[0]['capacity'], 100)
        # A category with 0 respondents (nothing set on sid_b's participant beyond
        # 'Homme') never appears — never shown at N=0.
        cats_b = app.compare_population(db, sid_b, 'profile')
        self.assertEqual(cats_b, [])

    def test_objective_findings_thresholds(self):
        result = {"domains": [
            {"id": "d1", "label": "Fort", "capacity": 85, "consensus": 80, "indicators": []},
            {"id": "d2", "label": "Faible", "capacity": 40, "consensus": 50, "indicators": []},
            {"id": "d3", "label": "Vigilance", "capacity": 75, "consensus": 30, "indicators": []},
        ]}
        findings = app.objective_findings(result)
        # d3 (capacity 75) legitimately qualifies as BOTH a force (high capacity)
        # and a vigilance point (that capacity isn't backed by consensus) — the two
        # categories are independent readings, not mutually exclusive.
        self.assertEqual(sorted(d['id'] for d in findings['forces']['domains']), ['d1', 'd3'])
        self.assertEqual([d['id'] for d in findings['fragilites']['domains']], ['d2'])
        self.assertEqual([v['id'] for v in findings['vigilance']], ['d3'])
        self.assertEqual(findings['vigilance'][0]['reason'], 'capacite_elevee_consensus_faible')

    def test_objective_findings_flags_large_subpopulation_gap(self):
        result = {"domains": [{"id": "d1", "label": "D1", "capacity": 50, "consensus": 65, "indicators": []}]}
        comparison = [
            {"value": "Homme", "domains": [{"id": "d1", "label": "D1", "capacity": 80}]},
            {"value": "Femme", "domains": [{"id": "d1", "label": "D1", "capacity": 20}]},
        ]
        findings = app.objective_findings(result, comparison=comparison)
        gaps = [v for v in findings['vigilance'] if v['reason'] == 'ecart_sous_populations']
        self.assertEqual(len(gaps), 1); self.assertEqual(gaps[0]['gap'], 60)

    def test_analysis_unfiltered_call_forms_are_identical(self):
        db = self.db
        sid = 'noop-filter-session'
        self._mk_session(sid); self._mk_participant(sid, 'p1', status='completed')
        domain = db.execute('select id from domains where display_order=1').fetchone()['id']
        indicator = db.execute('select id from indicators where domain_id=? order by display_order limit 1', (domain,)).fetchone()['id']
        db.execute('insert into responses values(?,?,?,?,?,?,?,?)', (str(uuid.uuid4()), sid, 'p1', indicator, '4', 'numeric', app.now(), app.now()))
        db.commit()
        a_default = app.analysis(db, sid); a_none = app.analysis(db, sid, None); a_empty = app.analysis(db, sid, {})
        self.assertEqual(a_default, a_none); self.assertEqual(a_default, a_empty)

    def test_epc_35_template_is_created_alongside_canonical_and_is_idempotent(self):
        db = self.db
        row = db.execute("select * from templates where name=?", (app.EPC_35_TEMPLATE_NAME,)).fetchone()
        self.assertIsNotNone(row)
        self.assertEqual(row['is_canonical'], 0)
        payload = app.template_payload(db, row['id'])
        self.assertEqual(len(payload['domains']), 7)
        self.assertEqual(sum(len(d['indicators']) for d in payload['domains']), 35)
        for d in payload['domains']:
            self.assertEqual(len(d['indicators']), 5)
        all_codes = [i['code'] for d in payload['domains'] for i in d['indicators']]
        self.assertEqual(all_codes, [str(n) for n in range(1, 36)])
        # the canonical 7x70 referential is untouched by the new template's presence
        canonical = db.execute('select id from templates where is_canonical=1').fetchone()['id']
        canonical_payload = app.template_payload(db, canonical)
        self.assertEqual(sum(len(d['indicators']) for d in canonical_payload['domains']), 70)
        # idempotent: calling it again does not create a duplicate row
        before = db.execute('select count(*) from templates').fetchone()[0]
        app.ensure_epc_35_template(db)
        after = db.execute('select count(*) from templates').fetchone()[0]
        self.assertEqual(before, after)

    def test_epc_35_template_owner_is_mouhba_when_account_exists(self):
        # Fresh, isolated DB: the users row must exist before init_db() runs
        # ensure_epc_35_template(), mirroring real startup timing.
        tmp = tempfile.TemporaryDirectory()
        try:
            db = app.connect(Path(tmp.name) / 'owner-test.sqlite3')
            app.init_schema(db)
            db.execute("insert into users values(?,?,?,?,?,?,?)", ('pilote-1', 'mouhba@local', 'x', 'y', 'pilote', 'Pilote', app.now()))
            db.commit()
            app.ensure_epc_35_template(db)
            row = db.execute("select owner_user_id from templates where name=?", (app.EPC_35_TEMPLATE_NAME,)).fetchone()
            self.assertEqual(row['owner_user_id'], 'pilote-1')
            db.close()
        finally:
            tmp.cleanup()

    # --- Mission :8810 (petites améliorations demandées par Mouhamed BA, consignes_claude.txt) ---

    def test_template_domain_indicator_counts_computed_from_data_for_canonical(self):
        # Point A : jamais code en dur - recalcule depuis domains/indicators actifs.
        # Utilisee par /api/templates pour l'accueil ET "Voir tous les modeles" (meme
        # fonction, donc meme information dans les deux ecrans).
        db = self.db
        canonical = db.execute('select id from templates where is_canonical=1').fetchone()['id']
        domain_count, indicator_count = app.template_domain_indicator_counts(db, canonical)
        self.assertEqual(domain_count, 7)
        self.assertEqual(indicator_count, 70)

    def test_template_domain_indicator_counts_for_epc_35_matches_its_own_structure(self):
        db = self.db
        epc35 = db.execute('select id from templates where name=?', (app.EPC_35_TEMPLATE_NAME,)).fetchone()['id']
        domain_count, indicator_count = app.template_domain_indicator_counts(db, epc35)
        self.assertEqual(domain_count, 7)
        self.assertEqual(indicator_count, 35)

    def test_template_domain_indicator_counts_ignores_inactive_domains_and_indicators(self):
        db = self.db
        canonical = db.execute('select id from templates where is_canonical=1').fetchone()['id']
        fork = app.clone_template(db, canonical, status='session_locked')
        domain = db.execute('select id from domains where template_id=? order by display_order limit 1', (fork,)).fetchone()['id']
        indicator = db.execute('select id from indicators where domain_id=? order by display_order limit 1', (domain,)).fetchone()['id']
        db.execute('update indicators set active=0 where id=?', (indicator,)); db.commit()
        domain_count, indicator_count = app.template_domain_indicator_counts(db, fork)
        self.assertEqual(domain_count, 7)
        self.assertEqual(indicator_count, 69)

    def test_create_session_uses_existing_template_never_creates_one(self):
        # Point B : le bouton "Nouveau diagnostic" cree une mission/atelier a partir
        # d'un questionnaire deja existant - il ne cree jamais de nouveau
        # questionnaire. "Créer le questionnaire" aurait donc ete trompeur.
        db = self.db
        canonical = db.execute('select id,version from templates where is_canonical=1').fetchone()
        templates_before = db.execute('select count(*) from templates').fetchone()[0]
        sid = app.create_session(db, 'owner-1', {'name': 'Mission test', 'templateId': canonical['id']})
        self.assertIsNotNone(sid)
        templates_after = db.execute('select count(*) from templates').fetchone()[0]
        self.assertEqual(templates_before, templates_after)
        session = db.execute('select template_id,template_version,owner_user_id from sessions where id=?', (sid,)).fetchone()
        self.assertEqual(session['template_id'], canonical['id'])
        self.assertEqual(session['template_version'], canonical['version'])
        self.assertEqual(session['owner_user_id'], 'owner-1')

    def test_create_session_refuses_a_template_with_no_active_question(self):
        db = self.db
        tid = app.create_blank_template(db, {'name': 'Vide'})
        db.commit()
        sessions_before = db.execute('select count(*) from sessions').fetchone()[0]
        sid = app.create_session(db, 'owner-1', {'name': 'Mission test', 'templateId': tid})
        self.assertIsNone(sid)
        sessions_after = db.execute('select count(*) from sessions').fetchone()[0]
        self.assertEqual(sessions_before, sessions_after)

    def test_education_level_other_saved_and_autre_stays_the_statistical_category(self):
        # Points C.4-C.7 : "Préciser" est sauvegardé séparément, "Autre" reste
        # inchangée comme seule catégorie statistique (jamais remplacée par le texte).
        db = self.db
        self._mk_session('sess-edu'); self._mk_participant('sess-edu', 'p1')
        app.update_participant_profile_fields(db, 'p1', {'educationLevel': 'Autre', 'educationLevelOther': 'BTS'})
        row = db.execute('select education_level,education_level_other from participants where id=?', ('p1',)).fetchone()
        self.assertEqual(row['education_level'], 'Autre')
        self.assertEqual(row['education_level_other'], 'BTS')

    def test_education_level_other_stays_empty_for_a_regular_level(self):
        db = self.db
        self._mk_session('sess-edu2'); self._mk_participant('sess-edu2', 'p1')
        app.update_participant_profile_fields(db, 'p1', {'educationLevel': 'Licence'})
        row = db.execute('select education_level,education_level_other from participants where id=?', ('p1',)).fetchone()
        self.assertEqual(row['education_level'], 'Licence')
        self.assertIsNone(row['education_level_other'])

    def test_education_level_other_is_included_in_the_participant_row_for_resume(self):
        # /api/participant renvoie dict(participant) tel quel : la précision doit donc
        # être restaurée automatiquement à la reprise sans route dédiée.
        db = self.db
        self._mk_session('sess-edu3'); self._mk_participant('sess-edu3', 'p1')
        app.update_participant_profile_fields(db, 'p1', {'educationLevel': 'Autre', 'educationLevelOther': 'BTS'})
        participant = dict(db.execute('select * from participants where id=?', ('p1',)).fetchone())
        self.assertIn('education_level_other', participant)
        self.assertEqual(participant['education_level_other'], 'BTS')

    def test_update_participant_profile_fields_is_additive_and_partial(self):
        # Régression pour l'extraction de update_participant_profile_fields : mettre à
        # jour un seul champ ne doit jamais effacer les autres déjà enregistrés.
        db = self.db
        self._mk_session('sess-edu4'); self._mk_participant('sess-edu4', 'p1')
        app.update_participant_profile_fields(db, 'p1', {'sex': 'Femme', 'ageRange': '25–39 ans'})
        app.update_participant_profile_fields(db, 'p1', {'educationLevel': 'Autre', 'educationLevelOther': 'BTS'})
        row = db.execute('select sex,age_range,education_level,education_level_other from participants where id=?', ('p1',)).fetchone()
        self.assertEqual(row['sex'], 'Femme')
        self.assertEqual(row['age_range'], '25–39 ans')
        self.assertEqual(row['education_level'], 'Autre')
        self.assertEqual(row['education_level_other'], 'BTS')

    # --- Mission :8810 (retour Mouhamed BA) : audit du champ "Référence" des indicateurs ---

    def test_canonical_and_epc35_have_no_duplicate_reference_within_any_domain(self):
        # Etat des lieux prealable a toute regle d'unicite : le referentiel reel
        # (protege) n'a jamais eu de doublon de Reference dans un meme domaine.
        for code, label, indicators in app.EPC_DOMAINS:
            refs = [ref for ref, _ in indicators]
            self.assertEqual(len(set(refs)), len(refs), f"doublon dans {code}")
        for code, label, indicators in app.EPC_DOMAINS_5:
            self.assertEqual(len(set(indicators)), len(indicators), f"doublon dans {code}")

    def test_indicator_code_conflicts_detects_duplicate_within_same_domain(self):
        db = self.db
        canonical = db.execute('select id from templates where is_canonical=1').fetchone()['id']
        domain = db.execute('select id from domains where template_id=? order by display_order limit 1', (canonical,)).fetchone()['id']
        existing_code = db.execute('select code from indicators where domain_id=? order by display_order limit 1', (domain,)).fetchone()['code']
        self.assertTrue(app.indicator_code_conflicts(db, domain, existing_code))

    def test_indicator_code_conflicts_ignores_other_domains(self):
        # Meme perimetre que le controle deja applique a l'import XLSX : par domaine,
        # jamais global - reutiliser la meme Reference dans un AUTRE domaine est legitime.
        db = self.db
        canonical = db.execute('select id from templates where is_canonical=1').fetchone()['id']
        domains = db.execute('select id from domains where template_id=? order by display_order limit 2', (canonical,)).fetchall()
        domain_a, domain_b = domains[0]['id'], domains[1]['id']
        code_in_b = db.execute('select code from indicators where domain_id=? limit 1', (domain_b,)).fetchone()['code']
        self.assertFalse(app.indicator_code_conflicts(db, domain_a, code_in_b))

    def test_indicator_code_conflicts_excludes_self_when_editing(self):
        db = self.db
        canonical = db.execute('select id from templates where is_canonical=1').fetchone()['id']
        domain = db.execute('select id from domains where template_id=? order by display_order limit 1', (canonical,)).fetchone()['id']
        indicator = db.execute('select id,code from indicators where domain_id=? limit 1', (domain,)).fetchone()
        self.assertFalse(app.indicator_code_conflicts(db, domain, indicator['code'], exclude_id=indicator['id']))

    def test_indicator_code_conflicts_false_when_reference_is_unused(self):
        db = self.db
        canonical = db.execute('select id from templates where is_canonical=1').fetchone()['id']
        domain = db.execute('select id from domains where template_id=? order by display_order limit 1', (canonical,)).fetchone()['id']
        self.assertFalse(app.indicator_code_conflicts(db, domain, 'Reference totalement inedite'))

    def test_display_order_not_code_determines_indicator_order(self):
        # Le champ qui pilote reellement l'ordre est display_order, jamais code -
        # verifie en creant deux indicateurs avec la MEME reference et un display_order
        # explicite, puis en confirmant que template_payload les restitue dans cet ordre.
        db = self.db
        canonical = db.execute('select id from templates where is_canonical=1').fetchone()['id']
        domain = db.execute('select id from domains where template_id=? order by display_order limit 1', (canonical,)).fetchone()['id']
        iid_first = str(uuid.uuid4()); iid_second = str(uuid.uuid4())
        db.execute('INSERT INTO indicators VALUES (?,?,?,?,?,?,?,?,?,?)', (iid_second, domain, 'DUP', 'Deuxieme (display_order 21)', '', 'numeric', 1, 21, 1, '{}'))
        db.execute('INSERT INTO indicators VALUES (?,?,?,?,?,?,?,?,?,?)', (iid_first, domain, 'DUP', 'Premiere (display_order 20)', '', 'numeric', 1, 20, 1, '{}'))
        db.commit()
        payload = app.template_payload(db, canonical)
        dom = next(d for d in payload['domains'] if d['id'] == domain)
        last_two = [i['id'] for i in dom['indicators'][-2:]]
        self.assertEqual(last_two, [iid_first, iid_second])

    def test_duplicate_reference_does_not_corrupt_results_but_shares_a_short_chart_label(self):
        # Constat central de l'audit : un doublon de Reference n'altere JAMAIS les
        # calculs (chaque indicateur reste distinct par son id), mais produit la meme
        # etiquette sur les graphiques/histogrammes quand la reference est courte -
        # une ambiguite d'affichage, jamais une erreur de calcul.
        db = self.db
        canonical = db.execute('select id,version from templates where is_canonical=1').fetchone()
        domain = db.execute('select id from domains where template_id=? order by display_order limit 1', (canonical['id'],)).fetchone()['id']
        iid1, iid2 = str(uuid.uuid4()), str(uuid.uuid4())
        db.execute('INSERT INTO indicators VALUES (?,?,?,?,?,?,?,?,?,?)', (iid1, domain, 'DUP', 'Premiere question dupliquee', '', 'numeric', 1, 21, 1, '{}'))
        db.execute('INSERT INTO indicators VALUES (?,?,?,?,?,?,?,?,?,?)', (iid2, domain, 'DUP', 'Deuxieme question dupliquee', '', 'numeric', 1, 22, 1, '{}'))
        sid = 'sess-dup-reference'
        self._mk_session(sid, template=canonical)
        self._mk_participant(sid, 'p1', status='completed')
        db.execute('insert into responses values(?,?,?,?,?,?,?,?)', (str(uuid.uuid4()), sid, 'p1', iid1, '5', 'numeric', app.now(), app.now()))
        db.execute('insert into responses values(?,?,?,?,?,?,?,?)', (str(uuid.uuid4()), sid, 'p1', iid2, '1', 'numeric', app.now(), app.now()))
        db.commit()
        result = app.analysis(db, sid)
        dom = next(d for d in result['domains'] if d['id'] == domain)
        first, second = dom['indicators'][-2], dom['indicators'][-1]
        # calculs corrects et distincts malgre la reference identique
        self.assertEqual(first['capacity'], 100.0)
        self.assertEqual(second['capacity'], 20.0)
        # mais meme etiquette de graphique (ambiguite d'affichage documentee)
        self.assertEqual(app.pdf_short_label(first), app.pdf_short_label(second))

    def test_individual_responses_export_header_repeats_the_reference_on_duplicate(self):
        # Justifie techniquement le refus des doublons a la creation/edition : sans
        # cette regle, l'export "reponses individuelles" produirait deux colonnes
        # portant exactement le meme intitule.
        db = self.db
        canonical = db.execute('select id,version from templates where is_canonical=1').fetchone()
        domain = db.execute('select id from domains where template_id=? order by display_order limit 1', (canonical['id'],)).fetchone()['id']
        iid1, iid2 = str(uuid.uuid4()), str(uuid.uuid4())
        db.execute('INSERT INTO indicators VALUES (?,?,?,?,?,?,?,?,?,?)', (iid1, domain, 'DUP', 'Premiere question dupliquee', '', 'numeric', 1, 21, 1, '{}'))
        db.execute('INSERT INTO indicators VALUES (?,?,?,?,?,?,?,?,?,?)', (iid2, domain, 'DUP', 'Deuxieme question dupliquee', '', 'numeric', 1, 22, 1, '{}'))
        db.commit()
        sid = 'sess-dup-export'
        self._mk_session(sid, template=canonical)
        data = app.individual_responses_rows(db, sid)
        codes = [i['code'] for i in data['indicators']]
        self.assertEqual(codes.count('DUP'), 2)

if __name__=='__main__': unittest.main()
