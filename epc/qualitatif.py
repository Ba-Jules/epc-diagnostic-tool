"""Chaine qualitative : priorites, constat, causes/consequences/leviers,
recommandations, themes de formation, note de rapport ; plus les anciennes
routes V1 (analysis_notes/recommendations), conservees a l'identique.

Extrait de app.py (lot 1f de la modularisation, cf. AUDIT_MODULARISATION_8800.md) :
aucune formule ni aucun contrat de route ne change ici, seule l'emplacement du
code metier bouge. app.py reimporte ces symboles a l'identique.

Les fonctions de suppression avec dependances (delete_analysis_entry,
delete_workshop_recommendation) suivent le meme idiome force-bypass que
delete_group_cascade/delete_campaign_cascade dans epc/campaigns.py : elles
renvoient (deleted, dependent_count) et laissent Handler composer le message
d'erreur specifique a la route.
"""
from __future__ import annotations

import sqlite3
import uuid

from .db import now, rows


def toggle_priority(db: sqlite3.Connection, session_id: str, data: dict) -> None:
    db.execute("INSERT INTO priorities VALUES (?,?,?,?,?,?) ON CONFLICT(session_id,indicator_id) DO UPDATE SET votes=excluded.votes",
        (str(uuid.uuid4()), session_id, data["domainId"], data["indicatorId"], int(data.get("votes", 0)), now()))
    db.commit()


def delete_priority(db: sqlite3.Connection, session_id: str, indicator_id: str) -> None:
    db.execute("DELETE FROM priorities WHERE session_id=? AND indicator_id=?", (session_id, indicator_id))
    db.commit()


def upsert_priority_analysis(db: sqlite3.Connection, session_id: str, data: dict) -> None:
    stamp = now(); priority_id = data["priorityId"]
    db.execute("INSERT INTO priority_analyses VALUES (?,?,?,?,?,?) ON CONFLICT(session_id,priority_id) DO UPDATE SET problem=excluded.problem,updated_at=excluded.updated_at",
        (str(uuid.uuid4()), session_id, priority_id, data.get("problem", ""), stamp, stamp))
    db.commit()


def update_priority_analysis(db: sqlite3.Connection, analysis_id: str, problem: str | None) -> None:
    db.execute("UPDATE priority_analyses SET problem=?,updated_at=? WHERE id=?", (problem or "", now(), analysis_id))
    db.commit()


def create_analysis_entry(db: sqlite3.Connection, session_id: str, data: dict) -> str:
    stamp = now(); eid = str(uuid.uuid4())
    db.execute("INSERT INTO analysis_entries VALUES (?,?,?,?,?,?,?,?,?,?,?)",
        (eid, session_id, data["priorityId"], data.get("parentId") or None, data["kind"], data["content"], data.get("itemType") or None, data.get("comment") or None, data.get("validationStatus", "A_DISCUTER"), stamp, stamp))
    db.commit()
    return eid


def update_analysis_entry(db: sqlite3.Connection, entry_id: str, data: dict) -> None:
    db.execute("UPDATE analysis_entries SET parent_id=?,content=?,item_type=?,comment=?,validation_status=?,updated_at=? WHERE id=?",
        (data.get("parentId") or None, data["content"], data.get("itemType") or None, data.get("comment") or None, data.get("validationStatus", "A_DISCUTER"), now(), entry_id))
    db.commit()


def delete_analysis_entry(db: sqlite3.Connection, entry_id: str, force: bool = False) -> tuple[bool, int]:
    dependent = db.execute("SELECT COUNT(*) FROM workshop_recommendations WHERE cause_id=? OR lever_id=?", (entry_id, entry_id)).fetchone()[0]
    if dependent and not force:
        return False, dependent
    db.execute("UPDATE workshop_recommendations SET cause_id=NULL WHERE cause_id=?", (entry_id,))
    db.execute("UPDATE workshop_recommendations SET lever_id=NULL WHERE lever_id=?", (entry_id,))
    db.execute("UPDATE analysis_entries SET parent_id=NULL WHERE parent_id=?", (entry_id,))
    db.execute("DELETE FROM analysis_entries WHERE id=?", (entry_id,))
    db.commit()
    return True, dependent


def create_workshop_recommendation(db: sqlite3.Connection, session_id: str, data: dict) -> str:
    stamp = now(); rid = str(uuid.uuid4())
    db.execute("INSERT INTO workshop_recommendations VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
        (rid, session_id, data.get("priorityId") or None, data.get("causeId") or None, data.get("leverId") or None, data["title"], data["description"], data.get("category", "Autre"), data.get("priorityLevel", "Non définie"), data.get("owner") or None, data.get("horizon") or None, data.get("comment") or None, data.get("status", "Proposée"), stamp, stamp))
    db.commit()
    return rid


def update_workshop_recommendation(db: sqlite3.Connection, rec_id: str, data: dict) -> None:
    db.execute("UPDATE workshop_recommendations SET priority_id=?,cause_id=?,lever_id=?,title=?,description=?,category=?,priority_level=?,owner=?,horizon=?,comment=?,status=?,updated_at=? WHERE id=?",
        (data.get("priorityId") or None, data.get("causeId") or None, data.get("leverId") or None, data["title"], data["description"], data.get("category", "Autre"), data.get("priorityLevel", "Non définie"), data.get("owner") or None, data.get("horizon") or None, data.get("comment") or None, data.get("status", "Proposée"), now(), rec_id))
    db.commit()


def delete_workshop_recommendation(db: sqlite3.Connection, rec_id: str, force: bool = False) -> tuple[bool, int]:
    dependent = db.execute("SELECT COUNT(*) FROM training_topics WHERE recommendation_id=?", (rec_id,)).fetchone()[0]
    if dependent and not force:
        return False, dependent
    db.execute("UPDATE training_topics SET recommendation_id=NULL WHERE recommendation_id=?", (rec_id,))
    db.execute("DELETE FROM workshop_recommendations WHERE id=?", (rec_id,))
    db.commit()
    return True, dependent


def create_training_topic(db: sqlite3.Connection, session_id: str, data: dict) -> str:
    stamp = now(); tid = str(uuid.uuid4())
    db.execute("INSERT INTO training_topics VALUES (?,?,?,?,?,?,?,?,?,?,?)",
        (tid, session_id, data.get("priorityId") or None, data.get("recommendationId") or None, data["title"], data.get("needText") or None, data.get("targetAudience") or None, data.get("priorityLevel", "Non définie"), data.get("comment") or None, stamp, stamp))
    db.commit()
    return tid


def update_training_topic(db: sqlite3.Connection, topic_id: str, data: dict) -> None:
    db.execute("UPDATE training_topics SET priority_id=?,recommendation_id=?,title=?,need_text=?,target_audience=?,priority_level=?,comment=?,updated_at=? WHERE id=?",
        (data.get("priorityId") or None, data.get("recommendationId") or None, data["title"], data.get("needText") or None, data.get("targetAudience") or None, data.get("priorityLevel", "Non définie"), data.get("comment") or None, now(), topic_id))
    db.commit()


def delete_training_topic(db: sqlite3.Connection, topic_id: str) -> None:
    db.execute("DELETE FROM training_topics WHERE id=?", (topic_id,))
    db.commit()


def upsert_report_meta(db: sqlite3.Connection, session_id: str, data: dict) -> None:
    db.execute("INSERT INTO session_report_meta VALUES (?,?,?,?,?,?) ON CONFLICT(session_id) DO UPDATE SET facilitator=excluded.facilitator,audience=excluded.audience,context=excluded.context,conclusion=excluded.conclusion,updated_at=excluded.updated_at",
        (session_id, data.get("facilitator", ""), data.get("audience", ""), data.get("context", ""), data.get("conclusion", ""), now()))
    db.commit()


def create_analysis_note(db: sqlite3.Connection, session_id: str, data: dict) -> None:
    """Legacy V1 route (/analysis-notes), superseded by priority-analyses/analysis-entries
    but kept working as-is: some older sessions may still use it.

    Fix (previously broken since before this modularisation, see AUDIT_MODULARISATION_8800.md
    lot 1f commit message): the INSERT had 9 "?" placeholders for analysis_notes' 8 columns
    (id, session_id, indicator_id, kind, content, validation_status, created_at, updated_at),
    so this route always raised sqlite3.OperationalError. Placeholder count corrected to 8;
    no column, no value, no ordering changed."""
    stamp = now()
    db.execute("INSERT INTO analysis_notes VALUES (?,?,?,?,?,?,?,?)",
        (str(uuid.uuid4()), session_id, data.get("indicatorId"), data["kind"], data["content"], data.get("validationStatus", "HYPOTHESE"), stamp, stamp))
    db.commit()


def create_legacy_recommendation(db: sqlite3.Connection, session_id: str, data: dict) -> None:
    """Legacy V1 route (/recommendations), superseded by recommendations-v2 but
    kept working as-is: some older sessions may still use it."""
    db.execute("INSERT INTO recommendations VALUES (?,?,?,?,?,?,?,?,?,?)",
        (str(uuid.uuid4()), session_id, data.get("indicatorId"), data["title"], data.get("description", ""), data.get("lever", ""), data.get("kind", "action"), data.get("owner", ""), data.get("horizon", ""), now()))
    db.commit()


# --- Migration V1 -> V2 (lot 7, cf. AUDIT_MODULARISATION_8800.md) ---
# analysis_notes/recommendations ("V1") are superseded by analysis_entries/
# workshop_recommendations ("V2", the only source the final report reads via
# qualitative_data()) but were left both live in the UI - a moderator using
# the V1 "Causes et recommandations" form (now removed from domainDiagnostic,
# see the frontend bugfix commit) had their work silently absent from every
# export. migrate_legacy_qualitative_data() copies any V1 rows forward so
# nothing already captured is lost; it does not delete or alter V1 rows.

LEGACY_NOTE_KIND_MAP = {
    "Cause": "cause", "Cause racine potentielle": "cause", "Symptôme": "cause",
    "Facteur externe": "cause", "Hypothèse": "cause",
    "Conséquence": "consequence", "Levier": "lever",
}
LEGACY_NOTE_STATUS_MAP = {"HYPOTHESE": "A_DISCUTER", "FAIT_VALIDE": "RETENU"}
LEGACY_RECOMMENDATION_CATEGORY_MAP = {"formation": "Formation", "organisation": "Organisation", "gouvernance": "Gouvernance", "action": "Autre", "autre": "Autre"}


def _ensure_priority_for_indicator(db: sqlite3.Connection, session_id: str, indicator_id: str) -> str | None:
    """V1 rows attach to an indicator_id directly; V2 rows attach to a
    priority_id (a domain_id+indicator_id pair the group explicitly
    selected). If this indicator was never separately selected as a
    priority, create one (votes=0) so the migrated content has somewhere to
    attach - additive, matches this codebase's other historical-preservation
    migrations (never destructive, never re-interprets existing data)."""
    existing = db.execute("SELECT id FROM priorities WHERE session_id=? AND indicator_id=?", (session_id, indicator_id)).fetchone()
    if existing:
        return existing["id"]
    indicator = db.execute("SELECT domain_id FROM indicators WHERE id=?", (indicator_id,)).fetchone()
    if not indicator:
        return None
    pid = str(uuid.uuid4())
    db.execute("INSERT INTO priorities VALUES (?,?,?,?,?,?)", (pid, session_id, indicator["domain_id"], indicator_id, 0, now()))
    return pid


def migrate_legacy_qualitative_data(db: sqlite3.Connection, session_id: str | None = None) -> dict:
    """One-off, explicit migration from the legacy V1 qualitative tables to
    the V2 ones. NEVER called automatically (no init_db/startup hook) - run
    it deliberately via scripts/migrate_legacy_qualitative.py after
    reviewing what it reports, exactly what the audit calls for
    ("dépréciation... seulement après migration vérifiée"). session_id=None
    migrates every session in the database.

    Idempotent: a V1 row already matched by an existing V2 row with the same
    session_id and content/title+description is skipped, so running this
    twice does not duplicate data. V1 rows are only ever read, never
    modified or deleted - true backward compatibility, not a destructive
    cutover.
    """
    where = "WHERE session_id=?" if session_id else ""
    params = (session_id,) if session_id else ()
    notes = rows(db, f"SELECT * FROM analysis_notes {where}", params)
    recommendations = rows(db, f"SELECT * FROM recommendations {where}", params)
    migrated_entries = migrated_recommendations = skipped_entries = 0

    for n in notes:
        if not n["indicator_id"]:
            skipped_entries += 1
            continue
        if db.execute("SELECT 1 FROM analysis_entries WHERE session_id=? AND content=?", (n["session_id"], n["content"])).fetchone():
            continue
        priority_id = _ensure_priority_for_indicator(db, n["session_id"], n["indicator_id"])
        if not priority_id:
            skipped_entries += 1
            continue
        kind = LEGACY_NOTE_KIND_MAP.get(n["kind"], "cause")
        stamp = now()
        db.execute("INSERT INTO analysis_entries VALUES (?,?,?,?,?,?,?,?,?,?,?)",
            (str(uuid.uuid4()), n["session_id"], priority_id, None, kind, n["content"],
             n["kind"] if kind == "cause" else None, None, LEGACY_NOTE_STATUS_MAP.get(n["validation_status"], "A_DISCUTER"), stamp, stamp))
        migrated_entries += 1

    for r in recommendations:
        if db.execute("SELECT 1 FROM workshop_recommendations WHERE session_id=? AND title=? AND description=?", (r["session_id"], r["title"], r["description"])).fetchone():
            continue
        priority_id = _ensure_priority_for_indicator(db, r["session_id"], r["indicator_id"]) if r["indicator_id"] else None
        description = r["description"] or ""
        if r["lever"]:
            description = (description + f"\n\nLevier (V1) : {r['lever']}").strip()
        stamp = now()
        db.execute("INSERT INTO workshop_recommendations VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
            (str(uuid.uuid4()), r["session_id"], priority_id, None, None, r["title"], description,
             LEGACY_RECOMMENDATION_CATEGORY_MAP.get(r["kind"], "Autre"), "Non définie", r["owner"] or None, r["horizon"] or None, None, "Proposée", r["created_at"] or stamp, stamp))
        migrated_recommendations += 1

    db.commit()
    return {
        "totalNotes": len(notes), "totalRecommendations": len(recommendations),
        "migratedEntries": migrated_entries, "migratedRecommendations": migrated_recommendations,
        "skippedEntries": skipped_entries,
    }
