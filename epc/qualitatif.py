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

from .db import now


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
