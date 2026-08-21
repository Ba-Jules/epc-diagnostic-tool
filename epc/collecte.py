"""Collecte participant : inscription, reponses, validation, reprise.

Extrait de app.py (lot 1e de la modularisation, cf. AUDIT_MODULARISATION_8800.md) :
aucune formule ni aucun contrat de route ne change ici, seule l'emplacement du
code metier bouge. app.py reimporte ces symboles a l'identique.
"""
from __future__ import annotations

import json
import sqlite3
import uuid

from .db import now, template_payload
from .profile import get_participant_profile_values, profile_schema_payload


class CollecteClosedError(Exception):
    """Raised when trying to register a participant on a session whose status isn't 'open'."""


def create_participant(db: sqlite3.Connection, session_id: str, data: dict) -> dict:
    session = db.execute("SELECT status FROM sessions WHERE id=?", (session_id,)).fetchone()
    if not session or session["status"] != "open":
        raise CollecteClosedError()
    pid = str(uuid.uuid4())
    label = data.get("anonymousId") or f"P-{uuid.uuid4().hex[:6]}"
    db.execute("INSERT INTO participants VALUES (?,?,?,?,?,?,?)", (pid, session_id, label, "in_progress", now(), None, data.get("displayName") or None))
    db.commit()
    return {"id": pid, "anonymousId": label}


def submit_response(db: sqlite3.Connection, session_id: str, data: dict) -> None:
    stamp = now()
    db.execute("INSERT INTO responses VALUES (?,?,?,?,?,?,?,?) ON CONFLICT(participant_id,indicator_id) DO UPDATE SET value_json=excluded.value_json,value_type=excluded.value_type,updated_at=excluded.updated_at",
        (str(uuid.uuid4()), session_id, data["participantId"], data["indicatorId"], json.dumps(data["value"]), data.get("valueType", "numeric"), stamp, stamp))
    db.commit()


def complete_participant(db: sqlite3.Connection, participant_id: str) -> None:
    db.execute("UPDATE participants SET status='completed',completed_at=? WHERE id=?", (now(), participant_id))
    db.commit()


def update_participant_display_name(db: sqlite3.Connection, participant_id: str, display_name: str | None) -> None:
    db.execute("UPDATE participants SET display_name=? WHERE id=?", (display_name or None, participant_id))
    db.commit()


def participant_resume(db: sqlite3.Connection, session_id: str, participant_id: str) -> dict:
    participant = db.execute("SELECT * FROM participants WHERE id=? AND session_id=?", (participant_id, session_id)).fetchone()
    session = db.execute("SELECT * FROM sessions WHERE id=?", (session_id,)).fetchone()
    # profile/profileValues are additive (lot 4a): a session without profile_schema_id
    # (the default, unchanged for every pre-existing session) yields None/{} exactly
    # as if these two keys didn't exist, so older clients ignoring them are unaffected.
    schema_id = session["profile_schema_id"] if session else None
    return {
        "session": dict(session) if session else None,
        "participant": dict(participant) if participant else None,
        "template": template_payload(db, session["template_id"]) if session else None,
        "responses": {r["indicator_id"]: json.loads(r["value_json"]) for r in db.execute("SELECT * FROM responses WHERE participant_id=?", (participant_id,))},
        "profile": profile_schema_payload(db, schema_id) if schema_id else None,
        "profileValues": get_participant_profile_values(db, participant_id) if participant else {},
    }
