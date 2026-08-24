"""Profil participant composable (lot 4a de la modularisation, cf.
AUDIT_MODULARISATION_8800.md) : definition de schemas/champs types et
stockage/validation des valeurs par participant.

Code neuf (pas une extraction) : contrairement aux lots precedents, aucune
contrainte de "copie a l'identique" ne s'applique ici - c'est une nouvelle
fonctionnalite, desactivee par defaut (sessions.profile_schema_id est NULL
tant qu'un pilote ne choisit pas explicitement un profil). Backend seul dans
ce lot : aucun formulaire participant dynamique n'est branche cote frontend.
"""
from __future__ import annotations

import json
import sqlite3
import uuid

from .db import now, rows
from .templates import next_order
from .util import slugify

FIELD_TYPES = ("text", "number", "single_choice", "multi_choice")
CHOICE_FIELD_TYPES = ("single_choice", "multi_choice")


def _validate_choice_and_dimension(field_type: str, options: list, is_dimension: bool) -> None:
    """Shared by create_profile_field()/update_profile_field(): options are
    required for choice-type fields, and only choice-type fields may become
    a dimension - kept in one place so the two functions can't silently
    drift on which combinations are accepted."""
    if field_type in CHOICE_FIELD_TYPES and not options:
        raise ValueError("Les options sont obligatoires pour un champ à choix.")
    if is_dimension and field_type not in CHOICE_FIELD_TYPES:
        raise ValueError("Seuls les champs à choix unique ou multiple peuvent devenir une dimension d'analyse.")


def _session_profile_schema_id(db: sqlite3.Connection, session_id: str) -> str | None:
    """Shared by set_participant_profile_values()/available_dimensions()/
    resolve_dimension_field(): the session's attached profile schema id, or
    None if the session doesn't exist or has no profile attached."""
    session = db.execute("SELECT profile_schema_id FROM sessions WHERE id=?", (session_id,)).fetchone()
    return session["profile_schema_id"] if session else None


def create_profile_schema(db: sqlite3.Connection, owner_user_id: str | None, data: dict) -> str:
    name = (data.get("name") or "").strip()
    if not name:
        raise ValueError("Le nom du profil est obligatoire.")
    sid, stamp = str(uuid.uuid4()), now()
    db.execute("INSERT INTO profile_schemas (id,name,description,owner_user_id,created_at,updated_at) VALUES (?,?,?,?,?,?)",
        (sid, name, data.get("description", ""), owner_user_id, stamp, stamp))
    db.commit()
    return sid


def update_profile_schema(db: sqlite3.Connection, schema_id: str, data: dict) -> None:
    existing = db.execute("SELECT * FROM profile_schemas WHERE id=?", (schema_id,)).fetchone()
    if not existing:
        raise ValueError("Profil introuvable.")
    name = (data.get("name", existing["name"]) or "").strip() or existing["name"]
    db.execute("UPDATE profile_schemas SET name=?,description=?,updated_at=? WHERE id=?",
        (name, data.get("description", existing["description"]), now(), schema_id))
    db.commit()


def delete_profile_schema(db: sqlite3.Connection, schema_id: str) -> str:
    """Returns "in_use" (a session references it, or a participant already has
    a value under one of its fields) or "deleted"."""
    if db.execute("SELECT 1 FROM sessions WHERE profile_schema_id=? LIMIT 1", (schema_id,)).fetchone():
        return "in_use"
    if db.execute("SELECT 1 FROM participant_profile_values v JOIN profile_fields f ON f.id=v.field_id WHERE f.schema_id=? LIMIT 1", (schema_id,)).fetchone():
        return "in_use"
    db.execute("DELETE FROM profile_fields WHERE schema_id=?", (schema_id,))
    db.execute("DELETE FROM profile_schemas WHERE id=?", (schema_id,))
    db.commit()
    return "deleted"


def profile_schema_payload(db: sqlite3.Connection, schema_id: str):
    schema = db.execute("SELECT * FROM profile_schemas WHERE id=?", (schema_id,)).fetchone()
    if not schema:
        return None
    out = dict(schema)
    out["fields"] = rows(db, "SELECT * FROM profile_fields WHERE schema_id=? ORDER BY display_order", (schema_id,))
    for f in out["fields"]:
        f["required"] = bool(f["required"]); f["active"] = bool(f["active"]); f["is_dimension"] = bool(f["is_dimension"]); f["options"] = json.loads(f.pop("options_json"))
    return out


def create_profile_field(db: sqlite3.Connection, schema_id: str, data: dict) -> str:
    field_type = data.get("fieldType")
    if field_type not in FIELD_TYPES:
        raise ValueError(f"Type de champ invalide : {field_type}")
    label = (data.get("label") or "").strip()
    if not label:
        raise ValueError("Le libellé du champ est obligatoire.")
    options = data.get("options") or []
    is_dimension = bool(data.get("isDimension"))
    _validate_choice_and_dimension(field_type, options, is_dimension)
    fid = str(uuid.uuid4())
    # slugify() intentionally preserves case (export_filename() wants that for
    # readability) - a field_key is a machine identifier though, so lowercase it
    # explicitly here rather than changing slugify()'s shared behaviour.
    key = data.get("key") or slugify(label).lower()
    # `or` would treat an explicit displayOrder of 0 as falsy and silently
    # replace it with the auto-computed order - check for None instead.
    display_order = data.get("displayOrder")
    display_order = int(display_order) if display_order is not None else next_order(db, "profile_fields", "schema_id", schema_id)
    db.execute("INSERT INTO profile_fields (id,schema_id,field_key,field_type,label,required,options_json,display_order,active,is_dimension) VALUES (?,?,?,?,?,?,?,?,?,?)",
        (fid, schema_id, key, field_type, label, int(bool(data.get("required"))), json.dumps(options), display_order, int(data.get("active", True)), int(is_dimension)))
    db.commit()
    return fid


def update_profile_field(db: sqlite3.Connection, field_id: str, data: dict) -> None:
    existing = db.execute("SELECT * FROM profile_fields WHERE id=?", (field_id,)).fetchone()
    if not existing:
        raise ValueError("Champ introuvable.")
    field_type = data.get("fieldType", existing["field_type"])
    if field_type not in FIELD_TYPES:
        raise ValueError(f"Type de champ invalide : {field_type}")
    options = data.get("options", json.loads(existing["options_json"]))
    is_dimension = bool(data.get("isDimension", existing["is_dimension"]))
    _validate_choice_and_dimension(field_type, options, is_dimension)
    label = (data.get("label", existing["label"]) or "").strip() or existing["label"]
    db.execute("UPDATE profile_fields SET field_key=?,field_type=?,label=?,required=?,options_json=?,display_order=?,active=?,is_dimension=? WHERE id=?",
        (data.get("key", existing["field_key"]), field_type, label, int(bool(data.get("required", existing["required"]))), json.dumps(options), int(data.get("displayOrder", existing["display_order"])), int(data.get("active", existing["active"])), int(is_dimension), field_id))
    db.commit()


def delete_profile_field(db: sqlite3.Connection, field_id: str) -> tuple[bool, int]:
    """Refuses (returns False, used_count) if any participant already has a
    value stored under this field, same historical-preservation invariant as
    indicator/domain deletion."""
    used = db.execute("SELECT COUNT(*) FROM participant_profile_values WHERE field_id=?", (field_id,)).fetchone()[0]
    if used:
        return False, used
    db.execute("DELETE FROM profile_fields WHERE id=?", (field_id,))
    db.commit()
    return True, used


def validate_profile_value(field: dict, raw_value):
    """Normalizes and type-checks a single submitted value against its field
    definition; raises ValueError (caught by Handler's generic 400 handler)
    with a message safe to show the moderator/participant."""
    label = field["label"]
    if raw_value is None or raw_value == "" or raw_value == []:
        if field["required"]:
            raise ValueError(f"Le champ « {label} » est obligatoire.")
        return None
    field_type = field["field_type"]
    options = json.loads(field["options_json"]) if field["options_json"] else []
    # Compared as strings (not raw equality) so an option value stored/edited
    # as a non-string (e.g. numeric) still matches a submitted value of a
    # different-but-equal-looking type - avoids rejecting a valid choice as
    # invalid purely because of a JSON type mismatch.
    allowed = {str(o["value"] if isinstance(o, dict) else o) for o in options}
    if field_type == "text":
        if not isinstance(raw_value, str):
            raise ValueError(f"Le champ « {label} » doit être du texte.")
        return raw_value
    if field_type == "number":
        if not isinstance(raw_value, (int, float)) or isinstance(raw_value, bool):
            raise ValueError(f"Le champ « {label} » doit être un nombre.")
        return raw_value
    if field_type == "single_choice":
        if str(raw_value) not in allowed:
            raise ValueError(f"Valeur invalide pour « {label} ».")
        return raw_value
    if field_type == "multi_choice":
        if not isinstance(raw_value, list) or any(str(v) not in allowed for v in raw_value):
            raise ValueError(f"Valeur invalide pour « {label} ».")
        return raw_value
    raise ValueError(f"Type de champ inconnu : {field_type}")


def set_participant_profile_values(db: sqlite3.Connection, participant_id: str, values: dict) -> None:
    """Partial upsert: only the fields present in `values` are validated and
    stored (mirrors submit_response()'s per-indicator upsert), so a participant
    can fill their profile incrementally rather than all at once."""
    participant = db.execute("SELECT session_id FROM participants WHERE id=?", (participant_id,)).fetchone()
    if not participant:
        raise ValueError("Participant introuvable.")
    session_id = participant["session_id"]
    schema_id = _session_profile_schema_id(db, session_id)
    if not schema_id:
        raise ValueError("Cet atelier n'a pas de profil participant configuré.")
    fields = {f["field_key"]: f for f in rows(db, "SELECT * FROM profile_fields WHERE schema_id=? AND active=1", (schema_id,))}
    stamp = now()
    for key, raw_value in values.items():
        field = fields.get(key)
        if not field:
            raise ValueError(f"Champ de profil inconnu : {key}")
        normalized = validate_profile_value(field, raw_value)
        if normalized is None:
            # Optional field left blank: nothing to store: keeps
            # get_participant_profile_values() free of empty-string/None noise for
            # fields the participant simply didn't answer.
            db.execute("DELETE FROM participant_profile_values WHERE participant_id=? AND field_id=?", (participant_id, field["id"]))
            continue
        db.execute("INSERT INTO participant_profile_values (id,session_id,participant_id,field_id,value_json,created_at,updated_at) VALUES (?,?,?,?,?,?,?) "
            "ON CONFLICT(participant_id,field_id) DO UPDATE SET value_json=excluded.value_json,updated_at=excluded.updated_at",
            (str(uuid.uuid4()), session_id, participant_id, field["id"], json.dumps(normalized), stamp, stamp))
    db.commit()


def get_participant_profile_values(db: sqlite3.Connection, participant_id: str) -> dict:
    result = {}
    for r in db.execute("""SELECT f.field_key, v.value_json FROM participant_profile_values v
                            JOIN profile_fields f ON f.id=v.field_id WHERE v.participant_id=?""", (participant_id,)):
        result[r["field_key"]] = json.loads(r["value_json"])
    return result


def available_dimensions(db: sqlite3.Connection, session_id: str) -> list[dict]:
    """Categorical profile fields the pilot has explicitly flagged as an
    analytical dimension for this session's attached profile (lot 5, cf.
    AUDIT_MODULARISATION_8800.md) - empty list if the session has no profile
    schema, or if none of its fields are flagged. Powers the filter/comparison
    UI's "which dimension can I use" choices."""
    schema_id = _session_profile_schema_id(db, session_id)
    if not schema_id:
        return []
    payload = profile_schema_payload(db, schema_id)
    return [{"fieldKey": f["field_key"], "label": f["label"], "fieldType": f["field_type"], "options": f["options"]}
            for f in payload["fields"] if f["is_dimension"] and f["active"]]


def resolve_dimension_field(db: sqlite3.Connection, session_id: str, field_key: str) -> dict:
    """The sole gate that lets a query filter analysis by a categorical
    field: refuses any field_key that isn't an active, pilot-flagged
    dimension of the session's own attached profile. This is the privacy
    boundary the audit calls out for lot 5 - never resolve a dimension
    filter any other way (e.g. trusting a raw field_key from the client
    without this check would let anyone probe arbitrary profile values)."""
    schema_id = _session_profile_schema_id(db, session_id)
    if not schema_id:
        raise ValueError("Cet atelier n'a pas de profil participant configuré.")
    field = db.execute("SELECT * FROM profile_fields WHERE schema_id=? AND field_key=? AND active=1", (schema_id, field_key)).fetchone()
    if not field or not field["is_dimension"]:
        raise ValueError("Ce champ n'est pas configuré comme dimension d'analyse.")
    return dict(field)


def participants_matching_dimension_values(db: sqlite3.Connection, session_id: str, field_key: str, values: list) -> dict:
    """Same matching as participants_matching_dimension(), for several values
    of the same field at once from a single query over its rows - avoids
    re-scanning the identical session_id+field_key rows once per compared
    value (the comparison screen compares several values of one dimension
    in a single request). Caller must have already validated field_key via
    resolve_dimension_field."""
    result: dict = {value: set() for value in values}
    for r in db.execute("""SELECT v.participant_id, v.value_json FROM participant_profile_values v
                            JOIN profile_fields f ON f.id=v.field_id
                            WHERE v.session_id=? AND f.field_key=?""", (session_id, field_key)):
        stored = json.loads(r["value_json"])
        for value in values:
            if isinstance(stored, list):
                if value in stored:
                    result[value].add(r["participant_id"])
            elif stored == value:
                result[value].add(r["participant_id"])
    return result


def participants_matching_dimension(db: sqlite3.Connection, session_id: str, field_key: str, value) -> set[str]:
    """Participant ids of `session_id` whose profile value for `field_key`
    equals (single_choice) or contains (multi_choice) `value`. Caller must
    have already validated field_key via resolve_dimension_field."""
    return participants_matching_dimension_values(db, session_id, field_key, [value])[value]
