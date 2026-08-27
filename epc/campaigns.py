"""Campagnes/groupes/sessions : libelle, code de groupe, inventaire et
suppression en cascade, generation des kits relais.

Extrait de app.py (lot 1d de la modularisation, cf. AUDIT_MODULARISATION_8800.md) :
aucune formule ni aucun contrat de route ne change ici, seule l'emplacement du
code metier bouge. app.py reimporte ces symboles a l'identique.
"""
from __future__ import annotations

import secrets
import sqlite3
import uuid
import zipfile
from io import BytesIO

from .auth import relay_token_hash
from .db import now, rows, template_payload
from .profile import ensure_default_profile_schema
from .util import slugify

# Every table that hangs off a session row via session_id — deleting a session (alone,
# or as part of a campaign cascade) always means deleting exactly these first.
SESSION_CHILD_TABLES = ("training_topics", "workshop_recommendations", "analysis_entries", "priority_analyses",
                         "analysis_notes", "recommendations", "responses", "participant_profile_values",
                         "priorities", "participants", "session_report_meta", "retained_comparisons")

GROUP_COLORS = ["#2563eb", "#dc2626", "#16a34a", "#d97706", "#7c3aed", "#0891b2", "#db2777", "#65a30d", "#ea580c", "#4338ca"]


def esc_html(s) -> str:
    return (str(s or "")).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def session_label(db: sqlite3.Connection, sid: str) -> str:
    row = db.execute("SELECT name FROM sessions WHERE id=?", (sid,)).fetchone()
    return row["name"] if row else "atelier"


def generate_group_code(db, name):
    """group_code is only ever used for display/filenames, never as a data-selection
    key — but it's generated globally unique across every campaign anyway, defensively,
    so it can never become one without a future change silently reintroducing an ambiguity.
    """
    all_codes = {r["group_code"] for r in db.execute("SELECT group_code FROM sessions WHERE group_code IS NOT NULL")}
    base_code = slugify(name)[:3].upper() or "GRP"
    n = 1
    while f"{base_code}-{n:02d}" in all_codes: n += 1
    return f"{base_code}-{n:02d}"


def campaign_deletion_summary(db, campaign_id):
    """Read-only inventory shown to a pilote before they confirm a permanent
    campaign deletion (or select it in the test-data cleanup tool). Never
    mutates anything."""
    camp = db.execute("SELECT * FROM campaigns WHERE id=?", (campaign_id,)).fetchone()
    if not camp:
        return None
    groups = rows(db, """SELECT s.id,s.name,s.group_code,s.relay_name,s.expected_participants,
        (SELECT COUNT(*) FROM participants p WHERE p.session_id=s.id) AS participant_count,
        (SELECT COUNT(*) FROM participants p WHERE p.session_id=s.id AND p.status='completed') AS completed_count,
        (SELECT COUNT(*) FROM responses r WHERE r.session_id=s.id) AS response_count,
        (SELECT COUNT(*) FROM priority_analyses a WHERE a.session_id=s.id) AS analysis_count,
        (SELECT COUNT(*) FROM workshop_recommendations w WHERE w.session_id=s.id) AS recommendation_count
        FROM sessions s WHERE s.campaign_id=? ORDER BY s.created_at""", (campaign_id,))
    return {
        "campaign": dict(camp),
        "groups": groups,
        "responseCount": sum(g["response_count"] for g in groups),
        "participantCount": sum(g["participant_count"] for g in groups),
        "completedCount": sum(g["completed_count"] for g in groups),
        "analysisCount": sum(g["analysis_count"] + g["recommendation_count"] for g in groups),
    }


def delete_group_cascade(db, campaign_id, session_id, force=False):
    """Same force-bypass idiom as delete_campaign_cascade, but for a single
    group within a campaign. Scoped by both session_id and campaign_id so it
    can never delete a group belonging to a different campaign."""
    used = db.execute("SELECT COUNT(*) FROM responses WHERE session_id=?", (session_id,)).fetchone()[0]
    if used and not force:
        return False, used
    for table in SESSION_CHILD_TABLES:
        db.execute(f"DELETE FROM {table} WHERE session_id=?", (session_id,))
    db.execute("DELETE FROM sessions WHERE id=? AND campaign_id=?", (session_id, campaign_id))
    db.commit()
    return True, used


def delete_campaign_cascade(db, campaign_id, force=False):
    """Permanently deletes a campaign and everything scoped under it (groups =
    sessions with this campaign_id, and every SESSION_CHILD_TABLES row for
    each of those groups). Never touches templates/domains/indicators (the
    questionnaire is always shared, never campaign-owned) and never touches a
    row belonging to another campaign or another pilote, since every DELETE
    below is scoped by session_id/campaign_id.

    Without force=True, refuses (returns deleted=False) if the campaign
    already has recorded responses, mirroring the confirmation the pilote
    sees on screen before they can pass force=True.
    """
    used = db.execute("SELECT COUNT(*) FROM responses r JOIN sessions s ON s.id=r.session_id WHERE s.campaign_id=?", (campaign_id,)).fetchone()[0]
    if used and not force:
        return False, used
    for r in db.execute("SELECT id FROM sessions WHERE campaign_id=?", (campaign_id,)).fetchall():
        for table in SESSION_CHILD_TABLES:
            db.execute(f"DELETE FROM {table} WHERE session_id=?", (r["id"],))
        db.execute("DELETE FROM sessions WHERE id=?", (r["id"],))
    db.execute("DELETE FROM campaigns WHERE id=?", (campaign_id,))
    db.commit()
    return True, used


def campaign_kits_zip(db: sqlite3.Connection, campaign_id: str, base_url: str) -> bytes:
    """One simple standalone HTML per group: campaign/group/relay + the two
    links. The QR itself is not duplicated here — opening the relay link shows
    it live (same tested qrSvg()/generateQR() already used everywhere else).

    Relay tokens are only ever stored hashed (never retrievable), so building
    fresh relay links here means regenerating every group's token — this is
    the same "regenerate" capability exposed individually, just applied to
    the whole campaign at once. Any previously shared relay links stop working.
    """
    camp = db.execute("SELECT * FROM campaigns WHERE id=?", (campaign_id,)).fetchone()
    if not camp:
        raise ValueError("Campagne introuvable")
    groups = rows(db, "SELECT * FROM sessions WHERE campaign_id=? ORDER BY created_at", (campaign_id,))
    buf = BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for g in groups:
            raw_token = secrets.token_urlsafe(24)
            db.execute("UPDATE sessions SET relay_token_hash=? WHERE id=?", (relay_token_hash(raw_token), g["id"]))
            participant_link = f"{base_url}/?session={g['id']}"
            relay_link = f"{base_url}/?relay={raw_token}"
            html = f"""<!doctype html><html lang="fr"><meta charset="utf-8"><title>Kit relais — {esc_html(g['name'])}</title>
<body style="font-family:sans-serif;max-width:640px;margin:2rem auto;line-height:1.5">
<h1>{esc_html(camp['name'])}</h1>
<h2>Groupe : {esc_html(g['name'])} {f"({esc_html(g['group_code'])})" if g['group_code'] else ''}</h2>
<p><b>Relais :</b> {esc_html(g['relay_name'] or 'Non renseigné')}</p>
<p><b>Objectif (participants prévus) :</b> {g['expected_participants'] or 'Non défini'}</p>
<hr>
<p><b>Lien participant</b> (à transmettre / afficher en QR) :<br><a href="{participant_link}">{participant_link}</a></p>
<p><b>Lien de suivi relais</b> (votre tableau de bord + votre QR code) :<br><a href="{relay_link}">{relay_link}</a></p>
<hr>
<p>Consigne : partagez le lien participant (ou son QR, visible sur votre lien de suivi) aux personnes de votre groupe. Suivez l'avancée de la collecte depuis votre lien de suivi — aucune configuration ni accès aux autres groupes n'y est nécessaire.</p>
</body></html>"""
            zf.writestr(f"kit_{slugify(g['group_code'] or g['name'])}.html", html)
    db.commit()
    return buf.getvalue()


def create_campaign(db: sqlite3.Connection, owner_user_id: str, template_id: str, template_version: int, data: dict) -> str:
    cid, stamp = str(uuid.uuid4()), now()
    # The campaign's own default profile schema (correctifs cibles :8820, cf.
    # consignes_claude.txt) - created ONCE here and shared by every group of
    # this campaign (see create_group()), so a dimension is never created
    # independently per group in the first place.
    template = template_payload(db, template_id)
    profile_schema_id = ensure_default_profile_schema(db, owner_user_id, template.get("model_key")) if template else None
    db.execute("INSERT INTO campaigns VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
        (cid, owner_user_id, data["name"], data.get("description", ""), data.get("periodStart"), data.get("periodEnd"), template_id, template_version, "active", stamp, stamp, profile_schema_id))
    db.commit()
    return cid


def update_campaign(db: sqlite3.Connection, campaign, data: dict) -> None:
    db.execute("UPDATE campaigns SET name=?,description=?,period_start=?,period_end=?,status=?,updated_at=? WHERE id=?",
        (data.get("name", campaign["name"]), data.get("description", campaign["description"]), data.get("periodStart", campaign["period_start"]), data.get("periodEnd", campaign["period_end"]), data.get("status", campaign["status"]), now(), campaign["id"]))
    db.commit()


def create_group(db: sqlite3.Connection, campaign, owner_user_id: str, data: dict) -> dict:
    campaign_codes = {r["group_code"] for r in db.execute("SELECT group_code FROM sessions WHERE campaign_id=?", (campaign["id"],))}
    group_code = generate_group_code(db, data["name"])
    group_color = GROUP_COLORS[len(campaign_codes) % len(GROUP_COLORS)]
    sid = str(uuid.uuid4())
    raw_token = secrets.token_urlsafe(24)
    profile_schema_id = data.get("profileSchemaId")
    if profile_schema_id is None:
        # The CAMPAIGN's own schema is the one source of truth for every group
        # (correctifs cibles :8820, PRIORITE CRITIQUE) - never create a fresh
        # default schema per group, or a dimension added to one group would
        # silently never exist for its siblings. A campaign created before
        # this fix (or whose template had no default profile) may still have
        # none yet: create it once here AND persist it on the campaign so
        # every subsequent group in this campaign reuses the exact same row.
        profile_schema_id = campaign["profile_schema_id"]
        if profile_schema_id is None:
            template = template_payload(db, campaign["template_id"])
            profile_schema_id = ensure_default_profile_schema(db, owner_user_id, template.get("model_key"))
            if profile_schema_id:
                db.execute("UPDATE campaigns SET profile_schema_id=? WHERE id=?", (profile_schema_id, campaign["id"]))
    db.execute("INSERT INTO sessions (id,template_id,template_version,name,organization,location,date,status,created_at,closed_at,description,expected_participants,owner_user_id,campaign_id,group_code,group_color,relay_name,relay_token_hash,profile_schema_id) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
        (sid, campaign["template_id"], campaign["template_version"], data["name"], "", "", "", "open", now(), None, "",
         int(data["expectedParticipants"]) if data.get("expectedParticipants") not in (None, "") else None,
         owner_user_id, campaign["id"], group_code, group_color, data.get("relayName") or "", relay_token_hash(raw_token), profile_schema_id))
    db.commit()
    return {"id": sid, "groupCode": group_code, "groupColor": group_color, "relayToken": raw_token}


def regenerate_group_relay(db: sqlite3.Connection, session_id: str) -> str:
    raw_token = secrets.token_urlsafe(24)
    db.execute("UPDATE sessions SET relay_token_hash=? WHERE id=?", (relay_token_hash(raw_token), session_id))
    db.commit()
    return raw_token


def create_session(db: sqlite3.Connection, owner_user_id: str, data: dict):
    """Standalone (non-campaign) session. Returns None if the questionnaire has
    no active domain with an active question, same refusal as the historical route."""
    template = template_payload(db, data["templateId"])
    if not template or not any(d["active"] and any(i["active"] for i in d["indicators"]) for d in template["domains"]):
        return None
    sid = str(uuid.uuid4())
    profile_schema_id = data.get("profileSchemaId")
    if profile_schema_id is None:
        # The template's OWN model_key (never the restitution fallback that
        # treats every untagged/custom questionnaire as EPC/SENEVAL for report
        # purposes) - a default profile is only a fit for the genuine model.
        profile_schema_id = ensure_default_profile_schema(db, owner_user_id, template.get("model_key"))
    db.execute("INSERT INTO sessions (id,template_id,template_version,name,organization,location,date,status,created_at,closed_at,description,expected_participants,owner_user_id,campaign_id,group_code,group_color,relay_name,relay_token_hash,profile_schema_id) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
        (sid, template["id"], template["version"], data["name"], data.get("organization", ""), data.get("location", ""), data.get("date", ""), "open", now(), None, data.get("description", ""),
         int(data["expectedParticipants"]) if data.get("expectedParticipants") not in (None, "") else None, owner_user_id, None, None, None, None, None, profile_schema_id))
    db.commit()
    return sid


def update_session(db: sqlite3.Connection, session_id: str, data: dict) -> bool:
    """Returns False (no mutation) if a requested templateId doesn't exist,
    same refusal as the historical route."""
    expected = int(data["expectedParticipants"]) if data.get("expectedParticipants") not in (None, "") else None
    # profileSchemaId is a new, not-yet-frontend-wired field (lot 4a): unlike the
    # older fields below (which the existing edit form always resends in full),
    # a caller that omits it entirely must not silently wipe a value set elsewhere
    # (e.g. by create_group/create_session) - only an explicit key (including an
    # explicit null, to clear it) changes it.
    existing_session = db.execute("SELECT profile_schema_id, campaign_id FROM sessions WHERE id=?", (session_id,)).fetchone()
    if "profileSchemaId" in data:
        profile_schema_id = data["profileSchemaId"]
    else:
        profile_schema_id = existing_session["profile_schema_id"] if existing_session else None
    if data.get("templateId"):
        tpl = db.execute("SELECT version FROM templates WHERE id=?", (data["templateId"],)).fetchone()
        if not tpl:
            return False
        db.execute("UPDATE sessions SET name=?,organization=?,location=?,date=?,description=?,expected_participants=?,template_id=?,template_version=?,profile_schema_id=? WHERE id=?",
            (data["name"], data.get("organization", ''), data.get("location", ''), data.get("date", ''), data.get("description", ''), expected, data["templateId"], tpl["version"], profile_schema_id, session_id))
    else:
        db.execute("UPDATE sessions SET name=?,organization=?,location=?,date=?,description=?,expected_participants=?,profile_schema_id=? WHERE id=?",
            (data["name"], data.get("organization", ''), data.get("location", ''), data.get("date", ''), data.get("description", ''), expected, profile_schema_id, session_id))
    # A profile attached/detached from a campaign group must propagate to the
    # whole campaign (correctifs cibles :8820) - the campaign stays the one
    # source of truth every group resolves through (resolve_session_profile_
    # schema_id, epc/profile.py), so writing it here is what makes an edit
    # made from any single group's Configuration screen visible to its siblings.
    if "profileSchemaId" in data and existing_session and existing_session["campaign_id"]:
        db.execute("UPDATE campaigns SET profile_schema_id=? WHERE id=?", (profile_schema_id, existing_session["campaign_id"]))
    db.commit()
    return True
