"""Campagnes/groupes/sessions : libelle, code de groupe, inventaire et
suppression en cascade, generation des kits relais.

Extrait de app.py (lot 1d de la modularisation, cf. AUDIT_MODULARISATION_8800.md) :
aucune formule ni aucun contrat de route ne change ici, seule l'emplacement du
code metier bouge. app.py reimporte ces symboles a l'identique.
"""
from __future__ import annotations

import secrets
import sqlite3
import zipfile
from io import BytesIO

from .auth import relay_token_hash
from .db import rows
from .util import slugify

# Every table that hangs off a session row via session_id — deleting a session (alone,
# or as part of a campaign cascade) always means deleting exactly these first.
SESSION_CHILD_TABLES = ("training_topics", "workshop_recommendations", "analysis_entries", "priority_analyses",
                         "analysis_notes", "recommendations", "responses", "priorities", "participants", "session_report_meta")


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
