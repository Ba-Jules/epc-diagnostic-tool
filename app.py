"""Local workshop diagnosis engine.

Run with: python app.py
The application uses only the Python standard library and SQLite.  It is meant
to be a dependable local-first starting point, not a simulated real-time app.
"""
from __future__ import annotations

import csv
import base64
import hashlib
import json
import math
import os
import re
import secrets
import sqlite3
import sys
import time
import uuid
from datetime import datetime, timezone
from http import HTTPStatus
from http.cookies import SimpleCookie
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from io import StringIO
from io import BytesIO
from pathlib import Path
from urllib.parse import parse_qs, urlparse
import urllib.request
import urllib.error

from epc.db import (
    DATABASE, EPC_DOMAINS, GRADING, now, rows, connect, init_db,
    ensure_reference_questionnaire_version, migrate_v2_ownership, seed_epc, template_payload,
    MODEL_KEY_EPC_SENEVAL, ensure_model_identity,
)
from epc.restitution import (
    restitution_manifest, session_restitution_manifest, resolve_model_key,
    qualitative_data, report_data, report_xlsx, report_docx, _n, _c,
    individual_responses_rows, individual_responses_xlsx, individual_responses_csv,
    filtered_analysis_rows, filtered_analysis_xlsx, filtered_analysis_csv,
)
from epc.auth import (
    AuthRequiredError, PermissionDeniedError, PUBLIC_API_EXACT, is_public_api,
    PBKDF2_ITERATIONS, AUTH_TOKEN_TTL_DAYS, hash_password, verify_password,
    create_auth_token, relay_token_hash, session_cookie_header,
    resolve_current_user, enforce_ownership, resolve_auth,
)
from epc.templates import (
    MATRIX_COLUMNS, PARAMETERS, IMPORTS, next_order, clone_template,
    create_blank_template, matrix_xlsx, blank_matrix_xlsx, read_xlsx,
    import_preview, save_import, update_template, delete_template,
    create_domain, update_domain, delete_domain, create_indicator,
    update_indicator, delete_indicator,
)
from epc.util import slugify
from epc.campaigns import (
    SESSION_CHILD_TABLES, esc_html, session_label, generate_group_code,
    campaign_deletion_summary, delete_group_cascade, delete_campaign_cascade,
    campaign_kits_zip, create_campaign, update_campaign, create_group,
    regenerate_group_relay, create_session, update_session,
)
from epc.collecte import (
    CollecteClosedError, create_participant, submit_response, complete_participant,
    update_participant_display_name, participant_resume, list_session_participants,
)
from epc.qualitatif import (
    toggle_priority, delete_priority, upsert_priority_analysis, update_priority_analysis,
    create_analysis_entry, update_analysis_entry, delete_analysis_entry,
    create_workshop_recommendation, update_workshop_recommendation, delete_workshop_recommendation,
    create_training_topic, update_training_topic, delete_training_topic,
    upsert_report_meta, create_analysis_note, create_legacy_recommendation,
    migrate_legacy_qualitative_data,
)
from epc.scoring import grade, analysis, analysis_for, dimension_analysis, dimension_analysis_multi, filtered_analysis, objective_findings, MIN_COHORT_N
from epc.profile import (
    create_profile_schema, update_profile_schema, delete_profile_schema, profile_schema_payload,
    create_profile_field, update_profile_field, delete_profile_field,
    set_participant_profile_values, get_participant_profile_values, available_dimensions,
    resolve_session_profile_schema_id,
)

ROOT = Path(__file__).resolve().parent
STATIC = ROOT / "static"


def export_filename(*parts, ext: str) -> str:
    slug = "_".join(slugify(p) for p in parts if p)
    return f"{slug}_{datetime.now().strftime('%Y-%m-%d')}.{ext}"


# ==================================================
# ASSISTANT IA OPTIONNEL — couche multi-fournisseur
# ==================================================
# Aucune dépendance externe : requêtes HTTPS via urllib.request (stdlib).
# La clé API ne quitte jamais ce module : jamais renvoyée par une route GET,
# jamais journalisée, jamais écrite dans un export.

AI_PROVIDERS = {
    "gemini": {"label": "Google Gemini", "pricing": "GRATUIT", "family": "gemini",
        "models": [("gemini-2.0-flash", "Recommandé"), ("gemini-1.5-flash", "Alternatif")],
        "key_url": "https://aistudio.google.com/apikey"},
    "groq": {"label": "Groq", "pricing": "GRATUIT", "family": "openai", "base_url": "https://api.groq.com/openai/v1",
        "models": [("llama-3.3-70b-versatile", "Recommandé"), ("llama-3.1-8b-instant", "Rapide")],
        "key_url": "https://console.groq.com/keys"},
    "openrouter": {"label": "OpenRouter", "pricing": "GRATUIT", "family": "openai", "base_url": "https://openrouter.ai/api/v1",
        "models": [("openrouter/free", "Recommandé — gratuit"), ("meta-llama/llama-3.1-8b-instruct:free", "Alternatif gratuit")],
        "key_url": "https://openrouter.ai/keys"},
    "cerebras": {"label": "Cerebras", "pricing": "ESSAI", "family": "openai", "base_url": "https://api.cerebras.ai/v1",
        "models": [("llama3.1-8b", "Recommandé"), ("llama3.1-70b", "Alternatif")],
        "key_url": "https://cloud.cerebras.ai/"},
    "openai": {"label": "OpenAI", "pricing": "PAYANT", "family": "openai", "base_url": "https://api.openai.com/v1",
        "models": [("gpt-4o-mini", "Recommandé"), ("gpt-4o", "Alternatif")],
        "key_url": "https://platform.openai.com/api-keys"},
    "anthropic": {"label": "Anthropic Claude", "pricing": "PAYANT", "family": "anthropic",
        "models": [("claude-3-5-haiku-latest", "Recommandé"), ("claude-3-5-sonnet-latest", "Alternatif")],
        "key_url": "https://console.anthropic.com/settings/keys"},
    "deepseek": {"label": "DeepSeek", "pricing": "PAYANT", "family": "openai", "base_url": "https://api.deepseek.com/v1",
        "models": [("deepseek-chat", "Recommandé"), ("deepseek-reasoner", "Raisonnement approfondi")],
        "key_url": "https://platform.deepseek.com/api_keys"},
    "xai": {"label": "xAI Grok", "pricing": "PAYANT", "family": "openai", "base_url": "https://api.x.ai/v1",
        "models": [("grok-2-latest", "Recommandé")],
        "key_url": "https://console.x.ai/"},
}


class AIError(Exception):
    """Raised with a message safe to show the moderator (never a raw stack trace)."""


def _http_json(url, headers, payload, timeout=25):
    req = urllib.request.Request(url, data=json.dumps(payload).encode("utf-8"), headers=headers, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", "ignore")
        if e.code == 401 or e.code == 403:
            raise AIError("Clé API invalide ou non autorisée.")
        if e.code == 429:
            raise AIError("Quota atteint pour ce fournisseur. Réessayez plus tard.")
        if e.code == 404:
            raise AIError("Modèle indisponible chez ce fournisseur.")
        raise AIError(f"Le fournisseur IA a refusé la requête (code {e.code}).") from None
    except urllib.error.URLError:
        raise AIError("Connexion au fournisseur IA impossible.") from None
    except TimeoutError:
        raise AIError("Le fournisseur IA n'a pas répondu à temps.") from None


def _call_openai_compatible(base_url, api_key, model, system_prompt, user_prompt):
    data = _http_json(f"{base_url}/chat/completions", {"Content-Type": "application/json", "Authorization": f"Bearer {api_key}"},
        {"model": model, "messages": [{"role": "system", "content": system_prompt}, {"role": "user", "content": user_prompt}], "temperature": 0.4})
    try:
        return data["choices"][0]["message"]["content"].strip()
    except (KeyError, IndexError, TypeError):
        raise AIError("Réponse du fournisseur IA illisible.") from None


def _call_gemini(api_key, model, system_prompt, user_prompt):
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={api_key}"
    data = _http_json(url, {"Content-Type": "application/json"},
        {"systemInstruction": {"parts": [{"text": system_prompt}]}, "contents": [{"role": "user", "parts": [{"text": user_prompt}]}],
         "generationConfig": {"temperature": 0.4}})
    try:
        return data["candidates"][0]["content"]["parts"][0]["text"].strip()
    except (KeyError, IndexError, TypeError):
        if data.get("promptFeedback", {}).get("blockReason"):
            raise AIError("La demande a été bloquée par le fournisseur IA.") from None
        raise AIError("Réponse du fournisseur IA illisible.") from None


def _call_anthropic(api_key, model, system_prompt, user_prompt):
    data = _http_json("https://api.anthropic.com/v1/messages",
        {"Content-Type": "application/json", "x-api-key": api_key, "anthropic-version": "2023-06-01"},
        {"model": model, "max_tokens": 1200, "system": system_prompt, "messages": [{"role": "user", "content": user_prompt}]})
    try:
        return data["content"][0]["text"].strip()
    except (KeyError, IndexError, TypeError):
        raise AIError("Réponse du fournisseur IA illisible.") from None


def generate_ai_response(provider: str, model: str, system_prompt: str, user_prompt: str, api_key: str) -> str:
    cfg = AI_PROVIDERS.get(provider)
    if not cfg:
        raise AIError("Fournisseur IA inconnu.")
    if not api_key:
        raise AIError("Aucune clé API configurée pour ce fournisseur.")
    family = cfg["family"]
    if family == "openai":
        return _call_openai_compatible(cfg["base_url"], api_key, model, system_prompt, user_prompt)
    if family == "gemini":
        return _call_gemini(api_key, model, system_prompt, user_prompt)
    if family == "anthropic":
        return _call_anthropic(api_key, model, system_prompt, user_prompt)
    raise AIError("Fournisseur IA non implémenté.")


def get_ai_config(db):
    row = db.execute("SELECT enabled,provider,model,api_key FROM ai_config WHERE id=1").fetchone()
    if not row:
        return {"enabled": False, "provider": None, "model": None, "api_key": None}
    return {"enabled": bool(row["enabled"]), "provider": row["provider"], "model": row["model"], "api_key": row["api_key"]}


def require_ai(db):
    cfg = get_ai_config(db)
    if not cfg["enabled"]:
        raise AIError("Assistant IA désactivé.")
    if not cfg["provider"] or not cfg["api_key"]:
        raise AIError("Assistant IA non configuré (fournisseur ou clé manquants).")
    return cfg


def ai_epc_context(db, sid):
    a = analysis(db, sid)
    g = a["global"]
    lines = [f"Atelier : {a['session']['name']}", f"Commencés : {a['participantCount']} · Validés : {a['completedCount']}",
        f"Capacité globale : {_n(g['capacity'])}/100 · Consensus global : {_c(g)}/100 "
        f"(graduées : capacité {g['gradedCapacity']}, consensus {g['gradedConsensus']})", "", "Résultats par domaine :"]
    for d in a["domains"]:
        if d["capacity"] is None: continue
        lines.append(f"- {d['label']} : capacité {_n(d['capacity'])}, consensus {_c(d)} "
            f"(graduées {d['gradedCapacity']}/{d['gradedConsensus']}), {d['responses']} répondant(s)")
        for i in d["indicators"]:
            if i["capacity"] is None: continue
            lines.append(f"    · {i['label']} : capacité {_n(i['capacity'])}, consensus {_c(i)}")
    return "\n".join(lines)


def ai_priority_context(db, sid, pid):
    q = qualitative_data(db, sid)
    p = next((x for x in q["priorities"] if x["id"] == pid), None)
    if not p:
        raise AIError("Priorité introuvable.")
    an = next((x for x in q["analyses"] if x["priority_id"] == pid), None)
    entries = [e for e in q["entries"] if e["priority_id"] == pid]
    def block(kind, label):
        items = [e for e in entries if e["kind"] == kind]
        if not items: return f"{label} déjà saisi(e)s : aucun."
        return f"{label} déjà saisi(e)s :\n" + "\n".join(f"  - [{e['validation_status']}] {e['content']}" for e in items)
    lines = [f"Domaine : {p['domain_label']}", f"Indicateur : {p['indicator_code']} — {p['indicator_label']}",
        f"Description : {p['indicator_description']}", f"Constat déjà saisi : {an['problem'] if an and an['problem'] else 'aucun'}",
        block("cause", "Causes"), block("consequence", "Conséquences"), block("lever", "Leviers")]
    return p, an, entries, "\n".join(lines)


def log_ai_suggestion(db, sid, kind, target_id, provider, model):
    sug_id = str(uuid.uuid4())
    db.execute("INSERT INTO ai_suggestions VALUES (?,?,?,?,?,?,?,?,?)", (sug_id, sid, kind, target_id, provider, model, "proposed", now(), now()))
    db.commit()
    return sug_id


def _parse_json_list(text):
    t = text.strip()
    if t.startswith("```"):
        t = t.strip("`")
        if t.lower().startswith("json"): t = t[4:]
        t = t.strip()
    try:
        val = json.loads(t)
        return val if isinstance(val, list) else [val]
    except (json.JSONDecodeError, ValueError):
        m = re.search(r"\[.*\]", t, re.S)
        if m:
            try: return json.loads(m.group(0))
            except (json.JSONDecodeError, ValueError): pass
        raise AIError("Réponse du fournisseur IA illisible (format inattendu).") from None


def ai_report_context(db, sid):
    r = report_data(db, sid)
    q, m = r["qualitative"], r["meta"]
    lines = [ai_epc_context(db, sid), "", f"Contexte de l'atelier : {m.get('context') or 'non renseigné'}"]
    # Profil agrege + constats automatiques (mission de parite :8810->:8820,
    # cf. consignes_claude.txt) : le contexte IA du rapport reste limite aux
    # donnees calculees/agregees, jamais aux identites individuelles.
    if r["profile"]:
        lines += ["", "Profil agrégé des participants :"] + [f"- {field} — {value} : {n}" for field, value, n in r["profile"]]
    findings = r["analysis"].get("findings")
    if findings:
        lines += ["", "Constats automatiques (forces / fragilités / points de vigilance) :"]
        lines += [f"- Force : {d['label']} (capacité {_n(d['capacity'])})" for d in findings["forces"]["domains"]]
        lines += [f"- Force : {i['domain']} — {i['label']} (capacité {_n(i['capacity'])})" for i in findings["forces"]["indicators"]]
        lines += [f"- Fragilité : {d['label']} (capacité {_n(d['capacity'])})" for d in findings["fragilites"]["domains"]]
        lines += [f"- Fragilité : {i['domain']} — {i['label']} (capacité {_n(i['capacity'])})" for i in findings["fragilites"]["indicators"]]
        for v in findings["vigilance"]:
            if v["reason"] == "ecart_sous_populations":
                lines.append(f"- Vigilance : écart de {_n(v.get('gap'))} points pour {v['label']} entre sous-populations comparées")
            else:
                lines.append(f"- Vigilance : {v['label']} (capacité {_n(v.get('capacity'))}, consensus {_n(v.get('consensus'))})")
    lines += ["", "Priorités et analyses validées :"]
    for p in q["priorities"]:
        an = next((x for x in q["analyses"] if x["priority_id"] == p["id"]), None)
        retained = [e for e in q["entries"] if e["priority_id"] == p["id"] and e["validation_status"] == "RETENU"]
        lines.append(f"- {p['domain_label']} — {p['indicator_label']} : constat="
            f"{an['problem'] if an and an['problem'] else 'non renseigné'} ; "
            f"causes={'; '.join(e['content'] for e in retained if e['kind']=='cause') or 'aucune'} ; "
            f"leviers={'; '.join(e['content'] for e in retained if e['kind']=='lever') or 'aucun'}")
    lines.append("\nRecommandations retenues :")
    kept_recs = [x for x in q["recommendations"] if x["status"] == "Retenue"]
    lines += ([f"- {x['title']} : {x['description']}" for x in kept_recs] or ["  aucune"])
    lines.append("\nThèmes de formation :")
    lines += ([f"- {t['title']} : {t['need_text'] or ''}" for t in q["trainingTopics"]] or ["  aucun"])
    return "\n".join(lines)


class Handler(SimpleHTTPRequestHandler):
    def db(self):
        db = connect(); init_db(db); return db

    def json(self, code, payload, cookie=None):
        raw = json.dumps(payload, ensure_ascii=False).encode()
        self.send_response(code); self.send_header("Content-Type", "application/json; charset=utf-8"); self.send_header("Content-Length", str(len(raw)))
        if cookie is not None: self.send_header("Set-Cookie", cookie)
        self.end_headers(); self.wfile.write(raw)

    def body(self):
        size = int(self.headers.get("Content-Length", 0)); return json.loads(self.rfile.read(size) or b"{}")

    def raw_body(self): return self.rfile.read(int(self.headers.get("Content-Length", 0)))

    def current_user(self, db):
        return resolve_current_user(db, self.headers.get("Cookie"))

    def check_ownership(self, path, db, user):
        enforce_ownership(path, db, user)

    def require_auth(self, path, db):
        """Call first inside each verb handler's try block. Returns the current
        user (or None for the small public whitelist) and enforces per-row
        ownership for /api/sessions|templates|campaigns/<id>... routes."""
        return resolve_auth(path, self.command, db, self.headers.get("Cookie"))

    def do_GET(self):
        path, query = urlparse(self.path).path, parse_qs(urlparse(self.path).query)
        db = self.db()
        try:
            user = self.require_auth(path, db)
            if path == "/api/auth/setup-status":
                return self.json(200, {"needsSetup": db.execute("SELECT 1 FROM users LIMIT 1").fetchone() is None})
            if path == "/api/auth/me":
                return self.json(200, {"user": {"id": user["id"], "email": user["email"], "role": user["role"], "displayName": user["display_name"]} if user else None})
            if path.startswith("/api/relay/"):
                token = path.split("/", 3)[3]
                g = db.execute("SELECT s.*, c.name AS campaign_name FROM sessions s LEFT JOIN campaigns c ON c.id=s.campaign_id WHERE s.relay_token_hash=?", (relay_token_hash(token),)).fetchone()
                if not g: return self.json(404, {"error": "Lien relais introuvable ou révoqué."})
                pc = db.execute("SELECT COUNT(*) FROM participants WHERE session_id=?", (g["id"],)).fetchone()[0]
                cc = db.execute("SELECT COUNT(*) FROM participants WHERE session_id=? AND status='completed'", (g["id"],)).fetchone()[0]
                link = f"{self.headers.get('X-Forwarded-Proto','http')}://{self.headers.get('Host','')}/?session={g['id']}"
                return self.json(200, {"campaignName": g["campaign_name"], "groupName": g["name"], "relayName": g["relay_name"], "groupCode": g["group_code"], "groupColor": g["group_color"], "expectedParticipants": g["expected_participants"], "participantCount": pc, "completedCount": cc, "participantLink": link})
            if path == "/api/campaigns":
                if user["role"] == "admin": return self.json(200, rows(db, "SELECT * FROM campaigns ORDER BY created_at DESC"))
                return self.json(200, rows(db, "SELECT * FROM campaigns WHERE owner_user_id=? ORDER BY created_at DESC", (user["id"],)))
            if path.startswith("/api/campaigns/") and path.endswith("/groups"):
                cid = path.split("/")[3]
                return self.json(200, rows(db, """SELECT s.id,s.name,s.organization,s.location,s.date,s.status,s.created_at,s.expected_participants,s.campaign_id,s.group_code,s.group_color,s.relay_name,s.owner_user_id,
                    (SELECT COUNT(*) FROM participants p WHERE p.session_id=s.id) AS participant_count,
                    (SELECT COUNT(*) FROM participants p WHERE p.session_id=s.id AND p.status='completed') AS completed_count
                    FROM sessions s WHERE s.campaign_id=? ORDER BY s.created_at""", (cid,)))
            if path.startswith("/api/campaigns/") and path.endswith("/deletion-summary"):
                s = campaign_deletion_summary(db, path.split("/")[3])
                return self.json(200, s) if s else self.json(404, {"error": "Campagne introuvable."})
            if path.startswith("/api/campaigns/") and path.endswith("/kits.zip"):
                cid = path.split("/")[3]
                base_url = f"{self.headers.get('X-Forwarded-Proto','http')}://{self.headers.get('Host','')}"
                try:
                    data = campaign_kits_zip(db, cid, base_url)
                except ValueError as e:
                    return self.json(404, {"error": str(e)})
                name = export_filename(db.execute("SELECT name FROM campaigns WHERE id=?", (cid,)).fetchone()["name"], "kits-relais", ext="zip")
                self.send_response(200); self.send_header("Content-Type", "application/zip"); self.send_header("Content-Disposition", f"attachment; filename={name}"); self.send_header("Content-Length", str(len(data))); self.end_headers(); self.wfile.write(data); return
            if path.startswith("/api/campaigns/"):
                cid = path.rsplit("/", 1)[1]
                camp = db.execute("SELECT * FROM campaigns WHERE id=?", (cid,)).fetchone()
                return self.json(200, dict(camp) if camp else {"error": "Campagne introuvable"})
            if path == "/api/templates":
                if user["role"] == "admin": return self.json(200, rows(db, "SELECT id,name,version,description,status,model_key,is_canonical,created_at,updated_at FROM templates WHERE status='active' ORDER BY name,version DESC"))
                return self.json(200, rows(db, "SELECT id,name,version,description,status,model_key,is_canonical,created_at,updated_at FROM templates WHERE status='active' AND (owner_user_id IS NULL OR owner_user_id=?) ORDER BY name,version DESC", (user["id"],)))
            if path == "/api/templates/matrix.xlsx":
                data=blank_matrix_xlsx(); name=export_filename("matrice-questionnaire-vierge", ext="xlsx"); self.send_response(200); self.send_header("Content-Type","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"); self.send_header("Content-Disposition",f"attachment; filename={name}"); self.send_header("Content-Length",str(len(data))); self.end_headers(); self.wfile.write(data); return
            if path.startswith("/api/templates/") and path.endswith("/matrix.xlsx"):
                template=template_payload(db,path.split("/")[3]); data=matrix_xlsx(template); name=export_filename(template["name"],"matrice-questionnaire", ext="xlsx"); self.send_response(200); self.send_header("Content-Type","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"); self.send_header("Content-Disposition",f"attachment; filename={name}"); self.send_header("Content-Length",str(len(data))); self.end_headers(); self.wfile.write(data); return
            if path.startswith("/api/templates/"): return self.json(200, template_payload(db, path.rsplit("/", 1)[1]) or {"error": "Configuration introuvable"})
            if path == "/api/profile-schemas":
                if user["role"] == "admin": return self.json(200, rows(db, "SELECT id,name,description,owner_user_id,created_at,updated_at FROM profile_schemas ORDER BY name"))
                return self.json(200, rows(db, "SELECT id,name,description,owner_user_id,created_at,updated_at FROM profile_schemas WHERE owner_user_id IS NULL OR owner_user_id=? ORDER BY name", (user["id"],)))
            if path.startswith("/api/profile-schemas/"): return self.json(200, profile_schema_payload(db, path.rsplit("/", 1)[1]) or {"error": "Profil introuvable"})
            if path == "/api/sessions":
                sessions = rows(db, "SELECT * FROM sessions ORDER BY created_at DESC") if user["role"] == "admin" else rows(db, "SELECT * FROM sessions WHERE owner_user_id=? ORDER BY created_at DESC", (user["id"],))
                # profile_schema_id resolved via the campaign for group sessions
                # (correctifs cibles :8820) - every frontend consumer of this list
                # (Configuration's "profil actif" card and its edit actions,
                # participants roster, etc.) must act on the campaign's shared
                # schema, never a possibly-stale value recorded on this one group.
                for s in sessions:
                    s["profile_schema_id"] = resolve_session_profile_schema_id(db, s["id"])
                return self.json(200, sessions)
            if path.startswith("/api/sessions/") and path.endswith("/participants"):
                return self.json(200, list_session_participants(db, path.split("/")[3]))
            if path.startswith("/api/sessions/") and path.endswith("/dimensions"):
                return self.json(200, available_dimensions(db, path.split("/")[3]))
            if path.startswith("/api/sessions/") and path.endswith("/analysis"):
                sid = path.split("/")[3]
                dimension, values = query.get("dimension", [None])[0], query.get("value", [])
                filter_params = query.get("filter", [])
                if filter_params:
                    # Combinable multi-dimension filtering: each ?filter=<field_key>:<v1,v2,...>
                    # is ANDed with the others (OR between the values of one filter) - see
                    # filtered_analysis()'s docstring. Additive alongside ?dimension=/?value=,
                    # which keeps comparing several values of a SINGLE dimension unchanged.
                    filters = {}
                    for raw in filter_params:
                        field_key, sep, values_part = raw.partition(":")
                        if not sep:
                            return self.json(400, {"error": "Paramètre filter invalide (attendu field_key:valeur[,valeur...])"})
                        filters[field_key] = values_part.split(",")
                    try:
                        result = filtered_analysis(db, sid, filters)
                    except ValueError as e:
                        return self.json(400, {"error": str(e)})
                    return self.json(200, result or {"error": "Session introuvable"})
                if dimension is not None:
                    # One or more ?value= params compare that many cohorts of the
                    # same dimension in a single call (see dimension_analysis_multi's
                    # docstring) - always returns {"results": [...]}, even for a
                    # single value, so callers don't need two response shapes.
                    try:
                        results = dimension_analysis_multi(db, sid, dimension, values)
                        # Findings base on the whole, unfiltered session (never on a
                        # single cohort) so "forces"/"fragilites" stay comparable across
                        # calls; "comparison" additionally flags capacity gaps between
                        # the compared cohorts themselves (see objective_findings()).
                        findings = objective_findings(analysis(db, sid), comparison=results)
                        return self.json(200, {"results": results, "findings": findings})
                    except ValueError as e:
                        return self.json(400, {"error": str(e)})
                result = analysis(db, sid)
                return self.json(200, result or {"error": "Session introuvable"})
            if path.startswith("/api/sessions/") and path.endswith("/workshop-data"):
                sid=path.split("/")[3]; return self.json(200,{"priorities":rows(db,"SELECT * FROM priorities WHERE session_id=?",(sid,)),"notes":rows(db,"SELECT * FROM analysis_notes WHERE session_id=?",(sid,)),"recommendations":rows(db,"SELECT * FROM recommendations WHERE session_id=?",(sid,))})
            if path.startswith("/api/sessions/") and path.endswith("/qualitative-data"):
                return self.json(200, qualitative_data(db, path.split("/")[3]))
            if path.startswith("/api/sessions/") and path.endswith("/report-data"):
                return self.json(200, report_data(db, path.split("/")[3]) or {"error":"Session introuvable"})
            if path.startswith("/api/sessions/") and path.endswith("/export.json"):
                sid = path.split("/")[3]; return self.json(200, {"report":report_data(db,sid), "responses": rows(db, "SELECT * FROM responses WHERE session_id=?", (sid,)), "analysisNotes": rows(db, "SELECT * FROM analysis_notes WHERE session_id=?", (sid,))})
            if path.startswith("/api/sessions/") and path.endswith("/report.xlsx"):
                sid=path.split("/")[3];data=report_xlsx(db,sid);mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';name=export_filename(session_label(db,sid),"diagnostic",ext="xlsx")
            elif path.startswith("/api/sessions/") and path.endswith("/report.docx"):
                sid=path.split("/")[3];data=report_docx(db,sid);mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document';name=export_filename(session_label(db,sid),"rapport",ext="docx")
            elif path.startswith("/api/sessions/") and path.endswith("/individual-responses.xlsx"):
                sid=path.split("/")[3];data=individual_responses_xlsx(db,sid);mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';name=export_filename(session_label(db,sid),"reponses-individuelles",ext="xlsx")
            elif path.startswith("/api/sessions/") and path.endswith("/individual-responses.csv"):
                sid=path.split("/")[3];data=individual_responses_csv(db,sid);mime='text/csv; charset=utf-8';name=export_filename(session_label(db,sid),"reponses-individuelles",ext="csv")
            elif path.startswith("/api/sessions/") and (path.endswith("/filtered-analysis.xlsx") or path.endswith("/filtered-analysis.csv")):
                sid=path.split("/")[3]
                filters={}
                for raw in query.get("filter", []):
                    field_key, sep, values_part = raw.partition(":")
                    if not sep:
                        return self.json(400, {"error": "Paramètre filter invalide (attendu field_key:valeur[,valeur...])"})
                    filters[field_key]=values_part.split(",")
                if path.endswith(".xlsx"):
                    data=filtered_analysis_xlsx(db,sid,filters);mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';name=export_filename(session_label(db,sid),"analyse-filtree",ext="xlsx")
                else:
                    data=filtered_analysis_csv(db,sid,filters);mime='text/csv; charset=utf-8';name=export_filename(session_label(db,sid),"analyse-filtree",ext="csv")
            else: data=None
            if data is not None:
                self.send_response(200);self.send_header('Content-Type',mime);self.send_header('Content-Disposition',f'attachment; filename={name}');self.send_header('Content-Length',str(len(data)));self.send_header('Cache-Control','no-store');self.end_headers();self.wfile.write(data);return
            if path.startswith("/api/sessions/") and path.endswith("/responses.csv"):
                sid = path.split("/")[3]; buf = StringIO(); writer = csv.writer(buf); writer.writerow(["participant", "nom / organisation", "indicator", "value", "updated_at"])
                for r in db.execute("SELECT p.anonymous_id,p.display_name,i.code,r.value_json,r.updated_at FROM responses r JOIN participants p ON p.id=r.participant_id JOIN indicators i ON i.id=r.indicator_id WHERE r.session_id=?", (sid,)): writer.writerow(r)
                data = buf.getvalue().encode(); name=export_filename(session_label(db,sid),"reponses",ext="csv"); self.send_response(200); self.send_header("Content-Type", "text/csv; charset=utf-8"); self.send_header("Content-Disposition", f"attachment; filename={name}"); self.end_headers(); self.wfile.write(data); return
            if path.startswith("/api/sessions/") and path.endswith("/ai/report-blocks"):
                sid = path.split("/")[3]
                return self.json(200, rows(db, "SELECT section_key,content,retained_at FROM report_ai_blocks WHERE session_id=?", (sid,)))
            if path == "/api/ai/config":
                cfg = get_ai_config(db)
                return self.json(200, {"enabled": cfg["enabled"], "provider": cfg["provider"], "model": cfg["model"], "keyConfigured": bool(cfg["api_key"]),
                    "providers": {k: {"label": v["label"], "pricing": v["pricing"], "models": v["models"], "keyUrl": v["key_url"]} for k, v in AI_PROVIDERS.items()}})
            if path == "/api/participant":
                sid, pid = query.get("session", [None])[0], query.get("participant", [None])[0]
                return self.json(200, participant_resume(db, sid, pid))
            return self.serve_static(path)
        except AuthRequiredError: return self.json(401, {"error": "Connexion requise."})
        except PermissionDeniedError: return self.json(403, {"error": "Accès refusé : cette ressource ne vous appartient pas."})
        except ValueError as e: return self.json(404, {"error": str(e)})
        finally: db.close()

    def do_POST(self):
        path = urlparse(self.path).path
        if path == "/api/templates/import/preview":
            guard_db = self.db()
            try:
                self.require_auth(path, guard_db)
            except AuthRequiredError:
                return self.json(401, {"error": "Connexion requise."})
            finally:
                guard_db.close()
            raw=self.raw_body(); marker=b"\r\n\r\n"; start=raw.find(marker)+len(marker); end=raw.rfind(b"\r\n--"); uploaded=raw[start:end]
            try:
                preview=import_preview(uploaded); token=str(uuid.uuid4()); IMPORTS[token]=preview; return self.json(200,{"token":token,**preview})
            except ValueError as e: return self.json(400,{"error":str(e)})
        data, db = self.body(), self.db()
        try:
            user = self.require_auth(path, db)
            if path == "/api/auth/setup":
                if db.execute("SELECT 1 FROM users LIMIT 1").fetchone() is not None:
                    return self.json(409, {"error": "Un compte existe déjà : utilisez la connexion."})
                email = (data.get("email") or "").strip().lower()
                password = data.get("password") or ""
                if not email or "@" not in email: return self.json(400, {"error": "Adresse email invalide."})
                if len(password) < 8: return self.json(400, {"error": "Le mot de passe doit contenir au moins 8 caractères."})
                uid = str(uuid.uuid4()); digest, salt = hash_password(password)
                db.execute("INSERT INTO users VALUES (?,?,?,?,?,?,?)", (uid, email, digest, salt, "admin", data.get("displayName") or "", now()))
                db.commit()
                migrate_v2_ownership(db)
                token = create_auth_token(db, uid)
                return self.json(201, {"user": {"id": uid, "email": email, "role": "admin"}}, cookie=session_cookie_header(token))
            if path == "/api/auth/login":
                email = (data.get("email") or "").strip().lower()
                row = db.execute("SELECT * FROM users WHERE email=?", (email,)).fetchone()
                if not row or not verify_password(data.get("password") or "", row["password_hash"], row["password_salt"]):
                    return self.json(401, {"error": "Identifiants incorrects."})
                token = create_auth_token(db, row["id"])
                return self.json(200, {"user": {"id": row["id"], "email": row["email"], "role": row["role"], "displayName": row["display_name"]}}, cookie=session_cookie_header(token))
            if path == "/api/auth/logout":
                raw = self.headers.get("Cookie")
                if raw:
                    jar = SimpleCookie(); jar.load(raw); morsel = jar.get("epc_session")
                    if morsel: db.execute("DELETE FROM auth_tokens WHERE token_hash=?", (hashlib.sha256(morsel.value.encode("utf-8")).hexdigest(),)); db.commit()
                return self.json(200, {"ok": True}, cookie=session_cookie_header(clear=True))
            if path == "/api/campaigns":
                if not (data.get("name") or "").strip(): return self.json(400, {"error": "Le nom de la campagne est obligatoire."})
                if not data.get("templateId"): return self.json(400, {"error": "Le questionnaire est obligatoire."})
                tpl = db.execute("SELECT version FROM templates WHERE id=?", (data["templateId"],)).fetchone()
                if not tpl: return self.json(404, {"error": "Questionnaire introuvable."})
                return self.json(201, {"id": create_campaign(db, user["id"], data["templateId"], tpl["version"], data)})
            if path.startswith("/api/campaigns/") and path.endswith("/groups"):
                cid = path.split("/")[3]
                camp = db.execute("SELECT * FROM campaigns WHERE id=?", (cid,)).fetchone()
                if not camp: return self.json(404, {"error": "Campagne introuvable."})
                if not (data.get("name") or "").strip(): return self.json(400, {"error": "Le nom du groupe est obligatoire."})
                return self.json(201, create_group(db, camp, user["id"], data))
            if path.startswith("/api/campaigns/") and path.endswith("/consolidate"):
                cid = path.split("/")[3]
                session_ids = data.get("sessionIds") or []
                if len(session_ids) < 1: return self.json(400, {"error": "Sélectionnez au moins un groupe."})
                placeholders = ",".join("?" * len(session_ids))
                sel = rows(db, f"SELECT id,template_id,template_version,campaign_id FROM sessions WHERE id IN ({placeholders})", session_ids)
                if len(sel) != len(session_ids) or any(s["campaign_id"] != cid for s in sel):
                    return self.json(400, {"error": "Sélection de groupes invalide."})
                templates_used = {(s["template_id"], s["template_version"]) for s in sel}
                if len(templates_used) > 1:
                    return self.json(409, {"error": "Ces groupes utilisent des questionnaires différents et ne peuvent pas être consolidés directement."})
                result = analysis_for(db, session_ids)
                result["groups"] = rows(db, f"""SELECT id,name,group_code,group_color,expected_participants,
                    (SELECT COUNT(*) FROM participants p WHERE p.session_id=sessions.id) AS participant_count,
                    (SELECT COUNT(*) FROM participants p WHERE p.session_id=sessions.id AND p.status='completed') AS completed_count
                    FROM sessions WHERE id IN ({placeholders})""", session_ids)
                return self.json(200, result)
            if path.startswith("/api/campaigns/") and "/groups/" in path and path.endswith("/regenerate-relay"):
                parts = path.split("/"); sid = parts[5]
                g = db.execute("SELECT id FROM sessions WHERE id=? AND campaign_id=?", (sid, parts[3])).fetchone()
                if not g: return self.json(404, {"error": "Groupe introuvable."})
                return self.json(200, {"relayToken": regenerate_group_relay(db, sid)})
            if path.startswith("/api/relay/") and path.endswith("/regenerate"):
                token = path.split("/")[3]
                g = db.execute("SELECT id, campaign_id FROM sessions WHERE relay_token_hash=?", (relay_token_hash(token),)).fetchone()
                if not g: return self.json(404, {"error": "Lien relais introuvable."})
                camp = db.execute("SELECT owner_user_id FROM campaigns WHERE id=?", (g["campaign_id"],)).fetchone()
                if not camp or (user["role"] != "admin" and camp["owner_user_id"] != user["id"]): raise PermissionDeniedError()
                new_token = regenerate_group_relay(db, g["id"])
                return self.json(200, {"relayToken": new_token})
            if path == "/api/ai/test":
                try:
                    cfg = require_ai(db)
                    t0 = time.time()
                    sample = generate_ai_response(cfg["provider"], cfg["model"], "Réponds uniquement par le mot OK.", "Confirme que la connexion fonctionne.", cfg["api_key"])
                    return self.json(200, {"ok": True, "provider": AI_PROVIDERS[cfg["provider"]]["label"], "model": cfg["model"], "latencyMs": int((time.time()-t0)*1000)})
                except AIError as e: return self.json(200, {"ok": False, "reason": str(e)})

            if path.startswith("/api/sessions/") and path.endswith("/ai/diagnostic"):
                sid = path.split("/")[3]
                try:
                    cfg = require_ai(db)
                    context = ai_epc_context(db, sid)
                    system = session_restitution_manifest(db, sid)["aiSystemPrompt"] + (" Analyse conjointement capacité et consensus pour les domaines remarquables "
                        "(faible capacité + consensus élevé = constat partagé par le groupe ; faible capacité + consensus faible = perceptions divergentes ; "
                        "capacité élevée + consensus élevé = force reconnue ; capacité élevée + consensus faible = expérience hétérogène). "
                        "Structure ta réponse avec exactement ces titres, en majuscules : POINTS SAILLANTS / POINTS DE VIGILANCE / "
                        "CONTRASTES CAPACITÉ / CONSENSUS / QUESTIONS À APPROFONDIR AVEC LE GROUPE. Reste concis et exploitable en atelier.")
                    text = generate_ai_response(cfg["provider"], cfg["model"], system, context, cfg["api_key"])
                    sug_id = log_ai_suggestion(db, sid, "diagnostic", None, cfg["provider"], cfg["model"])
                    return self.json(200, {"id": sug_id, "text": text})
                except AIError as e: return self.json(409, {"error": str(e)})

            if path.startswith("/api/sessions/") and "/ai/priority/" in path and path.endswith("/prepare"):
                parts = path.split("/"); sid, pid = parts[3], parts[6]
                try:
                    cfg = require_ai(db)
                    _, _, _, context = ai_priority_context(db, sid, pid)
                    system = session_restitution_manifest(db, sid)["aiSystemPrompt"] + (" Pour cette priorité, propose uniquement : 1) un constat reformulé à partir des seules données fournies, "
                        "2) des questions à poser au groupe, 3) les points nécessitant clarification. Ne propose ni cause ni recommandation ici.")
                    text = generate_ai_response(cfg["provider"], cfg["model"], system, context, cfg["api_key"])
                    sug_id = log_ai_suggestion(db, sid, "priority_prepare", pid, cfg["provider"], cfg["model"])
                    return self.json(200, {"id": sug_id, "text": text})
                except AIError as e: return self.json(409, {"error": str(e)})

            if path.startswith("/api/sessions/") and "/ai/priority/" in path and path.endswith("/entries"):
                parts = path.split("/"); sid, pid = parts[3], parts[6]
                kind = data.get("kind")
                if kind not in ("cause", "consequence", "lever"): return self.json(400, {"error": "Type d'hypothèse invalide."})
                try:
                    cfg = require_ai(db)
                    _, _, _, context = ai_priority_context(db, sid, pid)
                    kind_fr = {"cause": "des hypothèses de CAUSES", "consequence": "des CONSÉQUENCES possibles", "lever": "des LEVIERS d'action possibles"}[kind]
                    system = session_restitution_manifest(db, sid)["aiSystemPrompt"] + (f" Propose 3 à 5 {kind_fr} pour cette priorité, formulées comme des hypothèses à discuter — "
                        "jamais présentées comme établies (n'écris jamais « cause identifiée »). "
                        "Réponds uniquement par une liste, une hypothèse par ligne, sans numérotation ni tiret, sans autre texte.")
                    text = generate_ai_response(cfg["provider"], cfg["model"], system, context, cfg["api_key"])
                    items = [l.strip("-•* ").strip() for l in text.split("\n") if l.strip()]
                    sug_id = log_ai_suggestion(db, sid, f"priority_{kind}", pid, cfg["provider"], cfg["model"])
                    return self.json(200, {"id": sug_id, "items": items})
                except AIError as e: return self.json(409, {"error": str(e)})

            if path.startswith("/api/sessions/") and path.endswith("/ai/recommendations"):
                sid = path.split("/")[3]
                try:
                    cfg = require_ai(db)
                    q = qualitative_data(db, sid)
                    retained = [e for e in q["entries"] if e["validation_status"] == "RETENU"]
                    chain_lines = []
                    for p in q["priorities"]:
                        an = next((x for x in q["analyses"] if x["priority_id"] == p["id"]), None)
                        causes = [e["content"] for e in retained if e["priority_id"] == p["id"] and e["kind"] == "cause"]
                        levers = [e["content"] for e in retained if e["priority_id"] == p["id"] and e["kind"] == "lever"]
                        if not causes and not levers: continue
                        chain_lines.append(f"Priorité (id={p['id']}) : {p['domain_label']} — {p['indicator_label']}\n"
                            f"Constat : {an['problem'] if an and an['problem'] else 'non renseigné'}\n"
                            f"Causes validées : {'; '.join(causes) or 'aucune'}\nLeviers validés : {'; '.join(levers) or 'aucun'}")
                    if not chain_lines:
                        return self.json(200, {"id": None, "items": [], "note": "Aucune cause ni levier validé par le groupe pour l'instant : rien à proposer."})
                    context = "\n\n".join(chain_lines)
                    system = session_restitution_manifest(db, sid)["aiSystemPrompt"] + (" Pour chaque priorité listée, propose une ou deux recommandations d'action fondées UNIQUEMENT "
                        "sur les causes/leviers validés fournis — ne produis pas de liste générique de bonnes pratiques sans lien avec ces données. "
                        "Réponds en JSON strict, uniquement une liste d'objets avec ces clés : priorityId, title, description, category, "
                        "priorityLevel (Haute, Moyenne ou Basse), owner, horizon. Aucun texte hors JSON.")
                    text = generate_ai_response(cfg["provider"], cfg["model"], system, context, cfg["api_key"])
                    items = _parse_json_list(text)
                    sug_id = log_ai_suggestion(db, sid, "recommendation", None, cfg["provider"], cfg["model"])
                    return self.json(200, {"id": sug_id, "items": items})
                except AIError as e: return self.json(409, {"error": str(e)})

            if path.startswith("/api/sessions/") and path.endswith("/ai/training"):
                sid = path.split("/")[3]
                try:
                    cfg = require_ai(db)
                    q = qualitative_data(db, sid)
                    recs = q["recommendations"]
                    if not recs:
                        return self.json(200, {"id": None, "items": [], "note": "Aucune recommandation disponible : rien à proposer."})
                    context = "\n".join(f"- {r['title']} : {r['description']} (catégorie {r['category']})" for r in recs)
                    system = session_restitution_manifest(db, sid)["aiSystemPrompt"] + (" À partir de ces recommandations, identifie les besoins de formation qu'elles font apparaître. "
                        "Réponds en JSON strict, uniquement une liste d'objets avec ces clés : title, targetAudience, needText, "
                        "priorityLevel (Haute, Moyenne ou Basse). Aucun texte hors JSON.")
                    text = generate_ai_response(cfg["provider"], cfg["model"], system, context, cfg["api_key"])
                    items = _parse_json_list(text)
                    sug_id = log_ai_suggestion(db, sid, "training", None, cfg["provider"], cfg["model"])
                    return self.json(200, {"id": sug_id, "items": items})
                except AIError as e: return self.json(409, {"error": str(e)})

            if path.startswith("/api/sessions/") and path.endswith("/ai/plan"):
                sid = path.split("/")[3]
                try:
                    cfg = require_ai(db)
                    q = qualitative_data(db, sid)
                    retained = [r for r in q["recommendations"] if r["status"] == "Retenue"]
                    if not retained:
                        return self.json(200, {"id": None, "items": [], "note": "Aucune recommandation retenue pour l'instant : rien à structurer."})
                    context = "\n".join(f"- {r['title']} : {r['description']} (responsable : {r['owner'] or 'non renseigné'}, "
                        f"échéance : {r['horizon'] or 'non renseignée'})" for r in retained)
                    system = session_restitution_manifest(db, sid)["aiSystemPrompt"] + (" Structure ces recommandations retenues en plan d'action. N'invente aucun engagement organisationnel : "
                        "si un responsable ou une échéance ne sont pas fournis dans les données, indique-le explicitement plutôt que d'en inventer un. "
                        "Réponds en JSON strict, uniquement une liste d'objets avec ces clés : action, owner, horizon, expectedResult, "
                        "indicator, dependencies. Aucun texte hors JSON.")
                    text = generate_ai_response(cfg["provider"], cfg["model"], system, context, cfg["api_key"])
                    items = _parse_json_list(text)
                    sug_id = log_ai_suggestion(db, sid, "plan", None, cfg["provider"], cfg["model"])
                    return self.json(200, {"id": sug_id, "items": items})
                except AIError as e: return self.json(409, {"error": str(e)})

            if path.startswith("/api/sessions/") and path.endswith("/ai/report/section"):
                sid = path.split("/")[3]
                section = data.get("section")
                manifest = session_restitution_manifest(db, sid)
                if section not in manifest["aiReportSections"]: return self.json(400, {"error": "Section de rapport invalide."})
                try:
                    cfg = require_ai(db)
                    context = ai_report_context(db, sid)
                    system = manifest["aiSystemPrompt"] + " " + manifest["aiReportSections"][section] + (" N'invente aucune étape non réalisée : si les données "
                        "nécessaires sont absentes, indique-le sobrement plutôt que d'inventer un contenu.")
                    text = generate_ai_response(cfg["provider"], cfg["model"], system, context, cfg["api_key"])
                    sug_id = log_ai_suggestion(db, sid, f"report_{section}", None, cfg["provider"], cfg["model"])
                    return self.json(200, {"id": sug_id, "section": section, "text": text})
                except AIError as e: return self.json(409, {"error": str(e)})

            if path.startswith("/api/sessions/") and path.endswith("/ai/report/full"):
                sid = path.split("/")[3]
                try:
                    cfg = require_ai(db)
                    context = ai_report_context(db, sid)
                    manifest = session_restitution_manifest(db, sid)
                    results = {}
                    for section, instruction in manifest["aiReportSections"].items():
                        system = manifest["aiSystemPrompt"] + " " + instruction + (" N'invente aucune étape non réalisée : si les données "
                            "nécessaires sont absentes, indique-le sobrement plutôt que d'inventer un contenu.")
                        results[section] = generate_ai_response(cfg["provider"], cfg["model"], system, context, cfg["api_key"])
                    sug_id = log_ai_suggestion(db, sid, "report_full", None, cfg["provider"], cfg["model"])
                    return self.json(200, {"id": sug_id, "sections": results})
                except AIError as e: return self.json(409, {"error": str(e)})

            if path == "/api/templates": return self.json(201,{"id":create_blank_template(db,data,user["id"])})
            if path == "/api/templates/import/confirm":
                preview=IMPORTS.pop(data.get("token"),None)
                if not preview: return self.json(400,{"error":"Aperçu d'import introuvable ou expiré"})
                return self.json(201,{"id":save_import(db,preview,user["id"])})
            if path.startswith("/api/templates/") and path.endswith("/clone"):
                return self.json(201,{"id":clone_template(db,path.split("/")[3],data.get("name"))})
            if path.startswith("/api/templates/") and path.endswith("/domains"):
                tid=path.split("/")[3]; return self.json(201,{"id":create_domain(db, tid, data)})
            if path.startswith("/api/domains/") and path.endswith("/indicators"):
                did=path.split("/")[3]; return self.json(201,{"id":create_indicator(db, did, data)})
            if path == "/api/profile-schemas": return self.json(201,{"id":create_profile_schema(db,user["id"],data)})
            if path.startswith("/api/profile-schemas/") and path.endswith("/fields"):
                return self.json(201,{"id":create_profile_field(db,path.split("/")[3],data)})
            if path == "/api/sessions":
                sid = create_session(db, user["id"], data)
                if sid is None: return self.json(400,{"error":"Impossible de créer une session : le questionnaire ne contient aucun domaine avec question."})
                return self.json(201, {"id": sid})
            if path.endswith("/participants"):
                sid = path.split("/")[3]
                try:
                    return self.json(201, create_participant(db, sid, data))
                except CollecteClosedError:
                    return self.json(409, {"error": "Collecte fermée"})
            if path.endswith("/responses"):
                sid = path.split("/")[3]; submit_response(db, sid, data); return self.json(200, {"ok": True})
            if path.endswith("/complete"):
                complete_participant(db, data["participantId"]); return self.json(200, {"ok": True})
            if path.startswith("/api/participants/") and path.endswith("/profile"):
                pid=path.split("/")[3]; set_participant_profile_values(db, pid, data.get("values") or {}); return self.json(200, {"ok": True})
            if path.endswith("/status"):
                sid = path.split("/")[3]; status = data["status"]; db.execute("UPDATE sessions SET status=?,closed_at=? WHERE id=?", (status, now() if status == "closed" else None, sid)); db.commit(); return self.json(200, {"ok": True})
            if path.endswith("/priorities"):
                sid = path.split("/")[3]; toggle_priority(db, sid, data); return self.json(200, {"ok": True})
            if path.endswith("/priority-analyses"):
                sid=path.split("/")[3]; upsert_priority_analysis(db, sid, data); return self.json(201,{"ok":True})
            if path.endswith("/analysis-entries"):
                sid=path.split("/")[3]; return self.json(201,{"id":create_analysis_entry(db, sid, data)})
            if path.endswith("/recommendations-v2"):
                sid=path.split("/")[3]; return self.json(201,{"id":create_workshop_recommendation(db, sid, data)})
            if path.endswith("/training-topics"):
                sid=path.split("/")[3]; return self.json(201,{"id":create_training_topic(db, sid, data)})
            if path.endswith("/report-meta"):
                sid=path.split("/")[3]; upsert_report_meta(db, sid, data); return self.json(200,{"ok":True})
            if path.endswith("/analysis-notes"):
                sid = path.split("/")[3]; create_analysis_note(db, sid, data); return self.json(201, {"ok": True})
            if path.endswith("/recommendations"):
                sid=path.split("/")[3]; create_legacy_recommendation(db, sid, data); return self.json(201,{"ok":True})
            return self.json(404, {"error": "Route inconnue"})
        except (KeyError, ValueError) as e: return self.json(400, {"error": f"Requête invalide : champ manquant ou incorrect ({e})."})
        except sqlite3.IntegrityError: return self.json(409, {"error": "Action impossible : cette donnée est encore utilisée ailleurs."})
        except AuthRequiredError: return self.json(401, {"error": "Connexion requise."})
        except PermissionDeniedError: return self.json(403, {"error": "Accès refusé : cette ressource ne vous appartient pas."})
        except Exception: return self.json(500, {"error": "Erreur interne inattendue. Aucune donnée n'a été modifiée."})
        finally: db.close()

    def do_PUT(self):
        path, data, db = urlparse(self.path).path, self.body(), self.db()
        try:
            user = self.require_auth(path, db)
            if path.startswith("/api/campaigns/") and not path.endswith("/groups"):
                cid = path.rsplit("/", 1)[1]
                camp = db.execute("SELECT * FROM campaigns WHERE id=?", (cid,)).fetchone()
                if not camp: return self.json(404, {"error": "Campagne introuvable."})
                update_campaign(db, camp, data); return self.json(200, {"ok": True})
            if path == "/api/ai/config":
                if data.get("provider") and data["provider"] not in AI_PROVIDERS: return self.json(400, {"error": "Fournisseur IA inconnu."})
                cur = get_ai_config(db)
                api_key = data["apiKey"] if data.get("apiKey") else cur["api_key"]
                db.execute("INSERT INTO ai_config (id,enabled,provider,model,api_key,updated_at) VALUES (1,?,?,?,?,?) ON CONFLICT(id) DO UPDATE SET enabled=excluded.enabled,provider=excluded.provider,model=excluded.model,api_key=excluded.api_key,updated_at=excluded.updated_at",
                    (int(bool(data.get("enabled"))), data.get("provider"), data.get("model"), api_key, now()))
                db.commit(); return self.json(200, {"ok": True})
            if path.startswith("/api/sessions/") and path.endswith("/ai/report-block"):
                sid = path.split("/")[3]; section = data.get("sectionKey"); content = data.get("content", "")
                if section not in session_restitution_manifest(db, sid)["aiReportSections"]: return self.json(400, {"error": "Section de rapport invalide."})
                db.execute("INSERT INTO report_ai_blocks (id,session_id,section_key,content,retained_at) VALUES (?,?,?,?,?) "
                    "ON CONFLICT(session_id,section_key) DO UPDATE SET content=excluded.content,retained_at=excluded.retained_at",
                    (str(uuid.uuid4()), sid, section, content, now())); db.commit(); return self.json(200, {"ok": True})
            if path.startswith("/api/sessions/"):
                sid=path.split("/")[3]
                if not update_session(db, sid, data): return self.json(404,{"error":"Questionnaire introuvable"})
                return self.json(200,{"ok":True})
            if path.startswith("/api/participants/"):
                pid=path.split("/")[3]; update_participant_display_name(db, pid, data.get("displayName")); return self.json(200,{"ok":True})
            if path.startswith("/api/priority-analyses/"):
                update_priority_analysis(db, path.split("/")[3], data.get("problem")); return self.json(200,{"ok":True})
            if path.startswith("/api/ai-suggestions/"):
                status = data.get("status")
                if status not in ("proposed", "modified", "retained", "rejected"): return self.json(400, {"error": "Statut invalide."})
                db.execute("UPDATE ai_suggestions SET status=?,updated_at=? WHERE id=?", (status, now(), path.split("/")[3])); db.commit(); return self.json(200, {"ok": True})
            if path.startswith("/api/analysis-entries/"):
                update_analysis_entry(db, path.split("/")[3], data); return self.json(200,{"ok":True})
            if path.startswith("/api/recommendations-v2/"):
                update_workshop_recommendation(db, path.split("/")[3], data); return self.json(200,{"ok":True})
            if path.startswith("/api/training-topics/"):
                update_training_topic(db, path.split("/")[3], data); return self.json(200,{"ok":True})
            if path.startswith("/api/templates/"):
                tid, version_created = update_template(db, path.split("/")[3], data)
                return self.json(200,{"id":tid,"versionCreated":version_created})
            if path.startswith("/api/domains/"):
                did=path.split("/")[3]; update_domain(db, did, data); return self.json(200,{"ok":True})
            if path.startswith("/api/indicators/"):
                iid=path.split("/")[3]
                if not (data.get("code") or "").strip(): return self.json(400,{"error":"La référence est obligatoire."})
                if not (data.get("label") or "").strip(): return self.json(400,{"error":"La question est obligatoire."})
                update_indicator(db, iid, data); return self.json(200,{"ok":True})
            if path.startswith("/api/profile-schemas/"):
                update_profile_schema(db, path.split("/")[3], data); return self.json(200,{"ok":True})
            if path.startswith("/api/profile-fields/"):
                update_profile_field(db, path.split("/")[3], data); return self.json(200,{"ok":True})
            return self.json(404,{"error":"Route inconnue"})
        except (KeyError, ValueError) as e: return self.json(400, {"error": f"Requête invalide : champ manquant ou incorrect ({e})."})
        except sqlite3.IntegrityError: return self.json(409, {"error": "Action impossible : cette donnée est encore utilisée ailleurs."})
        except AuthRequiredError: return self.json(401, {"error": "Connexion requise."})
        except PermissionDeniedError: return self.json(403, {"error": "Accès refusé : cette ressource ne vous appartient pas."})
        except Exception: return self.json(500, {"error": "Erreur interne inattendue. Aucune donnée n'a été modifiée."})
        finally: db.close()

    def do_DELETE(self):
        path, db=urlparse(self.path).path,self.db()
        try:
            user = self.require_auth(path, db)
            if path.startswith("/api/campaigns/") and "/groups/" in path:
                cid, sid = path.split("/")[3], path.split("/")[5]
                force = parse_qs(urlparse(self.path).query).get("force", ["0"])[0] == "1"
                deleted, used = delete_group_cascade(db, cid, sid, force=force)
                if not deleted:
                    return self.json(409, {"error": f"Suppression impossible : {used} réponse(s) déjà enregistrées pour ce groupe. Vous pouvez le clôturer plutôt que le supprimer, ou confirmer la suppression définitive.", "dependencies": used})
                return self.json(200, {"ok": True})
            if path.startswith("/api/campaigns/"):
                cid = path.rsplit("/", 1)[1]
                force = parse_qs(urlparse(self.path).query).get("force", ["0"])[0] == "1"
                deleted, used = delete_campaign_cascade(db, cid, force=force)
                if not deleted:
                    return self.json(409, {"error": f"Suppression impossible : cette campagne contient {used} réponse(s) enregistrées. Clôturez-la plutôt que de la supprimer, ou confirmez la suppression définitive.", "dependencies": used})
                return self.json(200, {"ok": True, "forced": force, "responsesDeleted": used})
            if path.startswith("/api/sessions/") and "/ai/report-block/" in path:
                sid=path.split("/")[3]; section=path.rsplit("/",1)[1]
                db.execute("DELETE FROM report_ai_blocks WHERE session_id=? AND section_key=?",(sid,section)); db.commit(); return self.json(200,{"ok":True})
            if path.startswith("/api/analysis-entries/"):
                eid=path.split("/")[3]; force = parse_qs(urlparse(self.path).query).get("force",["0"])[0] == "1"
                deleted, dependent = delete_analysis_entry(db, eid, force=force)
                if not deleted: return self.json(409,{"error":f"Cette entrée est utilisée par {dependent} recommandation(s). Confirmez la suppression.","dependencies":dependent})
                return self.json(200,{"ok":True})
            if path.startswith("/api/recommendations-v2/"):
                rid=path.split("/")[3]; force = parse_qs(urlparse(self.path).query).get("force",["0"])[0] == "1"
                deleted, dependent = delete_workshop_recommendation(db, rid, force=force)
                if not deleted: return self.json(409,{"error":f"Cette recommandation est liée à {dependent} thème(s) de formation. Confirmez la suppression.","dependencies":dependent})
                return self.json(200,{"ok":True})
            if path.startswith("/api/training-topics/"):
                delete_training_topic(db, path.split("/")[3]); return self.json(200,{"ok":True})
            if path.startswith("/api/templates/"):
                tid=path.split("/")[3]
                force = parse_qs(urlparse(self.path).query).get("force",["0"])[0] == "1"
                result = delete_template(db, tid, force=force)
                if result == "protected": return self.json(409,{"error":"Suppression impossible. EPC/SENEVAL est le modèle de référence ; dupliquez-le pour le modifier."})
                if result == "archived": return self.json(200,{"ok":True,"archived":True})
                if result == "in_use": return self.json(409,{"error":"Suppression impossible. Ce questionnaire est utilisé par une ou plusieurs sessions d’atelier. Vous pouvez le conserver, créer une nouvelle version, ou confirmer son retrait de la liste des modèles."})
                return self.json(200,{"ok":True})
            if path.startswith("/api/domains/"):
                did=path.split("/")[3]
                deleted, affected = delete_domain(db, did)
                if not deleted:
                    names=", ".join(a["name"] for a in affected)
                    return self.json(409,{"error":f"Suppression impossible : ce domaine contient des réponses dans {len(affected)} atelier(s) ({names}). Désactivez-le plutôt pour préserver l'historique.","sessions":affected})
                return self.json(200,{"ok":True})
            if path.startswith("/api/indicators/"):
                iid=path.split("/")[3]
                deleted, used = delete_indicator(db, iid)
                if not deleted:
                    return self.json(409,{"error":f"Suppression impossible : {used} réponse(s) sont déjà enregistrées pour cette question. Désactivez-la plutôt pour préserver l'historique.","dependencies":used})
                return self.json(200,{"ok":True})
            if path.startswith("/api/sessions/") and "/priorities/" in path:
                parts=path.split("/"); delete_priority(db, parts[3], parts[5]); return self.json(200,{"ok":True})
            if path.startswith("/api/sessions/") and len(path.rstrip("/").split("/")) == 4:
                sid=path.rstrip("/").split("/")[3]
                for table in SESSION_CHILD_TABLES:
                    db.execute(f"DELETE FROM {table} WHERE session_id=?",(sid,))
                db.execute("DELETE FROM sessions WHERE id=?",(sid,)); db.commit(); return self.json(200,{"ok":True})
            if path.startswith("/api/profile-schemas/"):
                result = delete_profile_schema(db, path.split("/")[3])
                if result == "in_use": return self.json(409,{"error":"Suppression impossible. Ce profil est utilisé par un atelier ou par des réponses déjà enregistrées."})
                return self.json(200,{"ok":True})
            if path.startswith("/api/profile-fields/"):
                fid=path.split("/")[3]
                deleted, used = delete_profile_field(db, fid)
                if not deleted:
                    return self.json(409,{"error":f"Suppression impossible : {used} réponse(s) sont déjà enregistrées pour ce champ.","dependencies":used})
                return self.json(200,{"ok":True})
            return self.json(404,{"error":"Route inconnue"})
        except (KeyError, ValueError) as e: return self.json(400, {"error": f"Requête invalide : champ manquant ou incorrect ({e})."})
        except sqlite3.IntegrityError: return self.json(409, {"error": "Action impossible : cette donnée est encore utilisée ailleurs."})
        except AuthRequiredError: return self.json(401, {"error": "Connexion requise."})
        except PermissionDeniedError: return self.json(403, {"error": "Accès refusé : cette ressource ne vous appartient pas."})
        except Exception: return self.json(500, {"error": "Erreur interne inattendue. Aucune donnée n'a été modifiée."})
        finally: db.close()

    def serve_static(self, path):
        requested = "index.html" if path in ("/", "") else path.lstrip("/")
        target = (STATIC / requested).resolve()
        if not str(target).startswith(str(STATIC.resolve())) or not target.is_file(): self.send_error(404); return
        self.send_response(200); self.send_header("Content-Type", "text/html; charset=utf-8" if target.suffix == ".html" else "application/javascript; charset=utf-8" if target.suffix == ".js" else "text/css; charset=utf-8"); self.send_header("Cache-Control","no-store, max-age=0"); self.end_headers(); self.wfile.write(target.read_bytes())


def main():
    db=connect(); init_db(db); db.close()
    host=os.environ.get("HOST","127.0.0.1")
    port=int(os.environ.get("PORT","8000"))
    server=ThreadingHTTPServer((host, port), Handler)
    print(f"EPC Workshop Engine: http://{host}:{port}")
    server.serve_forever()

if __name__ == "__main__": main()
