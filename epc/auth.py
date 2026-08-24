"""Authentification : hachage de mot de passe, tokens de session, cookie,
whitelist des routes API publiques, resolution de l'utilisateur courant et
controle d'ownership par ligne.

Extrait de app.py (lots 1a et 1b de la modularisation, cf.
AUDIT_MODULARISATION_8800.md). Les fonctions ci-dessous sont pures (aucune
dependance a la requete HTTP) : Handler (app.py) reste le seul a connaitre
self.headers/self.command et delegue ici via de fins wrappers.
"""
from __future__ import annotations

import hashlib
import hmac
import secrets
import sqlite3
from datetime import datetime, timezone, timedelta
from http.cookies import SimpleCookie

from .db import now

PBKDF2_ITERATIONS = 200_000
AUTH_TOKEN_TTL_DAYS = 14

PUBLIC_API_EXACT = {"/api/auth/setup-status", "/api/auth/setup", "/api/auth/login", "/api/auth/logout", "/api/auth/me", "/api/participant"}


class AuthRequiredError(Exception):
    """Raised when an /api/ route needs a logged-in user and none is present."""


class PermissionDeniedError(Exception):
    """Raised when a logged-in user tries to reach a session/template/campaign they don't own."""


def is_public_api(path: str, method: str) -> bool:
    if path in PUBLIC_API_EXACT:
        return True
    if path.startswith("/api/relay/") and method == "GET" and path.count("/") == 3:
        return True
    if path.startswith("/api/sessions/") and method == "POST" and (path.endswith("/participants") or path.endswith("/responses") or path.endswith("/complete")):
        return True
    if path.startswith("/api/participants/") and method == "POST" and path.endswith("/profile"):
        return True
    return False


def hash_password(password: str) -> tuple[str, str]:
    salt = secrets.token_hex(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), bytes.fromhex(salt), PBKDF2_ITERATIONS).hex()
    return digest, salt


def verify_password(password: str, password_hash: str, password_salt: str) -> bool:
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), bytes.fromhex(password_salt), PBKDF2_ITERATIONS).hex()
    return hmac.compare_digest(digest, password_hash)


def create_auth_token(db: sqlite3.Connection, user_id: str) -> str:
    token = secrets.token_urlsafe(32)
    token_hash = hashlib.sha256(token.encode("utf-8")).hexdigest()
    expires = (datetime.now(timezone.utc) + timedelta(days=AUTH_TOKEN_TTL_DAYS)).isoformat()
    db.execute("INSERT INTO auth_tokens VALUES (?,?,?,?)", (token_hash, user_id, now(), expires))
    db.commit()
    return token


def relay_token_hash(token: str) -> str:
    return hashlib.sha256(token.encode("utf-8")).hexdigest()


def session_cookie_header(token: str | None = None, clear: bool = False) -> str:
    # No `Secure` flag: the VPS currently serves this app over plain HTTP (no TLS
    # termination in front of it), and a Secure cookie would simply never be sent
    # back by the browser, breaking login entirely. SameSite=Lax is the practical
    # CSRF mitigation available without adding a token scheme.
    if clear:
        return "epc_session=; Path=/; HttpOnly; SameSite=Lax; Max-Age=0"
    return f"epc_session={token}; Path=/; HttpOnly; SameSite=Lax; Max-Age={AUTH_TOKEN_TTL_DAYS*86400}"


def resolve_current_user(db: sqlite3.Connection, cookie_header: str | None):
    if not cookie_header:
        return None
    jar = SimpleCookie(); jar.load(cookie_header)
    morsel = jar.get("epc_session")
    if not morsel:
        return None
    token_hash = hashlib.sha256(morsel.value.encode("utf-8")).hexdigest()
    row = db.execute("SELECT u.* FROM auth_tokens t JOIN users u ON u.id=t.user_id WHERE t.token_hash=? AND t.expires_at>?", (token_hash, now())).fetchone()
    return dict(row) if row else None


def enforce_ownership(path: str, db: sqlite3.Connection, user) -> None:
    parts = path.split("/")
    if path.startswith("/api/sessions/") and len(parts) > 3 and parts[3]:
        row = db.execute("SELECT owner_user_id FROM sessions WHERE id=?", (parts[3],)).fetchone()
        if row and user["role"] != "admin" and row["owner_user_id"] not in (None, user["id"]):
            raise PermissionDeniedError()
    elif path.startswith("/api/templates/") and len(parts) > 3 and parts[3] not in ("matrix.xlsx", "import"):
        row = db.execute("SELECT owner_user_id FROM templates WHERE id=?", (parts[3],)).fetchone()
        if row and row["owner_user_id"] is not None and user["role"] != "admin" and row["owner_user_id"] != user["id"]:
            raise PermissionDeniedError()
    elif path.startswith("/api/campaigns/") and len(parts) > 3 and parts[3]:
        row = db.execute("SELECT owner_user_id FROM campaigns WHERE id=?", (parts[3],)).fetchone()
        if row and user["role"] != "admin" and row["owner_user_id"] != user["id"]:
            raise PermissionDeniedError()
    elif path.startswith("/api/profile-schemas/") and len(parts) > 3 and parts[3]:
        row = db.execute("SELECT owner_user_id FROM profile_schemas WHERE id=?", (parts[3],)).fetchone()
        if row and row["owner_user_id"] is not None and user["role"] != "admin" and row["owner_user_id"] != user["id"]:
            raise PermissionDeniedError()
    elif path.startswith("/api/profile-fields/") and len(parts) > 3 and parts[3]:
        row = db.execute("SELECT s.owner_user_id AS owner_user_id FROM profile_fields f JOIN profile_schemas s ON s.id=f.schema_id WHERE f.id=?", (parts[3],)).fetchone()
        if row and row["owner_user_id"] is not None and user["role"] != "admin" and row["owner_user_id"] != user["id"]:
            raise PermissionDeniedError()


def resolve_auth(path: str, method: str, db: sqlite3.Connection, cookie_header: str | None):
    """Call first inside each verb handler's try block. Returns the current
    user (or None for the small public whitelist) and enforces per-row
    ownership for /api/sessions|templates|campaigns/<id>... routes."""
    user = resolve_current_user(db, cookie_header)
    if path.startswith("/api/") and not is_public_api(path, method):
        if user is None:
            raise AuthRequiredError()
        enforce_ownership(path, db, user)
    return user
