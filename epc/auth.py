"""Primitives d'authentification : hachage de mot de passe, tokens de session,
cookie et whitelist des routes API publiques.

Extrait de app.py (lot 1a de la modularisation, cf. AUDIT_MODULARISATION_8800.md).
La resolution de l'utilisateur courant et le controle d'ownership par ligne
restent dans Handler (app.py) : ils sont couples a la requete HTTP (self) et
seront traites dans un lot dedie, avec ses propres tests.
"""
from __future__ import annotations

import hashlib
import hmac
import secrets
import sqlite3
from datetime import datetime, timezone, timedelta

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
