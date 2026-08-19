"""Tiny stdlib-only string helpers shared by several epc/* modules and app.py.

Extrait de app.py (lot 1d de la modularisation, cf. AUDIT_MODULARISATION_8800.md).
"""
from __future__ import annotations

import re
import unicodedata


def slugify(text: str) -> str:
    text = unicodedata.normalize("NFKD", text or "").encode("ascii", "ignore").decode("ascii")
    text = re.sub(r"[^A-Za-z0-9]+", "-", text).strip("-")
    return text or "sans-titre"
