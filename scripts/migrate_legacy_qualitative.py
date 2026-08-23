#!/usr/bin/env python
"""Migration V1 -> V2 des donnees qualitatives (lot 7, cf. AUDIT_MODULARISATION_8800.md).

Usage :
    python scripts/migrate_legacy_qualitative.py             # dry-run (aucune ecriture)
    python scripts/migrate_legacy_qualitative.py --apply      # applique reellement la migration
    python scripts/migrate_legacy_qualitative.py --session <id> [--apply]

A executer explicitement, jamais automatiquement (aucun hook au demarrage de
l'application) - voir epc/qualitatif.py:migrate_legacy_qualitative_data()
pour le detail de ce qui est copie et pourquoi. Sans danger a relancer
plusieurs fois (idempotent) ; les anciennes tables (analysis_notes/
recommendations) ne sont jamais modifiees ni supprimees, seulement lues.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from epc.db import connect, init_db
from epc.qualitatif import migrate_legacy_qualitative_data


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--session", help="Ne migrer qu'un seul atelier (par id). Par defaut : tous les ateliers.")
    parser.add_argument("--apply", action="store_true", help="Appliquer reellement la migration (sinon : dry-run, aucune ecriture).")
    args = parser.parse_args()

    db = connect()
    init_db(db)

    if not args.apply:
        where = "WHERE session_id=?" if args.session else ""
        params = (args.session,) if args.session else ()
        notes = db.execute(f"SELECT COUNT(*) FROM analysis_notes {where}", params).fetchone()[0]
        recs = db.execute(f"SELECT COUNT(*) FROM recommendations {where}", params).fetchone()[0]
        print(f"[dry-run] {notes} note(s) V1 et {recs} recommandation(s) V1 trouvee(s) (aucune ecriture).")
        print("Relancez avec --apply pour migrer reellement vers les tables V2.")
        db.close()
        return

    result = migrate_legacy_qualitative_data(db, session_id=args.session)
    db.close()
    print("Migration appliquee :")
    for key, value in result.items():
        print(f"  {key}: {value}")


if __name__ == "__main__":
    main()
