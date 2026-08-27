"""Connexion SQLite, schema et migrations, et lecture generique des questionnaires.

Extrait de app.py (lot 1a de la modularisation, cf. AUDIT_MODULARISATION_8800.md) :
aucune formule ni aucun contrat de route ne change ici, seule l'emplacement du code
technique bouge. app.py reimporte ces symboles a l'identique.
"""
from __future__ import annotations

import json
import os
import shutil
import sqlite3
import uuid
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DATABASE = Path(os.environ["DATA_DIR"]) / "workshops.sqlite3" if os.environ.get("DATA_DIR") else ROOT / "data" / "workshops.sqlite3"


def now() -> str:
    return datetime.now(timezone.utc).isoformat()


# Questionnaire EPC/SENEVAL de reference : source de verite absolue = le document
# local "5 Diagnostic de SenEval selon EPC PAB (1).docx" (extrait mecaniquement via
# python-docx, aucune reformulation). 7 domaines x 10 indicateurs = 70. Chaque
# indicateur porte deux textes distincts du DOCX : sa Reference courte et son enonce
# complet ("Indicateurs qualitatifs ou Capacites"), stockes respectivement dans les
# colonnes indicators.code et indicators.label.
EPC_DOMAINS = [
    ('grh', 'Gestion des Ressources Humaines', [
        ('Formation au personnel', 'Nous offrons régulièrement la formation au personnel'),
        ('Priorités de notre Organisation', 'Notre formation du personnel contribue directement à l’exécution des priorités de notre Organisation'),
        ('Compétences pour notre mission', 'Notre personnel a les compétences appropriées pour exécuter notre mission'),
        ('Nombre du personnel', 'Le nombre de personnel est approprié pour exécuter notre mission'),
        ('Diversité de nos bénéficiaires', 'Notre personnel reflète la diversité de nos bénéficiaires'),
        ('Recrutement des membres', 'Le système de recrutement des membres favorise l’adhésion à l’organisation'),
        ('Evaluation du Personnel', 'L’évaluation du Personnel encourage la fixation des membres'),
        ('Résolution des conflits', 'La politique de résolution des griefs et des conflits contribue à la rétention du personnel :'),
        ('Allocation des tâches', 'L’allocation des tâches et des responsabilités participent à la rétention du personnel :'),
        ('Pratiques de supervision', 'Les pratiques de supervision améliorent la capacité de notre personnel à atteindre les objectifs de l’Organisation.'),
    ]),
    ('grf', 'Gestion des Ressources Financières', [
        ('Equilibre des recettes et des dépenses', 'Nous utilisons régulièrement les procédures établies pour maintenir nos recettes et nos dépenses équilibrées'),
        ('Allocation des fonds', 'Le processus de budgétisation nous amène à allouer des fonds d’une manière qui reflète étroitement nos priorités organisationnelles'),
        ('Prévisions financières', 'Nos prévisions financières sont exactes'),
        ('Modification des dépenses', 'Nous modifions nos dépenses sur une base régulière chaque fois que nous avons des déficits de revenus'),
        ('Evitement des perturbations', 'Notre système financier et nos procédures nous évitent des perturbations opérationnelles'),
        ('Décaissements périodiques', 'Nos procédures de gestion des liquidités conduisent à des décaissements périodiques de fonds'),
        ('Appui financier des bailleurs', 'Le niveau de l’appui financier de la part des bailleurs reste stable ou croissant'),
        ('Moins de  dépendance', 'Nous prenons des mesures concrètes pour rendre notre organisation moins dépendante de quelques sources de financement'),
        ('Ressources pour les activités', 'Le niveau de nos ressources disponibles pour les activités du projet est approprié pour accomplir notre mission'),
        ('Ressources pour l’équipement', 'Le niveau de nos ressources disponibles pour l’équipement (bureaux, fournitures) est approprié pour accomplir notre mission'),
    ]),
    ('parteq', 'Gestion de la Participation Equitable', [
        ('Evaluation des besoins', 'Les niveaux de participation dans l’évaluation des besoins sont élevés'),
        ('Conception des projets', 'Les niveaux de participation dans la conception des projets sont élevés'),
        ('Mise en œuvre des projets', 'Les niveaux de participation dans la mise en œuvre des projets sont élevés'),
        ('Suivi et évaluation des projets', 'Les niveaux de participation dans le suivi et l’évaluation des projets sont élevés'),
        ('Groupes sous-représentés et accès aux activités', 'Les groupes sous-représentés ont un accès équitable aux activités du projet'),
        ('Groupes sous-représentés et bénéfice', 'Les groupes sous-représentés tirent un bénéfice équitable des activités du projet'),
        ('Promotion de l’équité', 'Nos projets font constamment la promotion de l’équité à tous les niveaux de la conception et de la mise en œuvre des projets'),
        ('Evaluation des changements', 'Nous examinons régulièrement les besoins des participants au projet pour évaluer s’ils changent.'),
        ('Besoins changeants des participants', 'Nous modifions les projets pour refléter les besoins changeant des participants'),
        ('Dialogue pour le développement équitable', 'Nous engageons régulièrement les décideurs et les institutions pertinents dans un dialogue qui contribue au développement équitable et participatif'),
    ]),
    ('dur', 'Gestion de la Durabilité des Acquis de l’organisation', [
        ('Conception et D. environnementale', 'Au moment de la conception de nos projets, nous accordons une attention adéquate à la durabilité environnementale'),
        ('Conception et D. économique', 'A la durabilité économique'),
        ('Conception et D. institutionnelle', 'A la durabilité institutionnelle'),
        ('Mise en œuvre et D. environnementale', 'En exécutant les projets, nous accordons une attention adéquate à la durabilité environnementale'),
        ('Mise en œuvre et D. économique', 'A la durabilité économique'),
        ('Mise en œuvre et D. institutionnelle', 'A la durabilité institutionnelle'),
        ('Evaluation et D. environnementale', 'En faisant le suivi du projet et l’évaluation de l’impact, nous accordions une attention adéquate à la durabilité environnementale'),
        ('Evaluation et D. économique', 'A la durabilité économique'),
        ('Evaluation et D. institutionnelle', 'A la durabilité institutionnelle'),
        ('Appui technique et durabilité', 'La qualité de l’appui technique pour nos activités de terrain contribue à la durabilité du projet'),
    ]),
    ('partn', 'Gestion du Partenariat', [
        ('Liens avec les décideurs politiques', 'Nous établissons de nouveaux liens précieux avec les décideurs politiques pertinents'),
        ('Liens avec le secteur privé', 'Nous établissons de nouveaux liens précieux avec les représentants du secteur privé'),
        ('Partenariats avec d’autres Organisations', 'Nous nous engageons activement dans des partenariats productifs avec d’autres Organisations'),
        ('Suivi de nos partenariats', 'Nous faisons le suivi de l’efficacité de nos partenariats avec les autres Organisations'),
        ('Avantages financiers', 'A travers le partenariat nous obtenons des avantages financiers qui améliorent notre capacité pour accomplir notre mission'),
        ('Compétences techniques', 'Et aussi des compétences techniques qui améliorent notre capacité à accomplir notre mission'),
        ('Nouveaux réseaux et relations', 'Et aussi de nouveaux réseaux et des relations qui améliorent notre capacité à accomplir notre mission'),
        ('Confiance et coopération', 'Les partenariats ont des mécanismes en place pour renforcer la confiance et la coopération'),
        ('Contribution aux objectifs partagés', 'Les partenaires individuels contribuent de manière appropriée aux objectifs partagés'),
        ('Effort de coopération', 'Les partenaires individuels participent aux bénéfices de l’effort de coopération'),
    ]),
    ('apporg', 'Gestion de l’Apprentissage Organisationnel', [
        ('Evaluation des projets', 'Nous faisons régulièrement le suivi et l’évaluation de la mise en œuvre de nos projets'),
        ('Implication des structures dans les défis', 'Nous impliquons régulièrement les structures dans la satisfaction des défis organisationnels majeurs'),
        ('Interdépendance des structures', 'Nous reconnaissons l’interdépendance des différentes structures de notre Organisation lorsque nous analysons les problèmes'),
        ('Informations pour le travail', 'Les membres ont régulièrement les informations dont ils ont besoin pour faire efficacement leur travail'),
        ('Informations pour les priorités', 'Nous disposons d’informations appropriées pour répondre à nos priorités'),
        ('Travail d’équipe pour les défis', 'Nous utilisons efficacement le travail d’équipe pour répondre aux défis organisationnels'),
        ('Travail d’équipe pour les défis organisationnels', 'Les responsables utilisent efficacement la travail d’équipe pour répondre aux défis organisationnels'),
        ('Réunions et apprentissage organisationnel', 'Nos réunions de personnel contribuent directement à l’apprentissage organisationnel.'),
        ('Expression libre lors des réunions', 'Les membres se sentent généralement à l’aise pour s’exprimer lors des réunions de personnel'),
        ('Prise de risque pour les innovateurs', 'Notre Organisation est une place sûre de prise de risque pour les innovateurs'),
    ]),
    ('gouv', 'Gestion Stratégique et Gouvernance', [
        ('Rapportage pour les bailleurs', 'Notre système de rapportage pour les bailleurs de fonds montre une compréhension claire de leurs besoins et exigences'),
        ('CC et mobilisation des fonds', 'Notre Comité de coordination a contribué à l’exécution des fonctions comme la mobilisation des fonds'),
        ('CC et relations publiques', 'Comme les relations publiques'),
        ('CC et plaidoyer', 'Comme le plaidoyer'),
        ('CC et définition de politique', 'Comme la définition de politique'),
        ('Représentation des bénéficiaires', 'Notre Comité de coordination a une représentation appropriée de nos principales bénéficiaires'),
        ('Engagement et décisions prises', 'L’engagement par rapport à notre mission, à nos objectifs et à nos valeurs est systématiquement reflété dans les décisions prises par le Comité de coordination et le personnel'),
        ('Planification stratégique et extérieur', 'Nous utilisons une planification stratégique pour nous examiner nous-mêmes par rapport à notre environnement externe'),
        ('Initiatives et plans stratégiques', 'Nos initiatives sont élaborées et mises en œuvre conformément à nos plans stratégiques et opérationnels'),
        ('Suivi du progrès', 'Régulièrement nous faisons le suivi du progrès dans l’accomplissement de nos objectifs stratégiques'),
    ]),
]

GRADING = [(0, 22, 5), (23, 32, 10), (33, 39, 15), (40, 45, 20), (46, 50, 25), (51, 55, 30), (56, 59, 35), (60, 63, 40), (64, 67, 45), (68, 71, 50), (72, 74, 55), (75, 78, 60), (79, 81, 65), (82, 84, 70), (85, 87, 75), (88, 89, 80), (90, 92, 85), (93, 95, 90), (96, 98, 95), (99, 100, 100)]


def connect(path: Path = DATABASE) -> sqlite3.Connection:
    path.parent.mkdir(parents=True, exist_ok=True)
    db = sqlite3.connect(path)
    db.row_factory = sqlite3.Row
    db.execute("PRAGMA foreign_keys = ON")
    return db


def init_db(db: sqlite3.Connection) -> None:
    db.executescript("""
    CREATE TABLE IF NOT EXISTS templates (id TEXT PRIMARY KEY, name TEXT NOT NULL, version INTEGER NOT NULL, description TEXT, status TEXT NOT NULL, scale_json TEXT NOT NULL, scoring_json TEXT NOT NULL, consensus_json TEXT NOT NULL, grading_json TEXT NOT NULL, priority_json TEXT NOT NULL, created_at TEXT NOT NULL, updated_at TEXT NOT NULL, UNIQUE(name, version));
    CREATE TABLE IF NOT EXISTS domains (id TEXT PRIMARY KEY, template_id TEXT NOT NULL REFERENCES templates(id), code TEXT NOT NULL, label TEXT NOT NULL, description TEXT, display_order INTEGER NOT NULL, active INTEGER NOT NULL DEFAULT 1);
    CREATE TABLE IF NOT EXISTS indicators (id TEXT PRIMARY KEY, domain_id TEXT NOT NULL REFERENCES domains(id), code TEXT NOT NULL, label TEXT NOT NULL, description TEXT, response_type TEXT NOT NULL, required INTEGER NOT NULL DEFAULT 1, display_order INTEGER NOT NULL, active INTEGER NOT NULL DEFAULT 1, configuration_json TEXT NOT NULL DEFAULT '{}');
    CREATE TABLE IF NOT EXISTS sessions (id TEXT PRIMARY KEY, template_id TEXT NOT NULL REFERENCES templates(id), template_version INTEGER NOT NULL, name TEXT NOT NULL, organization TEXT, location TEXT, date TEXT, status TEXT NOT NULL, created_at TEXT NOT NULL, closed_at TEXT, description TEXT, expected_participants INTEGER);
    CREATE TABLE IF NOT EXISTS participants (id TEXT PRIMARY KEY, session_id TEXT NOT NULL REFERENCES sessions(id), anonymous_id TEXT NOT NULL, status TEXT NOT NULL, started_at TEXT NOT NULL, completed_at TEXT, display_name TEXT, UNIQUE(session_id, anonymous_id));
    CREATE TABLE IF NOT EXISTS responses (id TEXT PRIMARY KEY, session_id TEXT NOT NULL REFERENCES sessions(id), participant_id TEXT NOT NULL REFERENCES participants(id), indicator_id TEXT NOT NULL REFERENCES indicators(id), value_json TEXT NOT NULL, value_type TEXT NOT NULL, created_at TEXT NOT NULL, updated_at TEXT NOT NULL, UNIQUE(participant_id, indicator_id));
    CREATE TABLE IF NOT EXISTS priorities (id TEXT PRIMARY KEY, session_id TEXT NOT NULL REFERENCES sessions(id), domain_id TEXT NOT NULL REFERENCES domains(id), indicator_id TEXT NOT NULL REFERENCES indicators(id), votes INTEGER NOT NULL DEFAULT 0, selected_at TEXT NOT NULL, UNIQUE(session_id, indicator_id));
    CREATE TABLE IF NOT EXISTS analysis_notes (id TEXT PRIMARY KEY, session_id TEXT NOT NULL REFERENCES sessions(id), indicator_id TEXT REFERENCES indicators(id), kind TEXT NOT NULL, content TEXT NOT NULL, validation_status TEXT NOT NULL DEFAULT 'HYPOTHESE', created_at TEXT NOT NULL, updated_at TEXT NOT NULL);
    CREATE TABLE IF NOT EXISTS recommendations (id TEXT PRIMARY KEY, session_id TEXT NOT NULL REFERENCES sessions(id), indicator_id TEXT REFERENCES indicators(id), title TEXT NOT NULL, description TEXT, lever TEXT, kind TEXT NOT NULL, owner TEXT, horizon TEXT, created_at TEXT NOT NULL);
    CREATE TABLE IF NOT EXISTS priority_analyses (id TEXT PRIMARY KEY, session_id TEXT NOT NULL REFERENCES sessions(id), priority_id TEXT NOT NULL REFERENCES priorities(id), problem TEXT NOT NULL DEFAULT '', created_at TEXT NOT NULL, updated_at TEXT NOT NULL, UNIQUE(session_id, priority_id));
    CREATE TABLE IF NOT EXISTS analysis_entries (id TEXT PRIMARY KEY, session_id TEXT NOT NULL REFERENCES sessions(id), priority_id TEXT NOT NULL REFERENCES priorities(id), parent_id TEXT REFERENCES analysis_entries(id), kind TEXT NOT NULL, content TEXT NOT NULL, item_type TEXT, comment TEXT, validation_status TEXT NOT NULL DEFAULT 'A_DISCUTER', created_at TEXT NOT NULL, updated_at TEXT NOT NULL);
    CREATE TABLE IF NOT EXISTS workshop_recommendations (id TEXT PRIMARY KEY, session_id TEXT NOT NULL REFERENCES sessions(id), priority_id TEXT REFERENCES priorities(id), cause_id TEXT REFERENCES analysis_entries(id), lever_id TEXT REFERENCES analysis_entries(id), title TEXT NOT NULL, description TEXT NOT NULL, category TEXT NOT NULL DEFAULT 'Autre', priority_level TEXT NOT NULL DEFAULT 'Non définie', owner TEXT, horizon TEXT, comment TEXT, status TEXT NOT NULL DEFAULT 'Proposée', created_at TEXT NOT NULL, updated_at TEXT NOT NULL);
    CREATE TABLE IF NOT EXISTS training_topics (id TEXT PRIMARY KEY, session_id TEXT NOT NULL REFERENCES sessions(id), priority_id TEXT REFERENCES priorities(id), recommendation_id TEXT REFERENCES workshop_recommendations(id), title TEXT NOT NULL, need_text TEXT, target_audience TEXT, priority_level TEXT NOT NULL DEFAULT 'Non définie', comment TEXT, created_at TEXT NOT NULL, updated_at TEXT NOT NULL);
    CREATE TABLE IF NOT EXISTS session_report_meta (session_id TEXT PRIMARY KEY REFERENCES sessions(id), facilitator TEXT, audience TEXT, context TEXT, conclusion TEXT, updated_at TEXT NOT NULL);
    CREATE TABLE IF NOT EXISTS ai_config (id INTEGER PRIMARY KEY CHECK (id=1), enabled INTEGER NOT NULL DEFAULT 0, provider TEXT, model TEXT, api_key TEXT, updated_at TEXT NOT NULL);
    CREATE TABLE IF NOT EXISTS ai_suggestions (id TEXT PRIMARY KEY, session_id TEXT NOT NULL REFERENCES sessions(id), kind TEXT NOT NULL, target_id TEXT, provider TEXT NOT NULL, model TEXT NOT NULL, status TEXT NOT NULL DEFAULT 'proposed', created_at TEXT NOT NULL, updated_at TEXT NOT NULL);
    CREATE TABLE IF NOT EXISTS report_ai_blocks (id TEXT PRIMARY KEY, session_id TEXT NOT NULL REFERENCES sessions(id), section_key TEXT NOT NULL, content TEXT NOT NULL, retained_at TEXT NOT NULL, UNIQUE(session_id, section_key));
    CREATE TABLE IF NOT EXISTS users (id TEXT PRIMARY KEY, email TEXT NOT NULL UNIQUE, password_hash TEXT NOT NULL, password_salt TEXT NOT NULL, role TEXT NOT NULL DEFAULT 'admin', display_name TEXT, created_at TEXT NOT NULL);
    CREATE TABLE IF NOT EXISTS auth_tokens (token_hash TEXT PRIMARY KEY, user_id TEXT NOT NULL REFERENCES users(id), created_at TEXT NOT NULL, expires_at TEXT NOT NULL);
    CREATE TABLE IF NOT EXISTS campaigns (id TEXT PRIMARY KEY, owner_user_id TEXT NOT NULL REFERENCES users(id), name TEXT NOT NULL, description TEXT, period_start TEXT, period_end TEXT, template_id TEXT NOT NULL REFERENCES templates(id), template_version INTEGER NOT NULL, status TEXT NOT NULL DEFAULT 'active', created_at TEXT NOT NULL, updated_at TEXT NOT NULL);
    CREATE TABLE IF NOT EXISTS profile_schemas (id TEXT PRIMARY KEY, name TEXT NOT NULL, description TEXT, owner_user_id TEXT, created_at TEXT NOT NULL, updated_at TEXT NOT NULL);
    CREATE TABLE IF NOT EXISTS profile_fields (id TEXT PRIMARY KEY, schema_id TEXT NOT NULL REFERENCES profile_schemas(id), field_key TEXT NOT NULL, field_type TEXT NOT NULL, label TEXT NOT NULL, required INTEGER NOT NULL DEFAULT 0, options_json TEXT NOT NULL DEFAULT '[]', display_order INTEGER NOT NULL, active INTEGER NOT NULL DEFAULT 1, is_dimension INTEGER NOT NULL DEFAULT 0);
    CREATE TABLE IF NOT EXISTS participant_profile_values (id TEXT PRIMARY KEY, session_id TEXT NOT NULL REFERENCES sessions(id), participant_id TEXT NOT NULL REFERENCES participants(id), field_id TEXT NOT NULL REFERENCES profile_fields(id), value_json TEXT NOT NULL, created_at TEXT NOT NULL, updated_at TEXT NOT NULL, UNIQUE(participant_id, field_id));
    """)
    db.commit()
    existing_columns = {r["name"] for r in db.execute("PRAGMA table_info(participants)")}
    if "display_name" not in existing_columns:
        db.execute("ALTER TABLE participants ADD COLUMN display_name TEXT")
        db.commit()
    session_columns = {r["name"] for r in db.execute("PRAGMA table_info(sessions)")}
    for col, decl in (("description", "TEXT"), ("expected_participants", "INTEGER"), ("owner_user_id", "TEXT"), ("campaign_id", "TEXT"), ("group_code", "TEXT"), ("group_color", "TEXT"), ("relay_name", "TEXT"), ("relay_token_hash", "TEXT"), ("profile_schema_id", "TEXT")):
        if col not in session_columns:
            db.execute(f"ALTER TABLE sessions ADD COLUMN {col} {decl}")
    campaign_columns = {r["name"] for r in db.execute("PRAGMA table_info(campaigns)")}
    if "profile_schema_id" not in campaign_columns:
        # Mission de correctifs cibles :8820 (cf. consignes_claude.txt) : la
        # campagne devient la source de verite du profil pour TOUS ses groupes -
        # avant cette colonne, chaque groupe recevait son propre schema cree
        # independamment (ensure_default_profile_schema appele par groupe), ce
        # qui faisait diverger silencieusement les dimensions entre groupes
        # d'une meme campagne.
        db.execute("ALTER TABLE campaigns ADD COLUMN profile_schema_id TEXT")
    template_columns = {r["name"] for r in db.execute("PRAGMA table_info(templates)")}
    if "owner_user_id" not in template_columns:
        db.execute("ALTER TABLE templates ADD COLUMN owner_user_id TEXT")
    if "model_key" not in template_columns:
        db.execute("ALTER TABLE templates ADD COLUMN model_key TEXT")
    if "is_canonical" not in template_columns:
        db.execute("ALTER TABLE templates ADD COLUMN is_canonical INTEGER NOT NULL DEFAULT 0")
    profile_field_columns = {r["name"] for r in db.execute("PRAGMA table_info(profile_fields)")}
    if profile_field_columns and "is_dimension" not in profile_field_columns:
        db.execute("ALTER TABLE profile_fields ADD COLUMN is_dimension INTEGER NOT NULL DEFAULT 0")
    ppv_columns = {r["name"] for r in db.execute("PRAGMA table_info(participant_profile_values)")}
    if ppv_columns and "session_id" not in ppv_columns:
        # Table already existed (created before session_id was added to its schema);
        # backfill from participants so cascade deletes (SESSION_CHILD_TABLES) can
        # scope by session_id like every other participant-linked table already does.
        db.execute("ALTER TABLE participant_profile_values ADD COLUMN session_id TEXT")
        db.execute("UPDATE participant_profile_values SET session_id=(SELECT session_id FROM participants WHERE participants.id=participant_profile_values.participant_id) WHERE session_id IS NULL")
    db.commit()
    if db.execute("SELECT 1 FROM templates LIMIT 1").fetchone() is None:
        seed_epc(db)
    ensure_reference_questionnaire_version(db)
    # Must run before migrate_v2_ownership(): that function now decides by model_key
    # (see ensure_model_identity's docstring), so the flag has to be current first —
    # otherwise, on the very first call after upgrading an existing database, the
    # EPC/SENEVAL rows would still show model_key=NULL and be wrongly treated as
    # ownerless-assignable.
    ensure_model_identity(db)
    migrate_v2_ownership(db)


def ensure_reference_questionnaire_version(db: sqlite3.Connection) -> None:
    """Keep an up-to-date "EPC / SENEVAL" template version available WITHOUT ever
    touching an existing version's domains/indicators/responses.

    Runs on every startup (idempotent: a no-op once the latest version already
    matches EPC_DOMAINS). Existing workshops stay pinned to whichever template
    version they were created on (sessions.template_version), so their
    historical responses are never altered by a referential correction here —
    only a brand-new version is added, and new workshops/campaigns pick up the
    latest version by default (see consignes_claude.txt: existing ateliers must
    not be silently altered when their questionnaire is versioned; an automatic
    migration that risks changing historical responses must NOT be done).

    Finds "the latest version" by model_key OR name (lot 2c, cf.
    AUDIT_MODULARISATION_8800.md): model_key is only populated after the first
    ensure_model_identity() call, so name is still needed as a bootstrap fallback,
    but model_key is what makes this resilient to a rename via PUT
    /api/templates/{id} — matching by name alone would have missed a renamed
    latest version and silently reseeded a duplicate default "EPC / SENEVAL" v1
    (or a duplicate next version number) alongside the renamed, orphaned one.
    """
    target_codes = [code for code, _, _ in EPC_DOMAINS]
    target_counts = {code: len(indicators) for code, _, indicators in EPC_DOMAINS}
    latest = db.execute("SELECT id FROM templates WHERE model_key=? OR name='EPC / SENEVAL' ORDER BY version DESC LIMIT 1", (MODEL_KEY_EPC_SENEVAL,)).fetchone()
    if not latest:
        seed_epc(db)
        return
    tid = latest["id"]
    domains = db.execute("SELECT id,code FROM domains WHERE template_id=? ORDER BY display_order", (tid,)).fetchall()
    up_to_date = [d["code"] for d in domains] == target_codes and all(
        db.execute("SELECT COUNT(*) FROM indicators WHERE domain_id=?", (d["id"],)).fetchone()[0] == target_counts[d["code"]]
        for d in domains
    )
    if up_to_date:
        return

    old = template_payload(db, tid)
    version = db.execute("SELECT COALESCE(MAX(version),0)+1 FROM templates WHERE model_key=? OR name='EPC / SENEVAL'", (MODEL_KEY_EPC_SENEVAL,)).fetchone()[0]
    new_tid, stamp = str(uuid.uuid4()), now()
    db.execute("INSERT INTO templates (id,name,version,description,status,scale_json,scoring_json,consensus_json,grading_json,priority_json,created_at,updated_at,owner_user_id) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", (new_tid, "EPC / SENEVAL", version, old["description"], "active", json.dumps(old["scale"]), json.dumps(old["scoring"]), json.dumps(old["consensus"]), json.dumps(old["grading"]), json.dumps(old["priority"]), stamp, stamp, old.get("owner_user_id")))
    for d_order, (code, label, indicators) in enumerate(EPC_DOMAINS, 1):
        did = str(uuid.uuid4())
        db.execute("INSERT INTO domains VALUES (?,?,?,?,?,?,?)", (did, new_tid, code, label, "", d_order, 1))
        for i_order, (reference, enonce) in enumerate(indicators, 1):
            db.execute("INSERT INTO indicators VALUES (?,?,?,?,?,?,?,?,?,?)", (str(uuid.uuid4()), did, reference, enonce, "", "numeric", 1, i_order, 1, "{}"))
    db.commit()


def migrate_v2_ownership(db: sqlite3.Connection) -> None:
    """Assign every ownerless session/template to the first account created.

    Runs on every startup (idempotent: a no-op once nothing is ownerless, or
    while no account exists yet). This is what lets the pilot who completes
    the first-run setup immediately keep the workshops/questionnaires that
    already existed before authentication was introduced.

    Identifies the shared EPC/SENEVAL model(s) by model_key rather than by name
    (lot 2b, cf. AUDIT_MODULARISATION_8800.md): unlike name, model_key survives a
    rename via PUT /api/templates/{id}, so a renamed reference questionnaire stays
    correctly excluded from personal ownership. Requires ensure_model_identity()
    to have already run in this same init_db() call (see caller).
    """
    user = db.execute("SELECT id FROM users ORDER BY created_at LIMIT 1").fetchone()
    if not user:
        return
    if db.execute("SELECT 1 FROM sessions WHERE owner_user_id IS NULL LIMIT 1").fetchone() is None \
            and db.execute("SELECT 1 FROM templates WHERE owner_user_id IS NULL AND (model_key IS NULL OR model_key!=?) LIMIT 1", (MODEL_KEY_EPC_SENEVAL,)).fetchone() is None:
        return
    if DATABASE.exists():
        backup = DATABASE.with_name(f"{DATABASE.stem}.bak-{datetime.now().strftime('%Y%m%d-%H%M%S')}{DATABASE.suffix}")
        try:
            shutil.copy2(DATABASE, backup)
        except OSError:
            pass
    db.execute("UPDATE sessions SET owner_user_id=? WHERE owner_user_id IS NULL", (user["id"],))
    # The reference EPC/SENEVAL template stays a common model (owner NULL); only
    # personal/custom questionnaires get attached to the first account.
    db.execute("UPDATE templates SET owner_user_id=? WHERE owner_user_id IS NULL AND (model_key IS NULL OR model_key!=?)", (user["id"], MODEL_KEY_EPC_SENEVAL))
    db.commit()


# Model registry, keyed by a stable identifier that survives a rename — the first step
# (lot 2a) of AUDIT_MODULARISATION_8800.md's Lot 2. Purely additive and currently inert:
# model_key/is_canonical are populated here but nothing else reads them yet. Every
# EPC_DOMAINS-derived version (any row named "EPC / SENEVAL", any `version`) is tagged
# model_key='epc_seneval'; only the single latest version is_canonical=1, matching
# exactly what ensure_reference_questionnaire_version()/migrate_v2_ownership() already
# treat as "the" reference row by name. Custom/imported templates are left untouched
# (model_key NULL, is_canonical 0) — the correct default for a model that isn't builtin.
MODEL_KEY_EPC_SENEVAL = "epc_seneval"


def ensure_model_identity(db: sqlite3.Connection) -> None:
    """Runs on every startup (idempotent: a no-op once model_key/is_canonical already
    match the current rows). Must run after ensure_reference_questionnaire_version()
    so "the latest version" reflects any version it just inserted."""
    db.execute("UPDATE templates SET model_key=? WHERE name='EPC / SENEVAL' AND (model_key IS NULL OR model_key!=?)", (MODEL_KEY_EPC_SENEVAL, MODEL_KEY_EPC_SENEVAL))
    latest = db.execute("SELECT id FROM templates WHERE model_key=? ORDER BY version DESC LIMIT 1", (MODEL_KEY_EPC_SENEVAL,)).fetchone()
    if latest:
        db.execute("UPDATE templates SET is_canonical=1 WHERE id=? AND is_canonical!=1", (latest["id"],))
        db.execute("UPDATE templates SET is_canonical=0 WHERE model_key=? AND id!=? AND is_canonical!=0", (MODEL_KEY_EPC_SENEVAL, latest["id"]))
    db.commit()


def seed_epc(db: sqlite3.Connection) -> str:
    tid, stamp = str(uuid.uuid4()), now()
    scale = {"type": "numeric", "min": 1, "max": 5, "labels": {"1": "Totalement en désaccord", "2": "En désaccord", "3": "Neutre", "4": "D’accord", "5": "Totalement d’accord"}}
    scoring = {"capacity": "mean_divided_by_scale_max", "outputRange": [0, 100]}
    consensus = {"method": "standard_deviation", "normalization": "theoretical_range", "factor": 2}
    db.execute("INSERT INTO templates (id,name,version,description,status,scale_json,scoring_json,consensus_json,grading_json,priority_json,created_at,updated_at,owner_user_id) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", (tid, "EPC / SENEVAL", 1, "Configuration initiale issue du questionnaire actuel.", "active", json.dumps(scale), json.dumps(scoring), json.dumps(consensus), json.dumps(GRADING), json.dumps({"maxPerDomain": 3}), stamp, stamp, None))
    for d_order, (code, label, indicators) in enumerate(EPC_DOMAINS, 1):
        did = str(uuid.uuid4())
        db.execute("INSERT INTO domains VALUES (?,?,?,?,?,?,?)", (did, tid, code, label, "", d_order, 1))
        for i_order, (reference, enonce) in enumerate(indicators, 1):
            db.execute("INSERT INTO indicators VALUES (?,?,?,?,?,?,?,?,?,?)", (str(uuid.uuid4()), did, reference, enonce, "", "numeric", 1, i_order, 1, "{}"))
    db.commit()
    return tid


def rows(db: sqlite3.Connection, sql: str, args=()):
    return [dict(r) for r in db.execute(sql, args).fetchall()]


def template_payload(db, template_id: str):
    template = db.execute("SELECT * FROM templates WHERE id=?", (template_id,)).fetchone()
    if not template:
        return None
    out = dict(template)
    for key in ["scale_json", "scoring_json", "consensus_json", "grading_json", "priority_json"]:
        out[key[:-5]] = json.loads(out.pop(key))
    out["domains"] = rows(db, "SELECT * FROM domains WHERE template_id=? ORDER BY display_order", (template_id,))
    for domain in out["domains"]:
        domain["active"] = bool(domain["active"])
        domain["indicators"] = rows(db, "SELECT * FROM indicators WHERE domain_id=? ORDER BY display_order", (domain["id"],))
        for indicator in domain["indicators"]:
            indicator["required"] = bool(indicator["required"]); indicator["active"] = bool(indicator["active"]); indicator["configuration"] = json.loads(indicator.pop("configuration_json"))
    return out
