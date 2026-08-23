"""Manifeste de restitution (lot 6, cf. AUDIT_MODULARISATION_8800.md) : quelles
sections de rapport, quels graphiques et quels prompts IA sont disponibles pour
un modele de questionnaire donne.

Un seul modele existe aujourd'hui (epc_seneval) - son manifeste liste
exactement ce qui existe deja, sans rien changer au rendu ("conserver
integralement le rapport EPC actuel", note de rollback de l'audit : "renderer
EPC historique selectionne pour model_key=epc_seneval"). Un futur second
modele apporterait son propre manifeste plutot que de brancher des
conditions supplementaires dans le code de rendu - c'est le seul point que
ce lot ajoute : un point d'indirection, pas une reecriture des generateurs
XLSX/DOCX (docx_bars_chart, docx_radar_chart, etc. restent inchanges, ils ne
savent toujours dessiner que ce que l'EPC dessine deja).
"""
from __future__ import annotations

from .db import MODEL_KEY_EPC_SENEVAL, template_payload

# Ordre et cles des feuilles du classeur XLSX (report_xlsx) - identique a
# l'ordre historique des onglets, jamais reordonne ni omis pour epc_seneval.
REPORT_SECTIONS_EPC_SENEVAL = [
    "synthese", "domaines", "indicateurs", "priorites", "analyses", "causes",
    "consequences", "leviers", "recommandations", "formations", "plan_action",
    "questionnaire",
]

AI_SYSTEM_PROMPT_EPC_SENEVAL = (
    "Tu assistes un modérateur d'atelier de diagnostic organisationnel EPC/SENEVAL. "
    "Tu interprètes des données déjà calculées ; tu ne recalcules jamais un score, tu n'inventes jamais un fait, "
    "une cause, une conséquence ou une recommandation absente des données fournies. "
    "Style professionnel, clair, factuel, sans jargon d'IA, sans formules comme « l'IA constate que ». Réponds en français."
)

# section_key -> (libelle affiche, instruction donnee a l'IA) ; l'ordre est
# l'ordre d'affichage dans le rapport DOCX/XLSX et de generation pour
# "generer tout le rapport".
AI_REPORT_SECTIONS_EPC_SENEVAL = {
    "resume_executif": ("Résumé exécutif", "Rédige un RÉSUMÉ EXÉCUTIF de l'atelier : situation générale, principaux constats, points forts, points de vigilance, priorités retenues, principales orientations."),
    "lecture_diagnostic": ("Lecture du diagnostic", "Rédige une LECTURE DU DIAGNOSTIC : tendances, écarts, convergences, divergences, domaines remarquables, à partir de la capacité et du consensus."),
    "synthese_domaines": ("Synthèse par domaine", "Rédige une SYNTHÈSE PAR DOMAINE : pour chaque domaine, résultat et interprétation prudente, en lien avec les analyses validées si disponibles."),
    "synthese_priorites": ("Synthèse des priorités", "Rédige une SYNTHÈSE DES PRIORITÉS : priorités retenues, constats, causes validées, leviers retenus."),
    "synthese_recommandations": ("Synthèse des recommandations", "Rédige une SYNTHÈSE DES RECOMMANDATIONS retenues, regroupées en catégories cohérentes si pertinent."),
    "synthese_formations": ("Synthèse des besoins de formation", "Rédige une SYNTHÈSE DES BESOINS DE FORMATION retenus."),
    "synthese_plan": ("Synthèse du plan d'action", "Rédige une SYNTHÈSE DU PLAN D'ACTION à partir des recommandations retenues."),
    "conclusion": ("Conclusion générale proposée", "Propose une CONCLUSION GÉNÉRALE concise et institutionnelle. Précise qu'il s'agit d'une proposition, pas d'une décision validée."),
}

RESTITUTION_MANIFESTS = {
    MODEL_KEY_EPC_SENEVAL: {
        "modelKey": MODEL_KEY_EPC_SENEVAL,
        "modelName": "EPC / SENEVAL",
        "reportSections": REPORT_SECTIONS_EPC_SENEVAL,
        "aiSystemPrompt": AI_SYSTEM_PROMPT_EPC_SENEVAL,
        "aiReportSections": {k: v[1] for k, v in AI_REPORT_SECTIONS_EPC_SENEVAL.items()},
        "aiSectionLabels": {k: v[0] for k, v in AI_REPORT_SECTIONS_EPC_SENEVAL.items()},
    },
}


def resolve_model_key(template: dict) -> str:
    """A template without its own model_key (a custom/blank questionnaire,
    cf. lot 2) still runs on the single EPC scoring engine that exists today
    (lot 3 deliberately deferred a multi-strategy registry) - so it gets the
    same restitution manifest as the canonical EPC/SENEVAL model."""
    return template.get("model_key") or MODEL_KEY_EPC_SENEVAL


def restitution_manifest(template: dict) -> dict:
    return RESTITUTION_MANIFESTS.get(resolve_model_key(template), RESTITUTION_MANIFESTS[MODEL_KEY_EPC_SENEVAL])


def session_restitution_manifest(db, session_id: str) -> dict:
    """Resolves the manifest for the model session_id's questionnaire belongs
    to - the one lookup point every route needing model-driven prompts,
    report sections or charts should call, instead of hardcoding a model."""
    session = db.execute("SELECT template_id FROM sessions WHERE id=?", (session_id,)).fetchone()
    if not session:
        return RESTITUTION_MANIFESTS[MODEL_KEY_EPC_SENEVAL]
    return restitution_manifest(template_payload(db, session["template_id"]))
