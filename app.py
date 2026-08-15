"""Local workshop diagnosis engine.

Run with: python app.py
The application uses only the Python standard library and SQLite.  It is meant
to be a dependable local-first starting point, not a simulated real-time app.
"""
from __future__ import annotations

import csv
import base64
import json
import math
import re
import sqlite3
import sys
import unicodedata
import uuid
import zipfile
from datetime import datetime, timezone
from http import HTTPStatus
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from io import StringIO
from io import BytesIO
from pathlib import Path
from urllib.parse import parse_qs, urlparse
from xml.etree import ElementTree as ET

try:  # Used only to generate the downloadable Excel template; the app stays local.
    import xlsxwriter
except ImportError:
    xlsxwriter = None
try:
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
except ImportError:
    Document = None
    Inches = None
try:  # Used to draw the report chart images (Word and Excel both embed the same PNGs); reports degrade to text/tables without it.
    from PIL import Image as PILImage, ImageDraw, ImageFont
except ImportError:
    PILImage = None

ROOT = Path(__file__).resolve().parent
STATIC = ROOT / "static"
DATABASE = ROOT / "data" / "workshops.sqlite3"
IMPORTS = {}
MATRIX_COLUMNS = ["Domaine", "Ordre domaine", "Code indicateur", "Indicateur", "Description", "Ordre indicateur", "Type réponse", "Obligatoire", "Actif"]
PARAMETERS = ["Nom questionnaire", "Description", "Version", "Type d'échelle", "Valeur minimum", "Valeur maximum", "Libellés des valeurs", "Nombre de priorités par domaine"]


def now() -> str:
    return datetime.now(timezone.utc).isoformat()


def slugify(text: str) -> str:
    text = unicodedata.normalize("NFKD", text or "").encode("ascii", "ignore").decode("ascii")
    text = re.sub(r"[^A-Za-z0-9]+", "-", text).strip("-")
    return text or "sans-titre"


def export_filename(*parts, ext: str) -> str:
    slug = "_".join(slugify(p) for p in parts if p)
    return f"{slug}_{datetime.now().strftime('%Y-%m-%d')}.{ext}"


EPC_DOMAINS = [
    ("grh", "Gestion des ressources humaines", [
        "Formation au personnel", "Priorités de notre organisation", "Compétences pour notre mission", "Nombre du personnel", "Diversité de nos bénéficiaires", "Recrutement des membres", "Évaluation du personnel", "Résolution des conflits", "Allocation des tâches", "Pratiques de supervision",
    ]),
    ("grf", "Gestion des ressources financières", [
        "Équilibre des recettes et des dépenses", "Allocation des fonds", "Prévisions financières", "Modification des dépenses", "Évitement des perturbations", "Décaissements périodiques", "Appui financier des bailleurs", "Moins de dépendance", "Ressources pour les activités", "Ressources pour l’équipement",
    ]),
    ("parteq", "Participation équitable", [
        "Évaluation des besoins", "Conception des projets", "Mise en œuvre des projets", "Suivi et évaluation des projets", "Accès aux activités", "Bénéfice équitable", "Promotion de l’équité", "Évaluation des changements", "Besoins changeants des participants", "Dialogue pour le développement équitable",
    ]),
    ("dur", "Durabilité des acquis", [
        "Durabilité environnementale à la conception", "Durabilité économique à la conception", "Durabilité institutionnelle à la conception", "Durabilité environnementale à la mise en œuvre", "Durabilité économique à la mise en œuvre", "Durabilité institutionnelle à la mise en œuvre", "Durabilité environnementale au suivi-évaluation", "Durabilité économique au suivi-évaluation", "Durabilité institutionnelle au suivi-évaluation", "Appui technique et durabilité",
    ]),
    ("partn", "Partenariat", [
        "Liens avec les décideurs politiques", "Liens avec le secteur privé", "Partenariats avec d’autres organisations", "Suivi de nos partenariats", "Avantages financiers", "Compétences techniques", "Nouveaux réseaux et relations", "Confiance et coopération", "Contribution aux objectifs partagés", "Effort de coopération",
    ]),
    ("apporg", "Apprentissage organisationnel", [
        "Évaluation des projets", "Implication des structures dans les défis", "Interdépendance des structures", "Informations pour le travail", "Informations pour les priorités", "Travail d’équipe pour les défis", "Travail d’équipe des responsables", "Réunions et apprentissage organisationnel", "Expression libre lors des réunions", "Prise de risque pour les innovateurs",
    ]),
    ("gouv", "Gestion stratégique et gouvernance", [
        "Rapportage pour les bailleurs", "Mobilisation des fonds", "Relations publiques", "Plaidoyer", "Définition de politique", "Représentation des bénéficiaires", "Engagement et décisions prises", "Planification stratégique et environnement externe", "Initiatives et plans stratégiques", "Suivi du progrès",
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
    CREATE TABLE IF NOT EXISTS sessions (id TEXT PRIMARY KEY, template_id TEXT NOT NULL REFERENCES templates(id), template_version INTEGER NOT NULL, name TEXT NOT NULL, organization TEXT, location TEXT, date TEXT, status TEXT NOT NULL, created_at TEXT NOT NULL, closed_at TEXT);
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
    """)
    db.commit()
    existing_columns = {r["name"] for r in db.execute("PRAGMA table_info(participants)")}
    if "display_name" not in existing_columns:
        db.execute("ALTER TABLE participants ADD COLUMN display_name TEXT")
        db.commit()
    if db.execute("SELECT 1 FROM templates LIMIT 1").fetchone() is None:
        seed_epc(db)


def seed_epc(db: sqlite3.Connection) -> str:
    tid, stamp = str(uuid.uuid4()), now()
    scale = {"type": "numeric", "min": 1, "max": 5, "labels": {"1": "Totalement en désaccord", "2": "En désaccord", "3": "Neutre", "4": "D’accord", "5": "Totalement d’accord"}}
    scoring = {"capacity": "mean_divided_by_scale_max", "outputRange": [0, 100]}
    consensus = {"method": "standard_deviation", "normalization": "theoretical_range", "factor": 2}
    db.execute("INSERT INTO templates VALUES (?,?,?,?,?,?,?,?,?,?,?,?)", (tid, "EPC / SENEVAL", 1, "Configuration initiale issue du questionnaire actuel.", "active", json.dumps(scale), json.dumps(scoring), json.dumps(consensus), json.dumps(GRADING), json.dumps({"maxPerDomain": 3}), stamp, stamp))
    for d_order, (code, label, indicators) in enumerate(EPC_DOMAINS, 1):
        did = str(uuid.uuid4())
        db.execute("INSERT INTO domains VALUES (?,?,?,?,?,?,?)", (did, tid, code, label, "", d_order, 1))
        for i_order, indicator in enumerate(indicators, 1):
            db.execute("INSERT INTO indicators VALUES (?,?,?,?,?,?,?,?,?,?)", (str(uuid.uuid4()), did, f"{code}-{i_order:02d}", indicator, "", "numeric", 1, i_order, 1, "{}"))
    db.commit()
    return tid


def rows(db: sqlite3.Connection, sql: str, args=()):
    return [dict(r) for r in db.execute(sql, args).fetchall()]


def session_label(db: sqlite3.Connection, sid: str) -> str:
    row = db.execute("SELECT name FROM sessions WHERE id=?", (sid,)).fetchone()
    return row["name"] if row else "atelier"


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


def next_order(db, table, where_col, where_value):
    return (db.execute(f"SELECT COALESCE(MAX(display_order),0)+1 FROM {table} WHERE {where_col}=?", (where_value,)).fetchone()[0])


def clone_template(db, template_id, name=None):
    old = template_payload(db, template_id)
    if not old: raise ValueError("Configuration introuvable")
    tid, stamp = str(uuid.uuid4()), now()
    version = db.execute("SELECT COALESCE(MAX(version),0)+1 FROM templates WHERE name=?", (name or old["name"],)).fetchone()[0]
    db.execute("INSERT INTO templates VALUES (?,?,?,?,?,?,?,?,?,?,?,?)", (tid, name or old["name"], version, old["description"], "active", json.dumps(old["scale"]), json.dumps(old["scoring"]), json.dumps(old["consensus"]), json.dumps(old["grading"]), json.dumps(old["priority"]), stamp, stamp))
    for d in old["domains"]:
        did=str(uuid.uuid4()); db.execute("INSERT INTO domains VALUES (?,?,?,?,?,?,?)",(did,tid,d["code"],d["label"],d["description"],d["display_order"],d["active"]))
        for i in d["indicators"]:
            db.execute("INSERT INTO indicators VALUES (?,?,?,?,?,?,?,?,?,?)",(str(uuid.uuid4()),did,i["code"],i["label"],i["description"],i["response_type"],i["required"],i["display_order"],i["active"],json.dumps(i["configuration"])))
    db.commit(); return tid


def create_blank_template(db, data):
    tid, stamp = str(uuid.uuid4()), now(); name=data.get("name", "Nouveau questionnaire").strip()
    if not name: raise ValueError("Le nom est obligatoire")
    version=db.execute("SELECT COALESCE(MAX(version),0)+1 FROM templates WHERE name=?",(name,)).fetchone()[0]
    scale={"type":"numeric","min":1,"max":5,"labels":{}}
    db.execute("INSERT INTO templates VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",(tid,name,version,data.get("description", ""),"active",json.dumps(scale),json.dumps({"capacity":"mean_divided_by_scale_max","outputRange":[0,100]}),json.dumps({"method":"standard_deviation","normalization":"theoretical_range","factor":2}),json.dumps(GRADING),json.dumps({"maxPerDomain":3}),stamp,stamp)); db.commit(); return tid


def matrix_xlsx(template):
    if not xlsxwriter: raise RuntimeError("Le générateur XLSX local n'est pas disponible")
    out=BytesIO(); wb=xlsxwriter.Workbook(out, {"in_memory": True}); head=wb.add_format({"bold":True,"bg_color":"#1F4E78","font_color":"#FFFFFF"}); wrap=wb.add_format({"text_wrap":True,"valign":"top"})
    guide=wb.add_worksheet("MODE D’EMPLOI"); guide.set_column(0,0,110); guide.write("A1","Cette matrice permet de préparer un questionnaire avant de l’importer dans l’outil.",head); guide.write_column("A3",["1. Dans la feuille PARAMETRES, remplacez la valeur d’exemple par le vrai nom de votre questionnaire.","2. Complétez la description (facultatif) et les libellés de l’échelle de notation (une seule fois, valables pour tout le questionnaire).","3. Dans la feuille QUESTIONNAIRE, saisissez une ligne par indicateur.","4. Répétez le nom du domaine pour les indicateurs appartenant au même domaine.","5. La numérotation sera générée automatiquement par l’outil.","6. Les lignes d’exemple (matrice PARAMETRES et QUESTIONNAIRE) peuvent être remplacées ou supprimées : elles servent uniquement de modèle, à l’image du questionnaire EPC/SENEVAL."],wrap)
    default_labels={"5":"Totalement d’accord","4":"D’accord","3":"Neutre","2":"Pas d’accord","1":"Totalement en désaccord"}
    ps=wb.add_worksheet("PARAMETRES"); ps.write_row(0,0,["Nom du questionnaire (à remplacer par le vôtre)",template["name"]],head); ps.write_row(1,0,["Description",template["description"]],wrap); ps.write_row(3,0,["Note","Libellé (exemple EPC/SENEVAL, à adapter)"],head); labels=template["scale"].get("labels",{}); [ps.write_row(4+(5-n),0,[n,labels.get(str(n)) or default_labels[str(n)]]) for n in range(5,0,-1)]; ps.set_column(0,0,38); ps.set_column(1,1,55)
    ws=wb.add_worksheet("QUESTIONNAIRE"); ws.write_row(0,0,["Domaine","Référence","Indicateur qualitatif ou Capacité"],head); ws.freeze_panes(1,0); row=1
    if not template["domains"]: ws.write_row(row,0,["EXEMPLE — Gestion des Ressources Humaines","EXEMPLE — Formation au personnel","EXEMPLE — Nous offrons régulièrement la formation au personnel"],wrap); row+=1
    for d in template["domains"]:
        for i in d["indicators"]:
            ws.write_row(row,0,[d["label"],i["label"],i["description"]],wrap); row+=1
    ws.set_column(0,0,34); ws.set_column(1,1,38); ws.set_column(2,2,75); wb.close(); return out.getvalue()


def blank_matrix_xlsx():
    return matrix_xlsx({"name":"Exemple à remplacer : Diagnostic EPC / [nom de l’atelier]", "description":"", "version":1, "scale":{"type":"numeric","min":1,"max":5,"labels":{}}, "priority":{"maxPerDomain":3}, "domains":[]})

def report_rows(db, sid):
    a=analysis(db,sid); return a, [[d['label'],d['capacity'],d['consensus'],d['gradedCapacity'],d['gradedConsensus'],d['responses']] for d in a['domains']]

def report_xlsx(db,sid):
    a,rs=report_rows(db,sid); q=qualitative_data(db,sid); meta=report_data(db,sid)["meta"]; template=template_payload(db,a['session']['template_id']); out=BytesIO(); wb=xlsxwriter.Workbook(out,{"in_memory":True}); h=wb.add_format({"bold":True,"bg_color":"#1F4E78","font_color":"#FFFFFF"})
    analyses={x['priority_id']:x for x in q['analyses']}; priority_rows=[[p['id'],p['domain_label'],p['indicator_code'],p['indicator_label'],analyses.get(p['id'],{}).get('problem','')] for p in q['priorities']]
    sheets=[("Synthèse",["Atelier","Organisation","Lieu","Date","Animateur","Public","Contexte","Conclusion","Capacité","Consensus"],[[a['session']['name'],a['session']['organization'],a['session']['location'],a['session']['date'],meta['facilitator'],meta['audience'],meta['context'],meta['conclusion'],a['global']['capacity'],a['global']['consensus']]]),("Domaines",["Domaine","Capacité","Consensus","Cap. graduée","Cons. gradué","Réponses"],rs),("Indicateurs",["Domaine","Référence","Capacité","Consensus","Réponses","Manquants"],[[d['label'],i['label'],i['capacity'],i['consensus'],i['responses'],i['missing']] for d in a['domains'] for i in d['indicators']]),("Priorités",["ID priorité","Domaine","Référence","Indicateur","Constat"],priority_rows),("Analyses",["ID","Priorité","Constat"],[[x['id'],x['priority_id'],x['problem']] for x in q['analyses']]),("Causes",["ID","Priorité","Parent","Cause","Type","Statut"],[[x['id'],x['priority_id'],x['parent_id'],x['content'],x['item_type'],x['validation_status']] for x in q['entries'] if x['kind']=='cause']),("Conséquences",["ID","Priorité","Conséquence","Statut"],[[x['id'],x['priority_id'],x['content'],x['validation_status']] for x in q['entries'] if x['kind']=='consequence']),("Leviers",["ID","Priorité","Levier","Commentaire","Statut"],[[x['id'],x['priority_id'],x['content'],x['comment'],x['validation_status']] for x in q['entries'] if x['kind']=='lever']),("Recommandations",["ID","Priorité","Cause","Levier","Titre","Description","Catégorie","Niveau","Responsable","Échéance","Statut"],[[x['id'],x['priority_id'],x['cause_id'],x['lever_id'],x['title'],x['description'],x['category'],x['priority_level'],x['owner'],x['horizon'],x['status']] for x in q['recommendations']]),("Formations",["ID","Priorité","Recommandation","Intitulé","Besoin","Public","Niveau","Commentaire"],[[x['id'],x['priority_id'],x['recommendation_id'],x['title'],x['need_text'],x['target_audience'],x['priority_level'],x['comment']] for x in q['trainingTopics']]),("Plan_action",["N°","Action / recommandation","Origine","Responsable","Échéance","Priorité","Statut"],[[n+1,x['title'],x['priority_id'] or '—',x['owner'] or '—',x['horizon'] or '—',x['priority_level'],x['status']] for n,x in enumerate(q['recommendations']) if x['status']=='Retenue']),("Questionnaire",["Domaine","Référence","Indicateur","Échelle"],[[d['label'],i['label'],i['description'],f"{template['scale']['min']}–{template['scale']['max']}"] for d in template['domains'] for i in d['indicators'] if i['active']])]
    for name,head,data in sheets:
        s=wb.add_worksheet(name);s.write_row(0,0,head,h);[s.write_row(n+1,0,row) for n,row in enumerate(data)];s.set_column(0,len(head)-1,24)
    domains=[d for d in a['domains'] if d.get('capacity') is not None]
    if PILImage is not None and domains:
        gs=wb.add_worksheet("Graphiques"); gs.set_column(0,0,4); row=1
        gs.write(0,0,"Vue synthétique",h)
        row=xlsx_add_image(gs,row,0,pil_chart_grid([
            docx_radar_chart(domains),
            docx_bars_chart(domains,"standard","Notes standardisées par domaine",with_mean=True),
            docx_bars_chart(domains,"graded","Notes graduées par domaine",with_mean=True),
            docx_grid_chart(domains),
        ]))
        if len(domains)>1:
            row+=1; gs.write(row,0,"Analyse comparative — cohorte des domaines",h); row+=1
            row=xlsx_add_image(gs,row,0,docx_cohort_chart(domains))
        for i in range(0,len(domains),4):
            group=domains[i:i+4]
            tiles=[]
            for d in group:
                inds=[ind for ind in d["indicators"] if ind.get("capacity") is not None]
                if inds:
                    caption=f"Capacité {pdf_fmt(d['capacity'])} · Consensus {pdf_fmt(d['consensus'])} · {d['responses']} répondant(s)"
                    tiles.append((docx_bars_chart(inds,"standard",d["label"]),caption))
            if tiles:
                row+=1; gs.write(row,0,"Détail par domaine",h); row+=1
                row=xlsx_add_image(gs,row,0,pil_chart_grid(tiles))
        gs.activate()
    wb.close();return out.getvalue()

DOCX_FONT_CACHE = {}


def docx_font(size, bold=False):
    key = (size, bold)
    if key in DOCX_FONT_CACHE:
        return DOCX_FONT_CACHE[key]
    names = ["arialbd.ttf", "Arial Bold.ttf", "DejaVuSans-Bold.ttf", "LiberationSans-Bold.ttf"] if bold else \
            ["arial.ttf", "Arial.ttf", "DejaVuSans.ttf", "LiberationSans-Regular.ttf"]
    dirs = ["C:/Windows/Fonts", "/usr/share/fonts/truetype/dejavu", "/usr/share/fonts/truetype/liberation",
            "/System/Library/Fonts/Supplemental", "/System/Library/Fonts"]
    font = None
    for d in dirs:
        for n in names:
            try:
                font = ImageFont.truetype(f"{d}/{n}", size); break
            except Exception:
                continue
        if font:
            break
    if font is None:
        try:
            font = ImageFont.load_default(size=size)
        except TypeError:
            font = ImageFont.load_default()
    DOCX_FONT_CACHE[key] = font
    return font


def docx_text_w(draw, text, font):
    b = draw.textbbox((0, 0), text, font=font)
    return b[2] - b[0]


def docx_rotated_text(img, text, x, y, angle, font, fill, anchor_end=False):
    draw0 = ImageDraw.Draw(img)
    b = draw0.textbbox((0, 0), text, font=font)
    w, h = b[2] - b[0], b[3] - b[1]
    pad = 4
    txt_img = PILImage.new("RGBA", (w + pad * 2, h + pad * 2), (255, 255, 255, 0))
    ImageDraw.Draw(txt_img).text((pad, pad), text, font=font, fill=fill)
    rot = txt_img.rotate(angle, expand=1, resample=PILImage.BICUBIC)
    px = x - (rot.width if anchor_end else 0)
    img.paste(rot, (int(px), int(y)), rot)


def docx_dashed_line(draw, p1, p2, color, dash=8, gap=5, width=2):
    x1, y1 = p1; x2, y2 = p2
    length = math.hypot(x2 - x1, y2 - y1)
    if length == 0:
        return
    ux, uy = (x2 - x1) / length, (y2 - y1) / length
    d = 0
    while d < length:
        e = min(d + dash, length)
        draw.line([x1 + ux * d, y1 + uy * d, x1 + ux * e, y1 + uy * e], fill=color, width=width)
        d += dash + gap


def docx_bars_chart(items, mode, title, with_mean=False):
    """Vertical grouped column chart, raster twin of static/app.js bars() / pdf_bars_chart()."""
    grad = mode == "graded"
    data = [d for d in items if d.get("capacity") is not None]
    if with_mean and data:
        mc = sum(d["capacity"] for d in data) / len(data)
        ms = sum(d["consensus"] for d in data) / len(data)
        gc = sum(d.get("gradedCapacity") or 0 for d in data) / len(data)
        gs = sum(d.get("gradedConsensus") or 0 for d in data) / len(data)
        data = data + [{"label": "Moyen", "code": "Moyen", "capacity": mc, "consensus": ms, "gradedCapacity": gc, "gradedConsensus": gs}]
    w = max(900, 90 + len(data) * 140) if data else 900
    h = 560
    img = PILImage.new("RGB", (w, h), "white")
    draw = ImageDraw.Draw(img)
    f_title, f_axis, f_legend, f_val, f_cat, f_cat_b, f_xtitle = docx_font(24, True), docx_font(15), docx_font(15), docx_font(14), docx_font(14), docx_font(14, True), docx_font(17, True)
    draw.text((w / 2, 30), title, font=f_title, fill="black", anchor="mm")
    if not data:
        draw.text((30, 66), "Pas encore de données disponibles.", font=f_axis, fill="black")
        return img
    left, right = 100, w - 30
    top, bottom = 72, h - 120
    plot_w, plot_h = right - left, bottom - top
    axis_min, axis_max = 20, 100
    scaleY = lambda v: bottom - (v - axis_min) / (axis_max - axis_min) * plot_h
    draw.rectangle([left, top, right, bottom], fill="#e9e9e9", outline="black")
    for v in (20, 40, 60, 80, 100):
        yy = scaleY(v)
        draw.line([left, yy, right, yy], fill="#bbbbbb")
        draw.text((left - 12, yy), str(v), font=f_axis, fill="black", anchor="rm")
    cap_label = "Capacité graduée" if grad else "Capacité standardisée"
    cons_label = "Consensus gradué" if grad else "Consensus standardisé"
    draw.rectangle([left, 40, left + 16, 56], fill="#9999FF")
    draw.text((left + 24, 48), cap_label, font=f_legend, fill="black", anchor="lm")
    draw.rectangle([left + 240, 40, left + 256, 56], fill="#993366")
    draw.text((left + 264, 48), cons_label, font=f_legend, fill="black", anchor="lm")
    n = len(data); step = plot_w / n
    for i, d in enumerate(data):
        x = left + i * step
        cap = d.get("gradedCapacity") if grad else d.get("capacity")
        cons = d.get("gradedConsensus") if grad else d.get("consensus")
        cv, sv = max(axis_min, cap or 0), max(axis_min, cons or 0)
        bw = step * 0.32
        y_cap, y_cons = scaleY(cv), scaleY(sv)
        draw.rectangle([x + step * .12, y_cap, x + step * .12 + bw, bottom], fill="#9999FF", outline="black")
        draw.rectangle([x + step * .12 + bw + 4, y_cons, x + step * .12 + bw + 4 + bw, bottom], fill="#993366", outline="black")
        cap_txt = "—" if cap is None else (str(round(cap)) if grad else f"{cap:.1f}")
        cons_txt = "—" if cons is None else (str(round(cons)) if grad else f"{cons:.1f}")
        draw.text((x + step * .12 + bw / 2, y_cap - 12), cap_txt, font=f_val, fill="black", anchor="mm")
        draw.text((x + step * .12 + bw + 4 + bw / 2, y_cons - 12), cons_txt, font=f_val, fill="black", anchor="mm")
        docx_rotated_text(img, pdf_short_label(d), x + step * .12 + bw + 2, bottom + 8, -40, f_cat_b if d.get("label") == "Moyen" else f_cat, "black")
    draw.text((left + plot_w / 2, h - 34), "Domaines de compétence", font=f_xtitle, fill="black", anchor="mm")
    return img


def docx_grid_chart(items):
    """Quadrant scatter with numbered markers, raster twin of graduatedGrid() / pdf_grid_chart()."""
    data = [d for d in items if d.get("capacity") is not None]
    w, h = 700, 620
    img = PILImage.new("RGB", (w, h), "white")
    draw = ImageDraw.Draw(img)
    f_title, f_axis, f_axis_title, f_quad, f_num, f_legend = docx_font(22, True), docx_font(14), docx_font(16, True), docx_font(13), docx_font(15, True), docx_font(13)
    draw.text((w / 2, 26), "Positionnement des domaines", font=f_title, fill="black", anchor="mm")
    if not data:
        draw.text((30, 60), "Pas encore de données disponibles pour la grille graduée.", font=f_axis, fill="black")
        return img
    left, right = 100, w - 40
    top, bottom = 60, h - 150
    plot_w, plot_h = right - left, bottom - top
    sx = lambda v: left + v / 100 * plot_w
    sy = lambda v: bottom - v / 100 * plot_h
    draw.rectangle([left, top, right, bottom], fill="#e9e9e9", outline="black")
    for v in (0, 20, 40, 60, 80, 100):
        draw.line([sx(v), top, sx(v), bottom], fill="#bbbbbb")
        draw.line([left, sy(v), right, sy(v)], fill="#bbbbbb")
        draw.text((sx(v), bottom + 16), str(v), font=f_axis, fill="black", anchor="mm")
        draw.text((left - 12, sy(v)), str(v), font=f_axis, fill="black", anchor="rm")
    draw.line([sx(50), top, sx(50), bottom], fill="#333333", width=2)
    draw.line([left, sy(50), right, sy(50)], fill="#333333", width=2)
    draw.text((left + plot_w / 2, bottom + 38), "Capacité", font=f_axis_title, fill="black", anchor="mm")
    docx_rotated_text(img, "Consensus", 22, top + plot_h / 2 + 45, 90, f_axis_title, "black")
    draw.text((left + 6, top + 8), "Faible capacité / consensus élevé", font=f_quad, fill="black")
    draw.text((right - 6 - docx_text_w(draw, "Capacité élevée / consensus élevé", f_quad), top + 8), "Capacité élevée / consensus élevé", font=f_quad, fill="black")
    draw.text((left + 6, bottom - 20), "Faible capacité / consensus faible", font=f_quad, fill="black")
    draw.text((right - 6 - docx_text_w(draw, "Capacité élevée / consensus faible", f_quad), bottom - 20), "Capacité élevée / consensus faible", font=f_quad, fill="black")
    avg_cap = sum(d["capacity"] for d in data) / len(data)
    avg_cons = sum(d["consensus"] for d in data) / len(data)
    order = sorted(range(len(data)), key=lambda i: data[i]["capacity"] + data[i]["consensus"])
    rank = {idx: r for r, idx in enumerate(order)}
    dirs = [(16, -14), (16, 18), (-16, -14), (-16, 18)]
    for i, d in enumerate(data):
        cx, cy = sx(d["capacity"]), sy(d["consensus"])
        r = 7
        draw.polygon([(cx, cy - r), (cx + r, cy), (cx, cy + r), (cx - r, cy)], fill="#000080")
        dx, dy = dirs[rank[i] % 4]
        if cx + dx > right - 16: dx = -abs(dx)
        elif cx + dx < left + 16: dx = abs(dx)
        if cy + dy < top + 12: dy = abs(dy)
        elif cy + dy > bottom - 12: dy = -abs(dy)
        draw.text((cx + dx, cy + dy), str(i + 1), font=f_num, fill="#000080", anchor="lm" if dx > 0 else "rm")
    ax, ay = sx(avg_cap), sy(avg_cons)
    label = "Moyenne générale"
    label_w = docx_text_w(draw, label, f_num)
    best = None
    for ddx, ddy, anchor in ((20, -16, "l"), (20, 22, "l"), (-20, -16, "r"), (-20, 22, "r")):
        lx, ly = ax + ddx, ay + ddy
        box_left = lx if anchor == "l" else lx - label_w
        box_right = box_left + label_w
        box_top, box_bottom = ly - 10, ly + 10
        in_bounds = box_left >= left + 2 and box_right <= right - 2 and box_top >= top + 2 and box_bottom <= bottom - 2
        min_dist = min(
            math.hypot(sx(d["capacity"]) - max(box_left, min(sx(d["capacity"]), box_right)),
                       sy(d["consensus"]) - max(box_top, min(sy(d["consensus"]), box_bottom)))
            for d in data
        )
        score = (in_bounds, min_dist)
        if best is None or score > best[0]:
            best = (score, lx, ly, anchor)
    _, lx, ly, anchor = best
    r = 7
    draw.polygon([(ax, ay - r), (ax + r, ay), (ax, ay + r), (ax - r, ay)], fill="white", outline="black", width=2)
    draw.text((lx, ly), label, font=f_num, fill="black", anchor="lm" if anchor == "l" else "rm")
    legend = " · ".join(f"{i + 1} {d['label']}" for i, d in enumerate(data))
    words, lines, cur = legend.split(" "), [], ""
    for word in words:
        trial = (cur + " " + word).strip()
        if not cur or docx_text_w(draw, trial, f_legend) <= w - 60:
            cur = trial
        else:
            lines.append(cur); cur = word
    if cur:
        lines.append(cur)
    ly2 = bottom + 74
    for line in lines:
        draw.text((30, ly2), line, font=f_legend, fill="#444444"); ly2 += 20
    return img


def docx_cohort_chart(items):
    """Standardised scores (bars) and graded scores (lines) across domains + the
    cohort mean as a 7th column, raster twin of the redesigned cohortChart()."""
    data = [d for d in items if d.get("capacity") is not None]
    w = max(900, 90 + (len(data) + 1) * 140) if data else 900
    h = 560
    img = PILImage.new("RGB", (w, h), "white")
    draw = ImageDraw.Draw(img)
    f_title, f_axis, f_legend, f_val, f_cat, f_cat_b, f_xtitle = docx_font(20, True), docx_font(15), docx_font(14), docx_font(14), docx_font(14), docx_font(14, True), docx_font(17, True)
    draw.text((w / 2, 24), "Notes standardisées et graduées — cohorte des domaines", font=f_title, fill="black", anchor="mm")
    if not data:
        draw.text((30, 60), "Pas encore de données disponibles pour l’analyse de cohorte.", font=f_axis, fill="black")
        return img
    mc = sum(d["capacity"] for d in data) / len(data)
    ms = sum(d["consensus"] for d in data) / len(data)
    # Mirrors static/app.js cohort(): missing graded values count as 0, divided by
    # the full domain count (not just the domains that have a graded value).
    gc = sum(d.get("gradedCapacity") or 0 for d in data) / len(data)
    gs = sum(d.get("gradedConsensus") or 0 for d in data) / len(data)
    all_data = data + [{"label": "Moyen", "code": "Moyen", "capacity": mc, "consensus": ms, "gradedCapacity": gc, "gradedConsensus": gs}]
    left, right = 100, w - 30
    top, bottom = 108, h - 120
    plot_w, plot_h = right - left, bottom - top
    axis_min, axis_max = 20, 100
    scaleY = lambda v: bottom - (v - axis_min) / (axis_max - axis_min) * plot_h
    draw.rectangle([left, top, right, bottom], fill="#e9e9e9", outline="black")
    for v in (20, 40, 60, 80, 100):
        yy = scaleY(v)
        draw.line([left, yy, right, yy], fill="#bbbbbb")
        draw.text((left - 12, yy), str(v), font=f_axis, fill="black", anchor="rm")
    draw.rectangle([left, 56, left + 16, 72], fill="#9999FF")
    draw.text((left + 24, 64), "Capacité standardisée", font=f_legend, fill="black", anchor="lm")
    draw.rectangle([left + 240, 56, left + 256, 72], fill="#993366")
    draw.text((left + 264, 64), "Consensus standardisé", font=f_legend, fill="black", anchor="lm")
    draw.line([left, 86, left + 32, 86], fill="#800000", width=3)
    draw.text((left + 38, 86), "Capacité graduée", font=f_legend, fill="black", anchor="lm")
    draw.line([left + 220, 86, left + 252, 86], fill="#0000ff", width=3)
    draw.text((left + 258, 86), "Consensus gradué", font=f_legend, fill="black", anchor="lm")
    n = len(all_data); step = plot_w / n; bw = step * .32
    cap_pts, cons_pts = [], []
    for i, d in enumerate(all_data):
        x = left + i * step
        c_, s_ = d["capacity"], d["consensus"]
        cv, sv = max(axis_min, c_ or 0), max(axis_min, s_ or 0)
        y_cap, y_cons = scaleY(cv), scaleY(sv)
        draw.rectangle([x + step * .12, y_cap, x + step * .12 + bw, bottom], fill="#9999FF", outline="black")
        draw.rectangle([x + step * .12 + bw + 4, y_cons, x + step * .12 + bw + 4 + bw, bottom], fill="#993366", outline="black")
        lx = x + step * .12 + bw + 2
        docx_rotated_text(img, pdf_short_label(d), lx, bottom + 8, -40, f_cat_b if d["label"] == "Moyen" else f_cat, "black")
        if d.get("gradedCapacity") is not None:
            cap_pts.append((lx, scaleY(d["gradedCapacity"])))
        if d.get("gradedConsensus") is not None:
            cons_pts.append((lx, scaleY(d["gradedConsensus"])))
    if len(cap_pts) > 1:
        draw.line(cap_pts, fill="#800000", width=3)
    for px, py in cap_pts:
        draw.polygon([(px, py - 6), (px + 6, py + 5), (px - 6, py + 5)], fill="#800000")
    if len(cons_pts) > 1:
        draw.line(cons_pts, fill="#0000ff", width=3)
    for px, py in cons_pts:
        draw.line([px - 5, py - 5, px + 5, py + 5], fill="#0000ff", width=2)
        draw.line([px - 5, py + 5, px + 5, py - 5], fill="#0000ff", width=2)
    draw.text((left + plot_w / 2, h - 34), "Domaines", font=f_xtitle, fill="black", anchor="mm")
    return img


def docx_radar_chart(items):
    """Complementary synthesis radar, raster twin of radar() / pdf_radar_chart()."""
    data = [d for d in items if d.get("capacity") is not None]
    w, h = 560, 580
    img = PILImage.new("RGB", (w, h), "white")
    draw = ImageDraw.Draw(img)
    f_legend, f_label = docx_font(15), docx_font(13)
    if len(data) < 3:
        draw.text((30, 30), "Radar non disponible : au moins 3 domaines avec des données sont nécessaires.", font=f_legend, fill="black")
        return img
    n = len(data)
    cx, cy, r = w / 2, h / 2 + 16, 190
    draw.rectangle([20, 20, 34, 34], fill="#176b4b")
    draw.text((42, 27), "Capacité", font=f_legend, fill="black", anchor="lm")
    draw.rectangle([170, 20, 184, 34], fill="#536271")
    draw.text((192, 27), "Consensus", font=f_legend, fill="black", anchor="lm")
    pts_cap, pts_cons = [], []
    for i, d in enumerate(data):
        ang = -math.pi / 2 + i * 2 * math.pi / n
        ex, ey = cx + math.cos(ang) * r, cy + math.sin(ang) * r
        draw.line([cx, cy, ex, ey], fill="#aaaabb")
        lx, ly = cx + math.cos(ang) * (r + 20), cy + math.sin(ang) * (r + 20)
        draw.text((lx, ly), pdf_short_label(d)[:9], font=f_label, fill="black", anchor="mm")
        cv, sv = (d.get("capacity") or 0) / 100, (d.get("consensus") or 0) / 100
        pts_cap.append((cx + math.cos(ang) * r * cv, cy + math.sin(ang) * r * cv))
        pts_cons.append((cx + math.cos(ang) * r * sv, cy + math.sin(ang) * r * sv))
    draw.polygon(pts_cap, fill="#cfe3da", outline="#176b4b")
    draw.polygon(pts_cons, outline="#536271", width=2)
    return img


def docx_add_chart(doc, img):
    buf = BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    doc.add_picture(buf, width=Inches(6.3))


DOCX_NAVY = RGBColor(0x1B, 0x3A, 0x5C) if RGBColor else None
DOCX_MUTED = RGBColor(0x64, 0x74, 0x8B) if RGBColor else None


def docx_shade_cell(cell, color_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color_hex)
    tcPr.append(shd)


def docx_style_heading(heading, color=None):
    for run in heading.runs:
        run.font.color.rgb = color or DOCX_NAVY
    return heading


def docx_note(doc, text, size=9, color=None, italic=True):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.font.italic = italic
    run.font.color.rgb = color or DOCX_MUTED
    return p


def docx_style_table(table, header_bg="1B3A5C"):
    table.style = "Table Grid"
    for cell in table.rows[0].cells:
        docx_shade_cell(cell, header_bg)
        for p in cell.paragraphs:
            for run in p.runs:
                run.font.bold = True
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    for i, row in enumerate(table.rows[1:]):
        if i % 2 == 1:
            for cell in row.cells:
                docx_shade_cell(cell, "F2F5F8")
    return table


def docx_metrics_row(doc, metrics):
    """metrics: list of (label, value) tuples, rendered as a row of KPI cards like the web app."""
    t = doc.add_table(rows=2, cols=len(metrics))
    t.autofit = True
    for col, (label, value) in enumerate(metrics):
        vcell = t.cell(0, col)
        vcell.text = str(value)
        vp = vcell.paragraphs[0]; vp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        vr = vp.runs[0]; vr.font.bold = True; vr.font.size = Pt(18); vr.font.color.rgb = DOCX_NAVY
        lcell = t.cell(1, col)
        lcell.text = label
        lp = lcell.paragraphs[0]; lp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        lr = lp.runs[0]; lr.font.size = Pt(9); lr.font.color.rgb = DOCX_MUTED
        docx_shade_cell(vcell, "EAF1F9"); docx_shade_cell(lcell, "EAF1F9")
    return t


def pdf_level(v):
    """Mirrors static/app.js level(): qualitative reading of a capacity score."""
    if v is None: return "—"
    if v < 20: return ""
    if v <= 39: return "Loin en dessous de la moyenne"
    if v <= 59: return "En dessous de la moyenne"
    if v <= 70: return "Moyen"
    if v <= 80: return "Au-dessus de la moyenne"
    return "Bien au-dessus de la moyenne"


def pil_chart_grid(images, cols=2, pad=24, bg="white"):
    """Compose several chart images (as returned by the docx_*_chart helpers) into one
    dashboard image, so a report page shows several charts side by side instead of
    one chart per page. Each entry is either a PIL image, or an (image, caption) tuple
    whose caption is drawn above the chart (used for per-domain stat lines)."""
    items = [x if isinstance(x, tuple) else (x, None) for x in images if x is not None]
    if not items:
        return None
    cell_w = max(img.width for img, _ in items)
    cell_h = max(img.height for img, _ in items)
    caption_h = 34 if any(cap for _, cap in items) else 0
    rows = math.ceil(len(items) / cols)
    grid_w = cols * cell_w + pad * (cols + 1)
    grid_h = rows * (cell_h + caption_h) + pad * (rows + 1)
    canvas_img = PILImage.new("RGB", (grid_w, grid_h), bg)
    draw = ImageDraw.Draw(canvas_img)
    font = docx_font(20, True)
    for idx, (img, caption) in enumerate(items):
        r, col = divmod(idx, cols)
        x0 = pad + col * (cell_w + pad)
        y0 = pad + r * (cell_h + caption_h + pad)
        if caption:
            draw.text((x0 + cell_w / 2, y0 + caption_h / 2), caption, font=font, fill="black", anchor="mm")
        x = x0 + (cell_w - img.width) // 2
        y = y0 + caption_h + (cell_h - img.height) // 2
        canvas_img.paste(img, (x, y))
    return canvas_img


def xlsx_add_image(ws, row, col, img, scale=0.55):
    """Embed a PIL image in an xlsxwriter sheet at (row, col); returns the next free row."""
    if img is None:
        return row
    buf = BytesIO(); img.save(buf, format="PNG"); buf.seek(0)
    ws.insert_image(row, col, "chart.png", {"image_data": buf, "x_scale": scale, "y_scale": scale})
    return row + math.ceil(img.height * scale / 15) + 2


def report_docx(db, sid):
    """Mirrors the web app's 'Diagnostic terminé' page (the one 'Imprimer / Télécharger en
    PDF' prints) section for section: title/meta, KPI cards, Vue synthétique, Synthèse par
    domaine, Priorités retenues — same order, same headings, same table columns."""
    a = analysis(db, sid)
    if not a:
        raise ValueError("Session introuvable")
    session, g = a["session"], a["global"]
    domains = [d for d in a["domains"] if d.get("capacity") is not None]
    priorities = qualitative_data(db, sid)["priorities"]
    doc = Document()

    docx_style_heading(doc.add_heading("Diagnostic terminé", level=0))
    docx_note(doc, f"{session['name']} · {a['participantCount']} participants · {a['completedCount']} validés · taux {round(a['completedCount']/a['participantCount']*100) if a['participantCount'] else 0}%", size=11, italic=False)
    doc.add_paragraph()
    docx_metrics_row(doc, [
        ("Capacité", pdf_fmt(g["capacity"])),
        ("Consensus", pdf_fmt(g["consensus"])),
        ("Capacité graduée", g["gradedCapacity"] if g["gradedCapacity"] is not None else "—"),
        ("Consensus gradué", g["gradedConsensus"] if g["gradedConsensus"] is not None else "—"),
    ])

    if not domains or PILImage is None:
        docx_style_heading(doc.add_heading("Synthèse par domaine", level=2))
        t = doc.add_table(rows=1, cols=6)
        for c, x in zip(t.rows[0].cells, ["Domaine", "Capacité", "Consensus", "Graduées", "Niveau", "Réponses"]):
            c.text = x
        for d in domains:
            row = t.add_row().cells
            graded = f"{d['gradedCapacity'] if d['gradedCapacity'] is not None else '—'} / {d['gradedConsensus'] if d['gradedConsensus'] is not None else '—'}"
            for c, v in zip(row, [d["label"], pdf_fmt(d["capacity"]), pdf_fmt(d["consensus"]), graded, pdf_level(d["capacity"]), d["responses"]]):
                c.text = str(v)
        docx_style_table(t)
        if not domains:
            doc.add_paragraph("Pas encore de données suffisantes pour la restitution graphique EPC.")
        elif PILImage is None:
            doc.add_paragraph("Graphiques indisponibles : le paquet optionnel Pillow n'est pas installé (voir README.md).")
        docx_style_heading(doc.add_heading("Priorités retenues", level=2))
        doc.add_paragraph(f"{len(priorities)} priorité(s) sélectionnée(s)." if priorities else "Aucune priorité sélectionnée.")
        out = BytesIO(); doc.save(out); return out.getvalue()

    doc.add_page_break()
    docx_style_heading(doc.add_heading("Vue synthétique", level=2))
    docx_note(doc, "Radar global : vue complémentaire, ne remplace pas les graphiques EPC ci-dessous.")
    docx_add_chart(doc, pil_chart_grid([
        docx_radar_chart(domains),
        docx_bars_chart(domains, "standard", "Notes standardisées par domaine", with_mean=True),
        docx_bars_chart(domains, "graded", "Notes graduées par domaine", with_mean=True),
        docx_grid_chart(domains),
    ]))
    if len(domains) > 1:
        docx_add_chart(doc, docx_cohort_chart(domains))
    docx_note(doc, "20–39 : Loin en dessous · 40–59 : En dessous · 60–70 : Moyen · 71–80 : Au-dessus · 81–100 : Bien au-dessus", size=9, color=RGBColor(0x18, 0x6E, 0x42))

    doc.add_page_break()
    docx_style_heading(doc.add_heading("Synthèse par domaine", level=2))
    t = doc.add_table(rows=1, cols=6)
    for c, x in zip(t.rows[0].cells, ["Domaine", "Capacité", "Consensus", "Graduées", "Niveau", "Réponses"]):
        c.text = x
    for d in domains:
        row = t.add_row().cells
        graded = f"{d['gradedCapacity'] if d['gradedCapacity'] is not None else '—'} / {d['gradedConsensus'] if d['gradedConsensus'] is not None else '—'}"
        for c, v in zip(row, [d["label"], pdf_fmt(d["capacity"]), pdf_fmt(d["consensus"]), graded, pdf_level(d["capacity"]), d["responses"]]):
            c.text = str(v)
    docx_style_table(t)

    docx_style_heading(doc.add_heading("Priorités retenues", level=2))
    doc.add_paragraph(f"{len(priorities)} priorité(s) sélectionnée(s)." if priorities else "Aucune priorité sélectionnée.")

    out = BytesIO(); doc.save(out); return out.getvalue()

PDF_STOPWORDS = {"de", "des", "du", "la", "le", "les", "et", "en", "au", "aux", "à", "a", "l"}


def pdf_fmt(v):
    return "—" if v is None else f"{v:.1f}"


def pdf_short_label(item):
    """Mirrors static/app.js shortLabel(): configured code if it reads as a short
    alphabetic label, otherwise an abbreviation generated from the item's own name."""
    code = (item.get("code") or "").strip()
    if re.fullmatch(r"[A-Za-zÀ-ÿ]{2,10}", code):
        return code.upper()
    label = (item.get("label") or "").strip()
    words = [w for w in label.split() if w and w.lower().replace("’", "").replace("'", "") not in PDF_STOPWORDS]
    if not words:
        words = label.split()
    if len(words) >= 2:
        return "".join(w[0].upper() for w in words[:4])
    return label[:4].upper()



def read_xlsx(raw):
    ns={"x":"http://schemas.openxmlformats.org/spreadsheetml/2006/main","r":"http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
    try: z=zipfile.ZipFile(BytesIO(raw)); wb=ET.fromstring(z.read("xl/workbook.xml")); rels=ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    except Exception as e: raise ValueError("Le fichier n'est pas un XLSX valide") from e
    targets={r.attrib["Id"]:r.attrib["Target"].lstrip("/") for r in rels}; shared=[]
    if "xl/sharedStrings.xml" in z.namelist(): shared=["".join(x.itertext()) for x in ET.fromstring(z.read("xl/sharedStrings.xml")).findall("x:si",ns)]
    sheets={}
    for sh in wb.findall("x:sheets/x:sheet",ns):
        rid=sh.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]; target=targets[rid]; target="xl/"+target if not target.startswith("xl/") else target
        grid=[]
        for row in ET.fromstring(z.read(target)).findall(".//x:sheetData/x:row",ns):
            vals=[]
            for c in row.findall("x:c",ns):
                v=c.find("x:v",ns); text="" if v is None else v.text or ""
                if c.attrib.get("t")=="s": text=shared[int(text)]
                elif c.attrib.get("t")=="inlineStr": text="".join(c.itertext())
                vals.append(text)
            grid.append(vals)
        sheets[sh.attrib["name"].upper()]=grid
    return sheets


def import_preview(raw):
    sheets=read_xlsx(raw); errors=[]
    if "QUESTIONNAIRE" in sheets:
        q=sheets["QUESTIONNAIRE"]; p=sheets.get("PARAMETRES",[])
        if not q or q[0][:3] != ["Domaine","Référence","Indicateur qualitatif ou Capacité"]: raise ValueError("Colonnes QUESTIONNAIRE invalides")
        values={r[0]:r[1] if len(r)>1 else "" for r in p[:2] if r}; labels={}
        for r in p[4:]:
            if len(r)>1 and r[0]: labels[str(r[0])]=r[1]
        grouped={}; order=[]
        for r in q[1:]:
            r=(r+["", "", ""])[:3]
            if not any(r) or r[0].startswith("EXEMPLE"): continue
            if not all(r): errors.append("Chaque ligne doit contenir Domaine, Référence et Indicateur qualitatif ou Capacité"); continue
            if r[0] not in grouped: grouped[r[0]]=[]; order.append(r[0])
            grouped[r[0]].append({"code":"","label":r[1],"description":r[2],"response_type":"numeric","required":True,"active":True,"display_order":len(grouped[r[0]])})
        if not values.get("Nom du questionnaire","").strip(): errors.append("Nom du questionnaire obligatoire")
        return {"errors":errors,"template":{"name":values.get("Nom du questionnaire",""),"description":values.get("Description",""),"scale":{"type":"numeric","min":1,"max":5,"labels":labels},"priority":{"maxPerDomain":3},"domains":[{"label":d,"display_order":n+1,"indicators":grouped[d]} for n,d in enumerate(order)]},"rows":sum(len(v) for v in grouped.values())}
    if "INDICATEURS" not in sheets or "PARAMETRES" not in sheets: raise ValueError("Les feuilles INDICATEURS et PARAMETRES sont obligatoires")
    rows_i=sheets["INDICATEURS"]; header_index=next((n for n,r in enumerate(rows_i) if r[:len(MATRIX_COLUMNS)]==MATRIX_COLUMNS),None)
    if header_index is None: errors.append("Colonnes INDICATEURS invalides ou dans un ordre incorrect"); header_index=len(rows_i)
    params={r[0]:r[1] if len(r)>1 else "" for r in sheets["PARAMETRES"] if r and r[0] in PARAMETERS}
    for p in PARAMETERS:
        if p not in params: errors.append(f"Paramètre manquant : {p}")
    domains={}
    for n,r in enumerate(rows_i[header_index+1:],header_index+2):
        r=(r+[""]*9)[:9]; dom,do,code,label,desc,io,typ,required,active=r
        if not any(r): continue
        if code == "EXEMPLE-01": continue
        if not dom or not code or not label: errors.append(f"Ligne {n}: Domaine, Code indicateur et Indicateur sont obligatoires"); continue
        try: do=int(float(do)); io=int(float(io))
        except: errors.append(f"Ligne {n}: les ordres doivent être des nombres entiers"); continue
        if typ not in ("numeric","text","boolean"): errors.append(f"Ligne {n}: Type réponse invalide ({typ})"); continue
        def flag(v): return str(v).strip().lower() in ("oui","true","1","yes")
        key=(dom,do); domains.setdefault(key,[]).append({"code":code,"label":label,"description":desc,"response_type":typ,"required":flag(required),"active":flag(active),"display_order":io})
    for key, inds in domains.items():
        if len({i["code"] for i in inds}) != len(inds): errors.append(f"Domaine {key[0]}: codes indicateurs dupliqués")
        if len({i["display_order"] for i in inds}) != len(inds): errors.append(f"Domaine {key[0]}: ordres indicateurs dupliqués")
    try: lo=float(params.get("Valeur minimum",1)); hi=float(params.get("Valeur maximum",5)); assert lo<=hi
    except: errors.append("Bornes d'échelle invalides")
    if not params.get("Nom questionnaire","").strip(): errors.append("Nom questionnaire obligatoire")
    return {"errors":errors,"template":{"name":params.get("Nom questionnaire",""),"description":params.get("Description",""),"scale":{"type":params.get("Type d'échelle","numeric"),"min":lo if not errors else 1,"max":hi if not errors else 5,"labels":json.loads(params.get("Libellés des valeurs") or "{}")},"priority":{"maxPerDomain":int(float(params.get("Nombre de priorités par domaine",3) or 3))},"domains":[{"label":k[0],"display_order":k[1],"indicators":sorted(v,key=lambda x:x["display_order"])} for k,v in sorted(domains.items(),key=lambda x:x[0][1])]},"rows":sum(len(x) for x in domains.values())}


def save_import(db, data):
    if data["errors"]: raise ValueError("La matrice comporte des erreurs")
    tid=create_blank_template(db,data["template"]); t=data["template"]; db.execute("UPDATE templates SET scale_json=?,priority_json=? WHERE id=?",(json.dumps(t["scale"]),json.dumps(t["priority"]),tid))
    for d in t["domains"]:
        did=str(uuid.uuid4()); code="domain-"+uuid.uuid4().hex[:8]; db.execute("INSERT INTO domains VALUES (?,?,?,?,?,?,?)",(did,tid,code,d["label"],"",d["display_order"],1))
        for i in d["indicators"]: db.execute("INSERT INTO indicators VALUES (?,?,?,?,?,?,?,?,?,?)",(str(uuid.uuid4()),did,i["code"],i["label"],i["description"],i["response_type"],int(i["required"]),i["display_order"],int(i["active"]),"{}"))
    db.commit(); return tid


def grade(value, norm):
    """Classify value into a graduated band from norm (low, high, result) tuples.

    The bands are authored with integer bounds (e.g. 0-22, then 23-32), leaving a
    gap for any fractional value landing between two consecutive bounds (e.g.
    22.5) — capacity/consensus are continuous means, not integers, so that gap
    silently dropped real scores to None. Each band's effective upper edge is
    therefore extended up to (but excluding) the next band's low bound, instead
    of relying on the band's own recorded high, which only closes the gaps
    without changing any authored bound or result value.
    """
    if value is None:
        return None
    v = max(0.0, min(100.0, float(value)))
    n = len(norm)
    for i, (low, high, result) in enumerate(norm):
        if i + 1 < n:
            if low <= v < norm[i + 1][0]:
                return result
        elif low <= v <= high:
            return result
    return None


def analysis(db, session_id: str):
    session = db.execute("SELECT * FROM sessions WHERE id=?", (session_id,)).fetchone()
    if not session:
        return None
    template = template_payload(db, session["template_id"])
    scale, rules, consensus, norm = template["scale"], template["scoring"], template["consensus"], template["grading"]
    low, high, amplitude = float(scale["min"]), float(scale["max"]), float(scale["max"] - scale["min"])
    output_max = float(rules.get("outputRange", [0, 100])[1])
    all_values, output_domains = [], []
    for domain in template["domains"]:
        indicators = [i for i in domain["indicators"] if i["active"]]
        output_indicators, participant_means = [], {}
        for indicator in indicators:
            response_rows = rows(db, "SELECT participant_id,value_json FROM responses WHERE session_id=? AND indicator_id=?", (session_id, indicator["id"]))
            values = [float(json.loads(r["value_json"])) for r in response_rows if isinstance(json.loads(r["value_json"]), (int, float))]
            for r in response_rows:
                value = json.loads(r["value_json"])
                if isinstance(value, (int, float)):
                    participant_means.setdefault(r["participant_id"], []).append(float(value))
            mean = sum(values) / len(values) if values else None
            sd = (sum((v - mean) ** 2 for v in values) / (len(values) - 1)) ** .5 if len(values) > 1 else 0 if values else None
            capacity = mean / high * output_max if mean is not None else None
            cons = max(0, output_max - consensus["factor"] * (sd / amplitude * output_max)) if sd is not None else None
            output_indicators.append({"id": indicator["id"], "code": indicator["code"], "label": indicator["label"], "responses": len(values), "missing": max(0, len(rows(db, "SELECT id FROM participants WHERE session_id=?", (session_id,))) - len(values)), "mean": mean, "capacity": capacity, "dispersion": sd, "consensus": cons, "distribution": {str(k): values.count(k) for k in range(int(low), int(high) + 1)}})
            all_values.extend(values)
        person_scores = [sum(values) / len(values) for values in participant_means.values() if values]
        dmean = sum(person_scores) / len(person_scores) if person_scores else None
        dsd = (sum((v - dmean) ** 2 for v in person_scores) / (len(person_scores) - 1)) ** .5 if len(person_scores) > 1 else 0 if person_scores else None
        dcap = dmean / high * output_max if dmean is not None else None
        dcons = max(0, output_max - consensus["factor"] * (dsd / amplitude * output_max)) if dsd is not None else None
        output_domains.append({"id": domain["id"], "code": domain["code"], "label": domain["label"], "responses": len(person_scores), "capacity": dcap, "dispersion": dsd, "consensus": dcons, "gradedCapacity": grade(dcap, norm) if dcap is not None else None, "gradedConsensus": grade(dcons, norm) if dcons is not None else None, "indicators": output_indicators})
    global_mean = sum(all_values) / len(all_values) if all_values else None
    global_capacity = global_mean / high * 100 if global_mean is not None else None
    domain_caps=[d["capacity"] for d in output_domains if d["capacity"] is not None]; domain_cons=[d["consensus"] for d in output_domains if d["consensus"] is not None]
    gc=sum(domain_cons)/len(domain_cons) if domain_cons else None
    participants=len(rows(db,"SELECT id FROM participants WHERE session_id=?",(session_id,))); completed=len(rows(db,"SELECT id FROM participants WHERE session_id=? AND status='completed'",(session_id,)))
    return {"session": dict(session), "participantCount": participants, "completedCount":completed, "domains": output_domains, "global": {"responses": len(all_values), "capacity": global_capacity, "consensus":gc, "gradedCapacity": grade(global_capacity, norm) if global_capacity is not None else None, "gradedConsensus":grade(gc,norm) if gc is not None else None}}


def qualitative_data(db, session_id: str):
    """Qualitative workshop data kept separate from EPC numerical results."""
    priorities = rows(db, """SELECT p.*,d.label AS domain_label,i.code AS indicator_code,
        i.label AS indicator_label,i.description AS indicator_description
        FROM priorities p JOIN domains d ON d.id=p.domain_id JOIN indicators i ON i.id=p.indicator_id
        WHERE p.session_id=? ORDER BY d.display_order,i.display_order""", (session_id,))
    return {
        "priorities": priorities,
        "analyses": rows(db, "SELECT * FROM priority_analyses WHERE session_id=?", (session_id,)),
        "entries": rows(db, "SELECT * FROM analysis_entries WHERE session_id=? ORDER BY created_at", (session_id,)),
        "recommendations": rows(db, "SELECT * FROM workshop_recommendations WHERE session_id=? ORDER BY created_at", (session_id,)),
        "trainingTopics": rows(db, "SELECT * FROM training_topics WHERE session_id=? ORDER BY created_at", (session_id,)),
    }

def report_data(db, session_id: str):
    a=analysis(db,session_id)
    if not a: return None
    meta=db.execute("SELECT * FROM session_report_meta WHERE session_id=?",(session_id,)).fetchone()
    return {"analysis":a,"template":template_payload(db,a["session"]["template_id"]),"qualitative":qualitative_data(db,session_id),"meta":dict(meta) if meta else {"facilitator":"","audience":"","context":"","conclusion":""}}


class Handler(SimpleHTTPRequestHandler):
    def db(self):
        db = connect(); init_db(db); return db

    def json(self, code, payload):
        raw = json.dumps(payload, ensure_ascii=False).encode()
        self.send_response(code); self.send_header("Content-Type", "application/json; charset=utf-8"); self.send_header("Content-Length", str(len(raw))); self.end_headers(); self.wfile.write(raw)

    def body(self):
        size = int(self.headers.get("Content-Length", 0)); return json.loads(self.rfile.read(size) or b"{}")

    def raw_body(self): return self.rfile.read(int(self.headers.get("Content-Length", 0)))

    def do_GET(self):
        path, query = urlparse(self.path).path, parse_qs(urlparse(self.path).query)
        db = self.db()
        try:
            if path == "/api/templates": return self.json(200, rows(db, "SELECT id,name,version,description,status,created_at,updated_at FROM templates WHERE status='active' ORDER BY name,version DESC"))
            if path == "/api/templates/matrix.xlsx":
                data=blank_matrix_xlsx(); name=export_filename("matrice-questionnaire-vierge", ext="xlsx"); self.send_response(200); self.send_header("Content-Type","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"); self.send_header("Content-Disposition",f"attachment; filename={name}"); self.send_header("Content-Length",str(len(data))); self.end_headers(); self.wfile.write(data); return
            if path.startswith("/api/templates/") and path.endswith("/matrix.xlsx"):
                template=template_payload(db,path.split("/")[3]); data=matrix_xlsx(template); name=export_filename(template["name"],"matrice-questionnaire", ext="xlsx"); self.send_response(200); self.send_header("Content-Type","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"); self.send_header("Content-Disposition",f"attachment; filename={name}"); self.send_header("Content-Length",str(len(data))); self.end_headers(); self.wfile.write(data); return
            if path.startswith("/api/templates/"): return self.json(200, template_payload(db, path.rsplit("/", 1)[1]) or {"error": "Configuration introuvable"})
            if path == "/api/sessions": return self.json(200, rows(db, "SELECT * FROM sessions ORDER BY created_at DESC"))
            if path.startswith("/api/sessions/") and path.endswith("/analysis"):
                result = analysis(db, path.split("/")[3]); return self.json(200, result or {"error": "Session introuvable"})
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
            else: data=None
            if data is not None:
                self.send_response(200);self.send_header('Content-Type',mime);self.send_header('Content-Disposition',f'attachment; filename={name}');self.send_header('Content-Length',str(len(data)));self.send_header('Cache-Control','no-store');self.end_headers();self.wfile.write(data);return
            if path.startswith("/api/sessions/") and path.endswith("/responses.csv"):
                sid = path.split("/")[3]; buf = StringIO(); writer = csv.writer(buf); writer.writerow(["participant", "nom / organisation", "indicator", "value", "updated_at"])
                for r in db.execute("SELECT p.anonymous_id,p.display_name,i.code,r.value_json,r.updated_at FROM responses r JOIN participants p ON p.id=r.participant_id JOIN indicators i ON i.id=r.indicator_id WHERE r.session_id=?", (sid,)): writer.writerow(r)
                data = buf.getvalue().encode(); name=export_filename(session_label(db,sid),"reponses",ext="csv"); self.send_response(200); self.send_header("Content-Type", "text/csv; charset=utf-8"); self.send_header("Content-Disposition", f"attachment; filename={name}"); self.end_headers(); self.wfile.write(data); return
            if path == "/api/participant":
                sid, pid = query.get("session", [None])[0], query.get("participant", [None])[0]
                participant = db.execute("SELECT * FROM participants WHERE id=? AND session_id=?", (pid, sid)).fetchone()
                session = db.execute("SELECT * FROM sessions WHERE id=?", (sid,)).fetchone()
                return self.json(200, {"session": dict(session) if session else None, "participant": dict(participant) if participant else None, "template": template_payload(db, session["template_id"]) if session else None, "responses": {r["indicator_id"]: json.loads(r["value_json"]) for r in db.execute("SELECT * FROM responses WHERE participant_id=?", (pid,))}})
            return self.serve_static(path)
        finally: db.close()

    def do_POST(self):
        path = urlparse(self.path).path
        if path == "/api/templates/import/preview":
            raw=self.raw_body(); marker=b"\r\n\r\n"; start=raw.find(marker)+len(marker); end=raw.rfind(b"\r\n--"); uploaded=raw[start:end]
            try:
                preview=import_preview(uploaded); token=str(uuid.uuid4()); IMPORTS[token]=preview; return self.json(200,{"token":token,**preview})
            except ValueError as e: return self.json(400,{"error":str(e)})
        data, db = self.body(), self.db()
        try:
            if path == "/api/templates": return self.json(201,{"id":create_blank_template(db,data)})
            if path == "/api/templates/import/confirm":
                preview=IMPORTS.pop(data.get("token"),None)
                if not preview: return self.json(400,{"error":"Aperçu d'import introuvable ou expiré"})
                return self.json(201,{"id":save_import(db,preview)})
            if path.startswith("/api/templates/") and path.endswith("/clone"):
                return self.json(201,{"id":clone_template(db,path.split("/")[3],data.get("name"))})
            if path.startswith("/api/templates/") and path.endswith("/domains"):
                tid=path.split("/")[3]; did=str(uuid.uuid4()); code=data.get("code") or "domain-"+uuid.uuid4().hex[:8]; db.execute("INSERT INTO domains VALUES (?,?,?,?,?,?,?)",(did,tid,code,data["label"],data.get("description",""),int(data.get("displayOrder") or next_order(db,"domains","template_id",tid)),int(data.get("active",True)))); db.commit(); return self.json(201,{"id":did})
            if path.startswith("/api/domains/") and path.endswith("/indicators"):
                did=path.split("/")[3]; iid=str(uuid.uuid4()); db.execute("INSERT INTO indicators VALUES (?,?,?,?,?,?,?,?,?,?)",(iid,did,data.get("code") or "indicator-"+uuid.uuid4().hex[:8],data["label"],data.get("description",""),data.get("responseType","numeric"),int(data.get("required",True)),int(data.get("displayOrder") or next_order(db,"indicators","domain_id",did)),int(data.get("active",True)),json.dumps(data.get("configuration",{})))); db.commit(); return self.json(201,{"id":iid})
            if path == "/api/sessions":
                template = template_payload(db, data["templateId"])
                if not template or not any(d["active"] and any(i["active"] for i in d["indicators"]) for d in template["domains"]): return self.json(400,{"error":"Impossible de créer une session : le questionnaire ne contient aucun domaine avec question."})
                sid = str(uuid.uuid4()); db.execute("INSERT INTO sessions VALUES (?,?,?,?,?,?,?,?,?,?)", (sid, template["id"], template["version"], data["name"], data.get("organization", ""), data.get("location", ""), data.get("date", ""), "open", now(), None)); db.commit(); return self.json(201, {"id": sid})
            if path.endswith("/participants"):
                sid = path.split("/")[3]; session = db.execute("SELECT status FROM sessions WHERE id=?", (sid,)).fetchone()
                if not session or session["status"] != "open": return self.json(409, {"error": "Collecte fermée"})
                pid, label = str(uuid.uuid4()), data.get("anonymousId") or f"P-{uuid.uuid4().hex[:6]}"; db.execute("INSERT INTO participants VALUES (?,?,?,?,?,?,?)", (pid, sid, label, "in_progress", now(), None, data.get("displayName") or None)); db.commit(); return self.json(201, {"id": pid, "anonymousId": label})
            if path.endswith("/responses"):
                sid = path.split("/")[3]; stamp = now(); db.execute("INSERT INTO responses VALUES (?,?,?,?,?,?,?,?) ON CONFLICT(participant_id,indicator_id) DO UPDATE SET value_json=excluded.value_json,value_type=excluded.value_type,updated_at=excluded.updated_at", (str(uuid.uuid4()), sid, data["participantId"], data["indicatorId"], json.dumps(data["value"]), data.get("valueType", "numeric"), stamp, stamp)); db.commit(); return self.json(200, {"ok": True})
            if path.endswith("/complete"):
                pid = data["participantId"]; db.execute("UPDATE participants SET status='completed',completed_at=? WHERE id=?", (now(), pid)); db.commit(); return self.json(200, {"ok": True})
            if path.endswith("/status"):
                sid = path.split("/")[3]; status = data["status"]; db.execute("UPDATE sessions SET status=?,closed_at=? WHERE id=?", (status, now() if status == "closed" else None, sid)); db.commit(); return self.json(200, {"ok": True})
            if path.endswith("/priorities"):
                sid = path.split("/")[3]; db.execute("INSERT INTO priorities VALUES (?,?,?,?,?,?) ON CONFLICT(session_id,indicator_id) DO UPDATE SET votes=excluded.votes", (str(uuid.uuid4()), sid, data["domainId"], data["indicatorId"], int(data.get("votes", 0)), now())); db.commit(); return self.json(200, {"ok": True})
            if path.endswith("/priority-analyses"):
                sid=path.split("/")[3]; stamp=now(); priority_id=data["priorityId"]
                db.execute("INSERT INTO priority_analyses VALUES (?,?,?,?,?,?) ON CONFLICT(session_id,priority_id) DO UPDATE SET problem=excluded.problem,updated_at=excluded.updated_at",(str(uuid.uuid4()),sid,priority_id,data.get("problem",""),stamp,stamp)); db.commit(); return self.json(201,{"ok":True})
            if path.endswith("/analysis-entries"):
                sid=path.split("/")[3]; stamp=now(); eid=str(uuid.uuid4())
                db.execute("INSERT INTO analysis_entries VALUES (?,?,?,?,?,?,?,?,?,?,?)",(eid,sid,data["priorityId"],data.get("parentId") or None,data["kind"],data["content"],data.get("itemType") or None,data.get("comment") or None,data.get("validationStatus","A_DISCUTER"),stamp,stamp)); db.commit(); return self.json(201,{"id":eid})
            if path.endswith("/recommendations-v2"):
                sid=path.split("/")[3]; stamp=now(); rid=str(uuid.uuid4())
                db.execute("INSERT INTO workshop_recommendations VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",(rid,sid,data.get("priorityId") or None,data.get("causeId") or None,data.get("leverId") or None,data["title"],data["description"],data.get("category","Autre"),data.get("priorityLevel","Non définie"),data.get("owner") or None,data.get("horizon") or None,data.get("comment") or None,data.get("status","Proposée"),stamp,stamp)); db.commit(); return self.json(201,{"id":rid})
            if path.endswith("/training-topics"):
                sid=path.split("/")[3]; stamp=now(); tid=str(uuid.uuid4())
                db.execute("INSERT INTO training_topics VALUES (?,?,?,?,?,?,?,?,?,?,?)",(tid,sid,data.get("priorityId") or None,data.get("recommendationId") or None,data["title"],data.get("needText") or None,data.get("targetAudience") or None,data.get("priorityLevel","Non définie"),data.get("comment") or None,stamp,stamp)); db.commit(); return self.json(201,{"id":tid})
            if path.endswith("/report-meta"):
                sid=path.split("/")[3]; db.execute("INSERT INTO session_report_meta VALUES (?,?,?,?,?,?) ON CONFLICT(session_id) DO UPDATE SET facilitator=excluded.facilitator,audience=excluded.audience,context=excluded.context,conclusion=excluded.conclusion,updated_at=excluded.updated_at",(sid,data.get("facilitator",""),data.get("audience",""),data.get("context",""),data.get("conclusion",""),now())); db.commit(); return self.json(200,{"ok":True})
            if path.endswith("/analysis-notes"):
                sid = path.split("/")[3]; stamp=now(); db.execute("INSERT INTO analysis_notes VALUES (?,?,?,?,?,?,?,?,?)", (str(uuid.uuid4()), sid, data.get("indicatorId"), data["kind"], data["content"], data.get("validationStatus", "HYPOTHESE"), stamp, stamp)); db.commit(); return self.json(201, {"ok": True})
            if path.endswith("/recommendations"):
                sid=path.split("/")[3]; db.execute("INSERT INTO recommendations VALUES (?,?,?,?,?,?,?,?,?,?)",(str(uuid.uuid4()),sid,data.get("indicatorId"),data["title"],data.get("description",""),data.get("lever",""),data.get("kind","action"),data.get("owner",""),data.get("horizon",""),now())); db.commit(); return self.json(201,{"ok":True})
            return self.json(404, {"error": "Route inconnue"})
        finally: db.close()

    def do_PUT(self):
        path, data, db = urlparse(self.path).path, self.body(), self.db()
        try:
            if path.startswith("/api/sessions/"):
                sid=path.split("/")[3]; db.execute("UPDATE sessions SET name=?,organization=?,location=?,date=? WHERE id=?",(data["name"],data.get("organization",''),data.get("location",''),data.get("date",''),sid)); db.commit(); return self.json(200,{"ok":True})
            if path.startswith("/api/participants/"):
                pid=path.split("/")[3]; db.execute("UPDATE participants SET display_name=? WHERE id=?",(data.get("displayName") or None,pid)); db.commit(); return self.json(200,{"ok":True})
            if path.startswith("/api/priority-analyses/"):
                db.execute("UPDATE priority_analyses SET problem=?,updated_at=? WHERE id=?",(data.get("problem",""),now(),path.split("/")[3])); db.commit(); return self.json(200,{"ok":True})
            if path.startswith("/api/analysis-entries/"):
                db.execute("UPDATE analysis_entries SET parent_id=?,content=?,item_type=?,comment=?,validation_status=?,updated_at=? WHERE id=?",(data.get("parentId") or None,data["content"],data.get("itemType") or None,data.get("comment") or None,data.get("validationStatus","A_DISCUTER"),now(),path.split("/")[3])); db.commit(); return self.json(200,{"ok":True})
            if path.startswith("/api/recommendations-v2/"):
                db.execute("UPDATE workshop_recommendations SET priority_id=?,cause_id=?,lever_id=?,title=?,description=?,category=?,priority_level=?,owner=?,horizon=?,comment=?,status=?,updated_at=? WHERE id=?",(data.get("priorityId") or None,data.get("causeId") or None,data.get("leverId") or None,data["title"],data["description"],data.get("category","Autre"),data.get("priorityLevel","Non définie"),data.get("owner") or None,data.get("horizon") or None,data.get("comment") or None,data.get("status","Proposée"),now(),path.split("/")[3])); db.commit(); return self.json(200,{"ok":True})
            if path.startswith("/api/training-topics/"):
                db.execute("UPDATE training_topics SET priority_id=?,recommendation_id=?,title=?,need_text=?,target_audience=?,priority_level=?,comment=?,updated_at=? WHERE id=?",(data.get("priorityId") or None,data.get("recommendationId") or None,data["title"],data.get("needText") or None,data.get("targetAudience") or None,data.get("priorityLevel","Non définie"),data.get("comment") or None,now(),path.split("/")[3])); db.commit(); return self.json(200,{"ok":True})
            if path.startswith("/api/templates/"):
                tid=path.split("/")[3]; used=db.execute("SELECT 1 FROM sessions WHERE template_id=? LIMIT 1",(tid,)).fetchone()
                if used:
                    tid=clone_template(db,tid); data["versionCreated"]=True
                old=template_payload(db,tid); scale=data.get("scale",old["scale"]); db.execute("UPDATE templates SET name=?,description=?,scale_json=?,priority_json=?,updated_at=? WHERE id=?",(data.get("name",old["name"]),data.get("description",old["description"]),json.dumps(scale),json.dumps(data.get("priority",old["priority"])),now(),tid)); db.commit(); return self.json(200,{"id":tid,"versionCreated":data.get("versionCreated",False)})
            if path.startswith("/api/domains/"):
                did=path.split("/")[3]; db.execute("UPDATE domains SET label=?,description=?,display_order=?,active=? WHERE id=?",(data["label"],data.get("description",""),int(data.get("displayOrder",1)),int(data.get("active",True)),did)); db.commit(); return self.json(200,{"ok":True})
            if path.startswith("/api/indicators/"):
                iid=path.split("/")[3]; db.execute("UPDATE indicators SET domain_id=?,code=?,label=?,description=?,response_type=?,required=?,display_order=?,active=?,configuration_json=? WHERE id=?",(data["domainId"],data["code"],data["label"],data.get("description",""),data.get("responseType","numeric"),int(data.get("required",True)),int(data.get("displayOrder",1)),int(data.get("active",True)),json.dumps(data.get("configuration",{})),iid)); db.commit(); return self.json(200,{"ok":True})
            return self.json(404,{"error":"Route inconnue"})
        finally: db.close()

    def do_DELETE(self):
        path, db=urlparse(self.path).path,self.db()
        try:
            if path.startswith("/api/analysis-entries/"):
                eid=path.split("/")[3]; dependent=db.execute("SELECT COUNT(*) FROM workshop_recommendations WHERE cause_id=? OR lever_id=?",(eid,eid)).fetchone()[0]
                if dependent and parse_qs(urlparse(self.path).query).get("force",["0"])[0] != "1": return self.json(409,{"error":f"Cette entrée est utilisée par {dependent} recommandation(s). Confirmez la suppression.","dependencies":dependent})
                db.execute("UPDATE workshop_recommendations SET cause_id=NULL WHERE cause_id=?",(eid,)); db.execute("UPDATE workshop_recommendations SET lever_id=NULL WHERE lever_id=?",(eid,)); db.execute("UPDATE analysis_entries SET parent_id=NULL WHERE parent_id=?",(eid,)); db.execute("DELETE FROM analysis_entries WHERE id=?",(eid,)); db.commit(); return self.json(200,{"ok":True})
            if path.startswith("/api/recommendations-v2/"):
                rid=path.split("/")[3]; dependent=db.execute("SELECT COUNT(*) FROM training_topics WHERE recommendation_id=?",(rid,)).fetchone()[0]
                if dependent and parse_qs(urlparse(self.path).query).get("force",["0"])[0] != "1": return self.json(409,{"error":f"Cette recommandation est liée à {dependent} thème(s) de formation. Confirmez la suppression.","dependencies":dependent})
                db.execute("UPDATE training_topics SET recommendation_id=NULL WHERE recommendation_id=?",(rid,)); db.execute("DELETE FROM workshop_recommendations WHERE id=?",(rid,)); db.commit(); return self.json(200,{"ok":True})
            if path.startswith("/api/training-topics/"):
                db.execute("DELETE FROM training_topics WHERE id=?",(path.split("/")[3],)); db.commit(); return self.json(200,{"ok":True})
            if path.startswith("/api/templates/"):
                tid=path.split("/")[3]
                protected=db.execute("SELECT name FROM templates WHERE id=?",(tid,)).fetchone()
                if protected and protected["name"] == "EPC / SENEVAL": return self.json(409,{"error":"Suppression impossible. EPC/SENEVAL est le modèle de référence ; dupliquez-le pour le modifier."})
                if db.execute("SELECT 1 FROM sessions WHERE template_id=? LIMIT 1",(tid,)).fetchone():
                    if parse_qs(urlparse(self.path).query).get("force",["0"])[0] == "1": db.execute("UPDATE templates SET status='archived',updated_at=? WHERE id=?",(now(),tid)); db.commit(); return self.json(200,{"ok":True,"archived":True})
                    return self.json(409,{"error":"Suppression impossible. Ce questionnaire est utilisé par une ou plusieurs sessions d’atelier. Vous pouvez le conserver, créer une nouvelle version, ou confirmer son retrait de la liste des modèles."})
                db.execute("DELETE FROM indicators WHERE domain_id IN (SELECT id FROM domains WHERE template_id=?)",(tid,)); db.execute("DELETE FROM domains WHERE template_id=?",(tid,)); db.execute("DELETE FROM templates WHERE id=?",(tid,)); db.commit(); return self.json(200,{"ok":True})
            if path.startswith("/api/indicators/"):
                iid=path.split("/")[3]; db.execute("DELETE FROM indicators WHERE id=?",(iid,)); db.commit(); return self.json(200,{"ok":True})
            if path.startswith("/api/sessions/") and "/priorities/" in path:
                parts=path.split("/"); db.execute("DELETE FROM priorities WHERE session_id=? AND indicator_id=?",(parts[3],parts[5])); db.commit(); return self.json(200,{"ok":True})
            if path.startswith("/api/sessions/") and len(path.rstrip("/").split("/")) == 4:
                sid=path.rstrip("/").split("/")[3]
                for table in ("training_topics","workshop_recommendations","analysis_entries","priority_analyses","analysis_notes","recommendations","responses","priorities","participants","session_report_meta"):
                    db.execute(f"DELETE FROM {table} WHERE session_id=?",(sid,))
                db.execute("DELETE FROM sessions WHERE id=?",(sid,)); db.commit(); return self.json(200,{"ok":True})
            return self.json(404,{"error":"Route inconnue"})
        finally: db.close()

    def serve_static(self, path):
        requested = "index.html" if path in ("/", "") else path.lstrip("/")
        target = (STATIC / requested).resolve()
        if not str(target).startswith(str(STATIC.resolve())) or not target.is_file(): self.send_error(404); return
        self.send_response(200); self.send_header("Content-Type", "text/html; charset=utf-8" if target.suffix == ".html" else "application/javascript; charset=utf-8" if target.suffix == ".js" else "text/css; charset=utf-8"); self.send_header("Cache-Control","no-store, max-age=0"); self.end_headers(); self.wfile.write(target.read_bytes())


def main():
    db=connect(); init_db(db); db.close()
    server=ThreadingHTTPServer(("127.0.0.1", 8000), Handler)
    print("EPC Workshop Engine: http://127.0.0.1:8000")
    server.serve_forever()

if __name__ == "__main__": main()
