"""Local workshop diagnosis engine.

Run with: python app.py
The application uses only the Python standard library and SQLite.  It is meant
to be a dependable local-first starting point, not a simulated real-time app.
"""
from __future__ import annotations

import csv
import base64
import json
import sqlite3
import sys
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
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas
except ImportError:
    Document = None
    canvas = None

ROOT = Path(__file__).resolve().parent
STATIC = ROOT / "static"
DATABASE = ROOT / "data" / "workshops.sqlite3"
IMPORTS = {}
MATRIX_COLUMNS = ["Domaine", "Ordre domaine", "Code indicateur", "Indicateur", "Description", "Ordre indicateur", "Type réponse", "Obligatoire", "Actif"]
PARAMETERS = ["Nom questionnaire", "Description", "Version", "Type d'échelle", "Valeur minimum", "Valeur maximum", "Libellés des valeurs", "Nombre de priorités par domaine"]


def now() -> str:
    return datetime.now(timezone.utc).isoformat()


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
    CREATE TABLE IF NOT EXISTS participants (id TEXT PRIMARY KEY, session_id TEXT NOT NULL REFERENCES sessions(id), anonymous_id TEXT NOT NULL, status TEXT NOT NULL, started_at TEXT NOT NULL, completed_at TEXT, UNIQUE(session_id, anonymous_id));
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
    guide=wb.add_worksheet("MODE D’EMPLOI"); guide.set_column(0,0,110); guide.write("A1","Cette matrice permet de préparer un questionnaire avant de l’importer dans l’outil.",head); guide.write_column("A3",["1. Donnez un nom au questionnaire dans la feuille PARAMETRES.","2. Définissez une seule fois l’échelle de notation.","3. Dans la feuille QUESTIONNAIRE, saisissez une ligne par indicateur.","4. Répétez le nom du domaine pour les indicateurs appartenant au même domaine.","5. La numérotation sera générée automatiquement par l’outil.","6. Les lignes d’exemple peuvent être remplacées ou supprimées."],wrap)
    ps=wb.add_worksheet("PARAMETRES"); ps.write_row(0,0,["Nom du questionnaire",template["name"]],head); ps.write_row(1,0,["Description",template["description"]],wrap); ps.write_row(3,0,["Note","Libellé"],head); labels=template["scale"].get("labels",{}); [ps.write_row(4+(5-n),0,[n,labels.get(str(n),"")]) for n in range(5,0,-1)]; ps.set_column(0,0,26); ps.set_column(1,1,55)
    ws=wb.add_worksheet("QUESTIONNAIRE"); ws.write_row(0,0,["Domaine","Référence","Indicateur qualitatif ou Capacité"],head); ws.freeze_panes(1,0); row=1
    if not template["domains"]: ws.write_row(row,0,["EXEMPLE — Gestion des Ressources Humaines","EXEMPLE — Formation au personnel","EXEMPLE — Nous offrons régulièrement la formation au personnel"],wrap); row+=1
    for d in template["domains"]:
        for i in d["indicators"]:
            ws.write_row(row,0,[d["label"],i["label"],i["description"]],wrap); row+=1
    ws.set_column(0,0,34); ws.set_column(1,1,38); ws.set_column(2,2,75); wb.close(); return out.getvalue()


def blank_matrix_xlsx():
    return matrix_xlsx({"name":"Nouveau questionnaire", "description":"", "version":1, "scale":{"type":"numeric","min":1,"max":5,"labels":{}}, "priority":{"maxPerDomain":3}, "domains":[]})

def report_rows(db, sid):
    a=analysis(db,sid); return a, [[d['label'],d['capacity'],d['consensus'],d['gradedCapacity'],d['gradedConsensus'],d['responses']] for d in a['domains']]

def report_xlsx(db,sid):
    a,rs=report_rows(db,sid); q=qualitative_data(db,sid); meta=report_data(db,sid)["meta"]; template=template_payload(db,a['session']['template_id']); out=BytesIO(); wb=xlsxwriter.Workbook(out,{"in_memory":True}); h=wb.add_format({"bold":True,"bg_color":"#1F4E78","font_color":"#FFFFFF"})
    analyses={x['priority_id']:x for x in q['analyses']}; priority_rows=[[p['id'],p['domain_label'],p['indicator_code'],p['indicator_label'],analyses.get(p['id'],{}).get('problem','')] for p in q['priorities']]
    sheets=[("Synthèse",["Atelier","Organisation","Lieu","Date","Animateur","Public","Contexte","Conclusion","Capacité","Consensus"],[[a['session']['name'],a['session']['organization'],a['session']['location'],a['session']['date'],meta['facilitator'],meta['audience'],meta['context'],meta['conclusion'],a['global']['capacity'],a['global']['consensus']]]),("Domaines",["Domaine","Capacité","Consensus","Cap. graduée","Cons. gradué","Réponses"],rs),("Indicateurs",["Domaine","Référence","Capacité","Consensus","Réponses","Manquants"],[[d['label'],i['label'],i['capacity'],i['consensus'],i['responses'],i['missing']] for d in a['domains'] for i in d['indicators']]),("Priorités",["ID priorité","Domaine","Référence","Indicateur","Constat"],priority_rows),("Analyses",["ID","Priorité","Constat"],[[x['id'],x['priority_id'],x['problem']] for x in q['analyses']]),("Causes",["ID","Priorité","Parent","Cause","Type","Statut"],[[x['id'],x['priority_id'],x['parent_id'],x['content'],x['item_type'],x['validation_status']] for x in q['entries'] if x['kind']=='cause']),("Conséquences",["ID","Priorité","Conséquence","Statut"],[[x['id'],x['priority_id'],x['content'],x['validation_status']] for x in q['entries'] if x['kind']=='consequence']),("Leviers",["ID","Priorité","Levier","Commentaire","Statut"],[[x['id'],x['priority_id'],x['content'],x['comment'],x['validation_status']] for x in q['entries'] if x['kind']=='lever']),("Recommandations",["ID","Priorité","Cause","Levier","Titre","Description","Catégorie","Niveau","Responsable","Échéance","Statut"],[[x['id'],x['priority_id'],x['cause_id'],x['lever_id'],x['title'],x['description'],x['category'],x['priority_level'],x['owner'],x['horizon'],x['status']] for x in q['recommendations']]),("Formations",["ID","Priorité","Recommandation","Intitulé","Besoin","Public","Niveau","Commentaire"],[[x['id'],x['priority_id'],x['recommendation_id'],x['title'],x['need_text'],x['target_audience'],x['priority_level'],x['comment']] for x in q['trainingTopics']]),("Plan_action",["N°","Action / recommandation","Origine","Responsable","Échéance","Priorité","Statut"],[[n+1,x['title'],x['priority_id'] or '—',x['owner'] or '—',x['horizon'] or '—',x['priority_level'],x['status']] for n,x in enumerate(q['recommendations']) if x['status']=='Retenue']),("Questionnaire",["Domaine","Référence","Indicateur","Échelle"],[[d['label'],i['label'],i['description'],f"{template['scale']['min']}–{template['scale']['max']}"] for d in template['domains'] for i in d['indicators'] if i['active']])]
    for name,head,data in sheets:
        s=wb.add_worksheet(name);s.write_row(0,0,head,h);[s.write_row(n+1,0,row) for n,row in enumerate(data)];s.set_column(0,len(head)-1,24)
    wb.close();return out.getvalue()

def report_docx(db,sid):
    a,rs=report_rows(db,sid); doc=Document();doc.add_heading('Diagnostic EPC',0);doc.add_paragraph(a['session']['name']);doc.add_paragraph(f"Capacité : {a['global']['capacity']:.1f} — Consensus : {a['global']['consensus']:.1f}");t=doc.add_table(rows=1,cols=4);[setattr(c,'text',x) for c,x in zip(t.rows[0].cells,['Domaine','Capacité','Consensus','Graduée'])];[ [setattr(c,'text',str(v if v is not None else '—')) for c,v in zip(t.add_row().cells,[r[0],round(r[1],1),round(r[2],1),r[3]])] for r in rs];out=BytesIO();doc.save(out);return out.getvalue()

def report_pdf(db,sid):
    a,rs=report_rows(db,sid);out=BytesIO();c=canvas.Canvas(out,pagesize=A4);c.setFont('Helvetica-Bold',16);c.drawString(45,800,'Diagnostic EPC');c.setFont('Helvetica',11);c.drawString(45,780,a['session']['name']);c.drawString(45,760,f"Capacité {a['global']['capacity']:.1f} - Consensus {a['global']['consensus']:.1f}");y=730
    for r in rs:
        c.drawString(45,y,f"{r[0]} : capacité {r[1]:.1f} - consensus {r[2]:.1f}"); y-=22
    c.save();return out.getvalue()


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
    for low, high, result in norm:
        if low <= value <= high:
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
                data=blank_matrix_xlsx(); self.send_response(200); self.send_header("Content-Type","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"); self.send_header("Content-Disposition","attachment; filename=matrice-questionnaire-vierge.xlsx"); self.send_header("Content-Length",str(len(data))); self.end_headers(); self.wfile.write(data); return
            if path.startswith("/api/templates/") and path.endswith("/matrix.xlsx"):
                data=matrix_xlsx(template_payload(db,path.split("/")[3])); self.send_response(200); self.send_header("Content-Type","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"); self.send_header("Content-Disposition","attachment; filename=matrice-questionnaire.xlsx"); self.send_header("Content-Length",str(len(data))); self.end_headers(); self.wfile.write(data); return
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
                sid=path.split("/")[3];data=report_xlsx(db,sid);mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';name='diagnostic.xlsx'
            elif path.startswith("/api/sessions/") and path.endswith("/report.docx"):
                sid=path.split("/")[3];data=report_docx(db,sid);mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document';name='diagnostic.docx'
            elif path.startswith("/api/sessions/") and path.endswith("/report.pdf"):
                sid=path.split("/")[3];data=report_pdf(db,sid);mime='application/pdf';name='diagnostic.pdf'
            else: data=None
            if data is not None:
                self.send_response(200);self.send_header('Content-Type',mime);self.send_header('Content-Disposition',f'attachment; filename={name}');self.send_header('Content-Length',str(len(data)));self.send_header('Cache-Control','no-store');self.end_headers();self.wfile.write(data);return
            if path.startswith("/api/sessions/") and path.endswith("/responses.csv"):
                sid = path.split("/")[3]; buf = StringIO(); writer = csv.writer(buf); writer.writerow(["participant", "indicator", "value", "updated_at"])
                for r in db.execute("SELECT p.anonymous_id,i.code,r.value_json,r.updated_at FROM responses r JOIN participants p ON p.id=r.participant_id JOIN indicators i ON i.id=r.indicator_id WHERE r.session_id=?", (sid,)): writer.writerow(r)
                data = buf.getvalue().encode(); self.send_response(200); self.send_header("Content-Type", "text/csv; charset=utf-8"); self.send_header("Content-Disposition", "attachment; filename=responses.csv"); self.end_headers(); self.wfile.write(data); return
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
                pid, label = str(uuid.uuid4()), data.get("anonymousId") or f"P-{uuid.uuid4().hex[:6]}"; db.execute("INSERT INTO participants VALUES (?,?,?,?,?,?)", (pid, sid, label, "in_progress", now(), None)); db.commit(); return self.json(201, {"id": pid, "anonymousId": label})
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
