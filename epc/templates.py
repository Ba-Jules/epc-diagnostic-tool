"""Questionnaires : clonage/creation, export de la matrice XLSX, import depuis
une matrice XLSX (lecture, previsualisation, confirmation).

Extrait de app.py (lot 1c de la modularisation, cf. AUDIT_MODULARISATION_8800.md) :
aucune formule ni aucun contrat de route ne change ici, seule l'emplacement du
code metier questionnaire bouge. app.py reimporte ces symboles a l'identique.

IMPORTS est un dictionnaire en memoire process-local (token -> previsualisation
d'import), tel qu'il l'etait dans app.py : la confirmation d'import doit se
faire sur le meme process que la previsualisation.
"""
from __future__ import annotations

import json
import uuid
import zipfile
from io import BytesIO
from xml.etree import ElementTree as ET

try:  # Used only to generate the downloadable Excel template; the app stays local.
    import xlsxwriter
except ImportError:
    xlsxwriter = None

from .db import GRADING, now, rows, template_payload

MATRIX_COLUMNS = ["Domaine", "Ordre domaine", "Code indicateur", "Indicateur", "Description", "Ordre indicateur", "Type réponse", "Obligatoire", "Actif"]
PARAMETERS = ["Nom questionnaire", "Description", "Version", "Type d'échelle", "Valeur minimum", "Valeur maximum", "Libellés des valeurs", "Nombre de priorités par domaine"]

IMPORTS = {}


def next_order(db, table, where_col, where_value):
    return (db.execute(f"SELECT COALESCE(MAX(display_order),0)+1 FROM {table} WHERE {where_col}=?", (where_value,)).fetchone()[0])


def clone_template(db, template_id, name=None):
    old = template_payload(db, template_id)
    if not old: raise ValueError("Configuration introuvable")
    tid, stamp = str(uuid.uuid4()), now()
    version = db.execute("SELECT COALESCE(MAX(version),0)+1 FROM templates WHERE name=?", (name or old["name"],)).fetchone()[0]
    db.execute("INSERT INTO templates VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", (tid, name or old["name"], version, old["description"], "active", json.dumps(old["scale"]), json.dumps(old["scoring"]), json.dumps(old["consensus"]), json.dumps(old["grading"]), json.dumps(old["priority"]), stamp, stamp, old.get("owner_user_id")))
    for d in old["domains"]:
        did=str(uuid.uuid4()); db.execute("INSERT INTO domains VALUES (?,?,?,?,?,?,?)",(did,tid,d["code"],d["label"],d["description"],d["display_order"],d["active"]))
        for i in d["indicators"]:
            db.execute("INSERT INTO indicators VALUES (?,?,?,?,?,?,?,?,?,?)",(str(uuid.uuid4()),did,i["code"],i["label"],i["description"],i["response_type"],i["required"],i["display_order"],i["active"],json.dumps(i["configuration"])))
    db.commit(); return tid


def create_blank_template(db, data, owner_user_id=None):
    tid, stamp = str(uuid.uuid4()), now(); name=data.get("name", "Nouveau questionnaire").strip()
    if not name: raise ValueError("Le nom est obligatoire")
    version=db.execute("SELECT COALESCE(MAX(version),0)+1 FROM templates WHERE name=?",(name,)).fetchone()[0]
    scale={"type":"numeric","min":1,"max":5,"labels":{}}
    db.execute("INSERT INTO templates VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)",(tid,name,version,data.get("description", ""),"active",json.dumps(scale),json.dumps({"capacity":"mean_divided_by_scale_max","outputRange":[0,100]}),json.dumps({"method":"standard_deviation","normalization":"theoretical_range","factor":2}),json.dumps(GRADING),json.dumps({"maxPerDomain":3}),stamp,stamp,owner_user_id)); db.commit(); return tid


def matrix_xlsx(template):
    if not xlsxwriter: raise RuntimeError("Le générateur XLSX local n'est pas disponible")
    out=BytesIO(); wb=xlsxwriter.Workbook(out, {"in_memory": True}); head=wb.add_format({"bold":True,"bg_color":"#1F4E78","font_color":"#FFFFFF"}); wrap=wb.add_format({"text_wrap":True,"valign":"top"})
    guide=wb.add_worksheet("MODE D’EMPLOI"); guide.set_column(0,0,110); guide.write("A1","Cette matrice permet de préparer un questionnaire avant de l’importer dans l’outil.",head); guide.write_column("A3",["1. Dans la feuille PARAMETRES, remplacez la valeur d’exemple par le vrai nom de votre questionnaire.","2. Complétez la description (facultatif) et les libellés de l’échelle de notation (une seule fois, valables pour tout le questionnaire).","3. Dans la feuille QUESTIONNAIRE, saisissez une ligne par indicateur.","4. Répétez le nom du domaine pour les indicateurs appartenant au même domaine.","5. La numérotation sera générée automatiquement par l’outil.","6. Les lignes d’exemple (matrice PARAMETRES et QUESTIONNAIRE) peuvent être remplacées ou supprimées : elles servent uniquement de modèle, à l’image du questionnaire EPC/SENEVAL."],wrap)
    default_labels={"5":"Totalement d’accord","4":"D’accord","3":"Neutre","2":"Pas d’accord","1":"Totalement en désaccord"}
    smin,smax=int(template["scale"]["min"]),int(template["scale"]["max"])
    ps=wb.add_worksheet("PARAMETRES"); ps.write_row(0,0,["Nom du questionnaire (à remplacer par le vôtre)",template["name"]],head); ps.write_row(1,0,["Description",template["description"]],wrap); ps.write_row(3,0,["Note","Libellé (exemple EPC/SENEVAL, à adapter)"],head); labels=template["scale"].get("labels",{}); [ps.write_row(4+(smax-n),0,[n,labels.get(str(n)) or default_labels.get(str(n),'')]) for n in range(smax,smin-1,-1)]; ps.set_column(0,0,38); ps.set_column(1,1,55)
    ws=wb.add_worksheet("QUESTIONNAIRE"); ws.write_row(0,0,["Domaine","Référence","Indicateur qualitatif ou Capacité"],head); ws.freeze_panes(1,0); row=1
    if not template["domains"]: ws.write_row(row,0,["EXEMPLE — Gestion des Ressources Humaines","EXEMPLE — Formation au personnel","EXEMPLE — Nous offrons régulièrement la formation au personnel"],wrap); row+=1
    for d in template["domains"]:
        for i in d["indicators"]:
            ws.write_row(row,0,[d["label"],i["label"],i["description"]],wrap); row+=1
    ws.set_column(0,0,34); ws.set_column(1,1,38); ws.set_column(2,2,75); wb.close(); return out.getvalue()


def blank_matrix_xlsx():
    return matrix_xlsx({"name":"Exemple à remplacer : Diagnostic EPC / [nom de l’atelier]", "description":"", "version":1, "scale":{"type":"numeric","min":1,"max":5,"labels":{}}, "priority":{"maxPerDomain":3}, "domains":[]})


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
        level_nums=sorted(int(k) for k in labels if str(k).lstrip('-').isdigit())
        smin,smax=(level_nums[0],level_nums[-1]) if level_nums else (1,5)
        grouped={}; order=[]
        for r in q[1:]:
            r=(r+["", "", ""])[:3]
            if not any(r) or r[0].startswith("EXEMPLE"): continue
            if not all(r): errors.append("Chaque ligne doit contenir Domaine, Référence et Indicateur qualitatif ou Capacité"); continue
            if r[0] not in grouped: grouped[r[0]]=[]; order.append(r[0])
            grouped[r[0]].append({"code":"","label":r[1],"description":r[2],"response_type":"numeric","required":True,"active":True,"display_order":len(grouped[r[0]])})
        if not values.get("Nom du questionnaire","").strip(): errors.append("Nom du questionnaire obligatoire")
        return {"errors":errors,"template":{"name":values.get("Nom du questionnaire",""),"description":values.get("Description",""),"scale":{"type":"numeric","min":smin,"max":smax,"labels":labels},"priority":{"maxPerDomain":3},"domains":[{"label":d,"display_order":n+1,"indicators":grouped[d]} for n,d in enumerate(order)]},"rows":sum(len(v) for v in grouped.values())}
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


def save_import(db, data, owner_user_id=None):
    if data["errors"]: raise ValueError("La matrice comporte des erreurs")
    tid=create_blank_template(db,data["template"],owner_user_id); t=data["template"]; db.execute("UPDATE templates SET scale_json=?,priority_json=? WHERE id=?",(json.dumps(t["scale"]),json.dumps(t["priority"]),tid))
    for d in t["domains"]:
        did=str(uuid.uuid4()); code="domain-"+uuid.uuid4().hex[:8]; db.execute("INSERT INTO domains VALUES (?,?,?,?,?,?,?)",(did,tid,code,d["label"],"",d["display_order"],1))
        for i in d["indicators"]: db.execute("INSERT INTO indicators VALUES (?,?,?,?,?,?,?,?,?,?)",(str(uuid.uuid4()),did,i["code"],i["label"],i["description"],i["response_type"],int(i["required"]),i["display_order"],int(i["active"]),"{}"))
    db.commit(); return tid


def update_template(db, template_id, data):
    """Editing a template already used by a session forks it first (clone_template),
    exactly like the historical PUT route did: the old version stays untouched."""
    used = db.execute("SELECT 1 FROM sessions WHERE template_id=? LIMIT 1", (template_id,)).fetchone()
    version_created = False
    if used:
        template_id = clone_template(db, template_id)
        version_created = True
    old = template_payload(db, template_id)
    scale = data.get("scale", old["scale"])
    db.execute("UPDATE templates SET name=?,description=?,scale_json=?,priority_json=?,updated_at=? WHERE id=?",
        (data.get("name", old["name"]), data.get("description", old["description"]), json.dumps(scale), json.dumps(data.get("priority", old["priority"])), now(), template_id))
    db.commit()
    return template_id, version_created


def delete_template(db, template_id, force=False):
    """Returns one of "protected" (EPC/SENEVAL, never deletable), "in_use" (blocked
    unless force=True, in which case it's archived instead of deleted), or "deleted"."""
    protected = db.execute("SELECT name FROM templates WHERE id=?", (template_id,)).fetchone()
    if protected and protected["name"] == "EPC / SENEVAL":
        return "protected"
    if db.execute("SELECT 1 FROM sessions WHERE template_id=? LIMIT 1", (template_id,)).fetchone():
        if force:
            db.execute("UPDATE templates SET status='archived',updated_at=? WHERE id=?", (now(), template_id))
            db.commit()
            return "archived"
        return "in_use"
    db.execute("DELETE FROM indicators WHERE domain_id IN (SELECT id FROM domains WHERE template_id=?)", (template_id,))
    db.execute("DELETE FROM domains WHERE template_id=?", (template_id,))
    db.execute("DELETE FROM templates WHERE id=?", (template_id,))
    db.commit()
    return "deleted"


def create_domain(db, template_id, data):
    did = str(uuid.uuid4())
    code = data.get("code") or "domain-" + uuid.uuid4().hex[:8]
    db.execute("INSERT INTO domains VALUES (?,?,?,?,?,?,?)",
        (did, template_id, code, data["label"], data.get("description", ""), int(data.get("displayOrder") or next_order(db, "domains", "template_id", template_id)), int(data.get("active", True))))
    db.commit()
    return did


def update_domain(db, domain_id, data):
    db.execute("UPDATE domains SET label=?,description=?,display_order=?,active=? WHERE id=?",
        (data["label"], data.get("description", ""), int(data.get("displayOrder", 1)), int(data.get("active", True)), domain_id))
    db.commit()


def delete_domain(db, domain_id):
    """Refuses (returns False, affected_sessions) if any response in any session
    references an indicator of this domain; affected_sessions is then the read-only
    list shown to the pilote (id/name), same as the historical DELETE route."""
    affected = rows(db, "SELECT DISTINCT s.id,s.name FROM sessions s JOIN responses r ON r.session_id=s.id JOIN indicators i ON i.id=r.indicator_id WHERE i.domain_id=?", (domain_id,))
    if affected:
        return False, affected
    db.execute("DELETE FROM indicators WHERE domain_id=?", (domain_id,))
    db.execute("DELETE FROM domains WHERE id=?", (domain_id,))
    db.commit()
    return True, affected


def create_indicator(db, domain_id, data):
    iid = str(uuid.uuid4())
    code = data.get("code") or "indicator-" + uuid.uuid4().hex[:8]
    db.execute("INSERT INTO indicators VALUES (?,?,?,?,?,?,?,?,?,?)",
        (iid, domain_id, code, data["label"], data.get("description", ""), data.get("responseType", "numeric"), int(data.get("required", True)), int(data.get("displayOrder") or next_order(db, "indicators", "domain_id", domain_id)), int(data.get("active", True)), json.dumps(data.get("configuration", {}))))
    db.commit()
    return iid


def update_indicator(db, indicator_id, data):
    db.execute("UPDATE indicators SET domain_id=?,code=?,label=?,description=?,response_type=?,required=?,display_order=?,active=?,configuration_json=? WHERE id=?",
        (data["domainId"], data["code"], data["label"], data.get("description", ""), data.get("responseType", "numeric"), int(data.get("required", True)), int(data.get("displayOrder", 1)), int(data.get("active", True)), json.dumps(data.get("configuration", {})), indicator_id))
    db.commit()


def delete_indicator(db, indicator_id):
    """Refuses (returns False, used_count) if responses already reference this
    indicator, same as the historical DELETE route."""
    used = db.execute("SELECT COUNT(*) FROM responses WHERE indicator_id=?", (indicator_id,)).fetchone()[0]
    if used:
        return False, used
    db.execute("DELETE FROM indicators WHERE id=?", (indicator_id,))
    db.commit()
    return True, used
