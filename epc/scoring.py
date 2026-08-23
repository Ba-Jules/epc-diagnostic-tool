"""Moteur de calcul EPC : graduation et analyse capacite/consensus.

Extrait de app.py (lot 3 de la modularisation, cf. AUDIT_MODULARISATION_8800.md) :
AUCUNE formule n'est modifiee ici, seule l'emplacement du code bouge - copie a
l'identique, verifiee par les golden tests decimaux de tests.py (bornes de
graduation, N=0/N=1/N>1, "manquants", ponderation entre domaines) qui passaient
deja contre le code d'origine avant ce deplacement.

analysis_for() reste le chemin de calcul unique : aucun autre code ne doit
recalculer un score EPC. analysis() est un simple adaptateur pour une session
seule ; c'est ce que l'audit appelle "encapsuler EPC sans modifier ses
formules" (Lot 3) - l'abstraction multi-strategies ("puis permettre d'autres
methodes") est delibbrement hors perimetre tant qu'un seul modele existe.
"""
from __future__ import annotations

import json

from .db import rows, template_payload
from .profile import participants_matching_dimension, resolve_dimension_field

# Cohorts smaller than this have every capacity/consensus number suppressed
# in dimension_analysis() (lot 5, cf. AUDIT_MODULARISATION_8800.md - flagged
# "risque critique pour confidentialite et biais petits N"). A tiny matching
# subgroup is often re-identifiable by the pilot reading the report, so the
# safe default is to hide the numbers outright, not merely flag them as
# low-confidence.
MIN_COHORT_N = 5


def grade(value, norm):
    """Classify value into a graduated band from norm (low, high, result) tuples.

    Matches the reference KOICA calculation tool: the continuous standardized
    score is rounded to the nearest whole number first, then looked up against
    the bands' own authored integer bounds (which are contiguous once rounded,
    e.g. ...59-63, 64-67, 68-71... — no gap-filling needed post-rounding).
    """
    if value is None:
        return None
    v = round(max(0.0, min(100.0, float(value))))
    for low, high, result in norm:
        if low <= v <= high:
            return result
    return None


def analysis(db, session_id: str):
    """Single-session analysis. Thin wrapper over analysis_for so a lone
    workshop and a multi-group consolidation share one calculation path."""
    session = db.execute("SELECT * FROM sessions WHERE id=?", (session_id,)).fetchone()
    if not session:
        return None
    result = analysis_for(db, [session_id])
    if result is None:
        return None
    result["session"] = dict(session)
    return result


def analysis_for(db, session_ids: list[str], participant_ids: set[str] | None = None):
    """Same EPC calculation as analysis(), pooling responses/participants over
    one or several session ids. A single id behaves exactly as before; several
    ids (same template) is what powers campaign consolidation — the maths are
    never a mean-of-means, they recompute directly from individual responses.

    Capacité/consensus are computed only from participants with status='completed'
    (a questionnaire opened but abandoned mid-way must never silently shift the
    published score) — participantCount below still counts every participant
    row (started or completed) so "commencés" stays visible separately from
    "validés"/completedCount.

    participant_ids (lot 5, optional) further restricts every query below to
    that id set — this is the one and only place a dimension filter actually
    touches the calculation: it recomputes from the same individual response
    rows, never from a pre-aggregated table, so a filtered cohort's numbers
    are exactly what analysis_for() would compute if only that cohort had
    ever answered. None (the default) means "no restriction", identical to
    the pre-lot-5 behaviour.
    """
    if not session_ids:
        return None
    first = db.execute("SELECT * FROM sessions WHERE id=?", (session_ids[0],)).fetchone()
    if not first:
        return None
    template = template_payload(db, first["template_id"])
    scale, rules, consensus, norm = template["scale"], template["scoring"], template["consensus"], template["grading"]
    low, high, amplitude = float(scale["min"]), float(scale["max"]), float(scale["max"] - scale["min"])
    output_max = float(rules.get("outputRange", [0, 100])[1])
    ph = ",".join("?" * len(session_ids))
    pid_clause, pid_params = "", ()
    if participant_ids is not None:
        pid_clause = f" AND p.id IN ({','.join('?' * len(participant_ids))})" if participant_ids else " AND 0"
        pid_params = tuple(participant_ids)
    all_values, output_domains, all_participant_ids = [], [], set()
    total_participants = len(rows(db, f"SELECT id FROM participants p WHERE p.session_id IN ({ph}){pid_clause}", (*session_ids, *pid_params)))
    for domain in template["domains"]:
        indicators = [i for i in domain["indicators"] if i["active"]]
        output_indicators, participant_means = [], {}
        for indicator in indicators:
            response_rows = rows(db, f"SELECT r.participant_id,r.value_json FROM responses r JOIN participants p ON p.id=r.participant_id WHERE r.session_id IN ({ph}) AND r.indicator_id=? AND p.status='completed'{pid_clause}", (*session_ids, indicator["id"], *pid_params))
            values = [float(json.loads(r["value_json"])) for r in response_rows if isinstance(json.loads(r["value_json"]), (int, float))]
            for r in response_rows:
                value = json.loads(r["value_json"])
                if isinstance(value, (int, float)):
                    participant_means.setdefault(r["participant_id"], []).append(float(value))
                    all_participant_ids.add(r["participant_id"])
            mean = sum(values) / len(values) if values else None
            sd = (sum((v - mean) ** 2 for v in values) / (len(values) - 1)) ** .5 if len(values) > 1 else None
            capacity = mean / high * output_max if mean is not None else None
            cons = max(0, output_max - consensus["factor"] * (sd / amplitude * output_max)) if sd is not None else None
            output_indicators.append({"id": indicator["id"], "code": indicator["code"], "label": indicator["label"], "responses": len(values), "missing": max(0, total_participants - len(values)), "mean": mean, "capacity": capacity, "dispersion": sd, "consensus": cons, "consensusNote": "single_respondent" if len(values) == 1 else None, "distribution": {str(k): values.count(k) for k in range(int(low), int(high) + 1)}})
            all_values.extend(values)
        person_scores = [sum(values) / len(values) for values in participant_means.values() if values]
        dmean = sum(person_scores) / len(person_scores) if person_scores else None
        dsd = (sum((v - dmean) ** 2 for v in person_scores) / (len(person_scores) - 1)) ** .5 if len(person_scores) > 1 else None
        dcap = dmean / high * output_max if dmean is not None else None
        dcons = max(0, output_max - consensus["factor"] * (dsd / amplitude * output_max)) if dsd is not None else None
        output_domains.append({"id": domain["id"], "code": domain["code"], "label": domain["label"], "responses": len(person_scores), "capacity": dcap, "dispersion": dsd, "consensus": dcons, "consensusNote": "single_respondent" if len(person_scores) == 1 else None, "gradedCapacity": grade(dcap, norm) if dcap is not None else None, "gradedConsensus": grade(dcons, norm) if dcons is not None else None, "indicators": output_indicators})
    # Mirrors the reference KOICA tool's "Moyenne" row: global capacity/consensus
    # (standardized and graduated alike) are unweighted averages across domains,
    # not a response-weighted pool of every individual answer. The graduated
    # global score in particular averages the domains' own graduated scores —
    # it is not grade() applied to the averaged standardized score.
    domain_caps=[d["capacity"] for d in output_domains if d["capacity"] is not None]; domain_cons=[d["consensus"] for d in output_domains if d["consensus"] is not None]
    global_capacity=sum(domain_caps)/len(domain_caps) if domain_caps else None
    gc=sum(domain_cons)/len(domain_cons) if domain_cons else None
    graded_caps=[d["gradedCapacity"] for d in output_domains if d["gradedCapacity"] is not None]; graded_cons=[d["gradedConsensus"] for d in output_domains if d["gradedConsensus"] is not None]
    global_graded_capacity=sum(graded_caps)/len(graded_caps) if graded_caps else None
    global_graded_consensus=sum(graded_cons)/len(graded_cons) if graded_cons else None
    completed=len(rows(db,f"SELECT id FROM participants p WHERE p.session_id IN ({ph}) AND p.status='completed'{pid_clause}",(*session_ids,*pid_params)))
    return {"sessionIds": session_ids, "participantCount": total_participants, "completedCount":completed, "domains": output_domains, "global": {"responses": len(all_values), "capacity": global_capacity, "consensus":gc, "consensusNote": "single_respondent" if len(all_participant_ids) == 1 else None, "gradedCapacity": global_graded_capacity, "gradedConsensus": global_graded_consensus}}


def _suppress_small_cohort(result: dict) -> None:
    """Nulls every capacity/consensus number in an analysis_for() result in
    place, keeping only counts and labels. Used by dimension_analysis() when
    a filtered cohort falls below MIN_COHORT_N."""
    for domain in result["domains"]:
        for indicator in domain["indicators"]:
            for key in ("mean", "capacity", "dispersion", "consensus"):
                indicator[key] = None
            indicator["distribution"] = {}
        for key in ("capacity", "dispersion", "consensus", "gradedCapacity", "gradedConsensus"):
            domain[key] = None
    for key in ("capacity", "consensus", "gradedCapacity", "gradedConsensus"):
        result["global"][key] = None


def dimension_analysis(db, session_id: str, field_key: str, value, min_n: int = MIN_COHORT_N):
    """Same EPC calculation as analysis(), restricted to the participants of
    `session_id` whose profile value for `field_key` matches `value` (lot 5:
    "tout champ categoriel autorise devient dimension analytique"). Raises
    ValueError (via resolve_dimension_field) if field_key isn't an active,
    pilot-flagged dimension of this session's attached profile — that check
    is the sole privacy gate, see its own docstring.

    Cohorts smaller than min_n come back with every capacity/consensus number
    suppressed (participant/completed counts stay visible, since those are
    already shown at the whole-session level and needed to explain why the
    numbers are hidden) — see MIN_COHORT_N.
    """
    field = resolve_dimension_field(db, session_id, field_key)
    matching_ids = participants_matching_dimension(db, session_id, field_key, value)
    result = analysis_for(db, [session_id], participant_ids=matching_ids)
    if result is None:
        return None
    suppressed = result["completedCount"] < min_n
    if suppressed:
        _suppress_small_cohort(result)
    result["dimension"] = {"fieldKey": field_key, "fieldLabel": field["label"], "value": value, "minRequired": min_n, "suppressed": suppressed}
    return result
