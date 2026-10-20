"""Resumen agregado (sin datos personales) de una extracción GHL.

Uso:
    python3 ghl/summarize.py                      # usa la extracción más reciente, solo imprime
    python3 ghl/summarize.py --write --corte 2026-10-01 \
        --ventas 234 --monto 15855823.17          # escribe snapshots/pipeline_2026-10.json

--ventas/--monto vienen del Excel de ventas (acumulado jun-2025 → corte) y se
copian tal cual, igual que en snapshots anteriores. Sin ellos quedan en null.

Esquema: mismas llaves que snapshots/pipeline_2026-07.json + bloques nuevos
(por_etapa, perdidas_por_motivo, fuentes, estancadas, actividades_por_mes,
actividades_por_usuario). Solo conteos y montos: nunca nombres de contactos.
"""
import argparse
import collections as C
import datetime as dt
import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
RAW = HERE / "raw"
SNAP = HERE.parent / "snapshots"
TZ = dt.timezone(dt.timedelta(hours=-6))
STALE_DAYS = 30

MSG_TYPE = {  # messageType de GHL → categoría del reporte
    "TYPE_CALL": "llamadas", "CALL": "llamadas",
    "TYPE_WHATSAPP": "whatsapp", "WHATSAPP": "whatsapp",
    "TYPE_EMAIL": "email", "EMAIL": "email",
    "TYPE_SMS": "sms", "SMS": "sms",
    "TYPE_FACEBOOK": "facebook", "TYPE_INSTAGRAM": "instagram",
}


def load(d, name):
    p = d / f"{name}.json"
    return json.loads(p.read_text()) if p.exists() else []


def parse_ts(v):
    if v in (None, ""):
        return None
    if isinstance(v, (int, float)):
        return dt.datetime.fromtimestamp(v / 1000, TZ)
    try:
        t = dt.datetime.fromisoformat(str(v).replace("Z", "+00:00"))
        return (t if t.tzinfo else t.replace(tzinfo=TZ)).astimezone(TZ)
    except ValueError:
        return None


def month(v):
    t = parse_ts(v)
    return t.strftime("%Y-%m") if t else None


def summarize(d, corte):
    pipes = load(d, "pipelines")
    opps = load(d, "opportunities")
    contacts = load(d, "contacts")
    tasks = load(d, "tasks")
    notes = load(d, "notes")
    appts = load(d, "appointments")
    msgs = load(d, "messages")
    users = {u.get("id"): u.get("name") or f"{u.get('firstName', '')} {u.get('lastName', '')}".strip()
             for u in load(d, "users")}

    stage_name, stage_pos, pipe_name = {}, {}, {}
    for p in pipes:
        for s in p.get("stages", []):
            stage_name[s["id"]] = s.get("name")
            stage_pos[s["id"]] = s.get("position", 0)
            pipe_name[s["id"]] = p.get("name")

    val = lambda o: float(o.get("monetaryValue") or 0)
    st = C.Counter(o.get("status") for o in opps)
    open_ = [o for o in opps if o.get("status") == "open"]

    etapas = C.defaultdict(lambda: {"opps": 0, "valor": 0.0})
    for o in open_:
        k = (pipe_name.get(o.get("pipelineStageId"), "?"), stage_name.get(o.get("pipelineStageId"), "?"),
             stage_pos.get(o.get("pipelineStageId"), 99))
        etapas[k]["opps"] += 1
        etapas[k]["valor"] += val(o)

    motivos = C.Counter((o.get("lostReason") or o.get("lostReasonId") or "sin motivo")
                        for o in opps if o.get("status") == "lost")

    cutoff = dt.datetime.combine(corte, dt.time(), TZ) - dt.timedelta(days=STALE_DAYS)
    stale = [o for o in open_
             if (parse_ts(o.get("lastStageChangeAt") or o.get("updatedAt") or o.get("createdAt")) or cutoff) < cutoff]

    act = C.defaultdict(C.Counter)
    por_usuario = C.defaultdict(C.Counter)
    for m in msgs:
        cat = MSG_TYPE.get(m.get("messageType")) or MSG_TYPE.get(str(m.get("type")))
        if not cat:
            continue
        mo = month(m.get("dateAdded"))
        act[mo][cat] += 1
        if m.get("direction") == "outbound":
            act[mo][cat + "_salientes"] += 1
            por_usuario[users.get(m.get("userId"), "sistema/automatización")][cat] += 1
    for t in tasks:
        act[month(t.get("dateAdded") or t.get("createdAt"))]["tareas_creadas"] += 1
        if t.get("completed"):
            act[month(t.get("completedAt") or t.get("updatedAt") or t.get("dueDate"))]["tareas_completadas"] += 1
    for n in notes:
        act[month(n.get("dateAdded"))]["notas"] += 1
        por_usuario[users.get(n.get("userId"), "sin usuario")]["notas"] += 1
    for ap in appts:
        act[month(ap.get("startTime"))]["citas"] += 1
    for o in opps:
        act[month(o.get("createdAt"))]["opps_creadas"] += 1
        if o.get("status") == "won":
            act[month(o.get("lastStatusChangeAt") or o.get("updatedAt"))]["opps_ganadas"] += 1
    act.pop(None, None)

    fuente_opp = C.Counter((o.get("source") or "sin fuente").strip().upper() for o in opps)
    tag_cont = C.Counter(t.strip().lower() for ct in contacts for t in (ct.get("tags") or []))

    total = len(opps)
    return {
        "corte": corte.isoformat(),
        "fuente": f"API Go High Level (extracción {d.name}) + Excel de ventas",
        "ventas_cerradas": None,
        "monto_vendido": None,
        "ticket_prom": None,
        "crm_opps_total": total,
        "crm_ganadas": st.get("won", 0),
        "crm_perdidas": st.get("lost", 0),
        "crm_abandonadas": st.get("abandoned", 0),
        "opps_abiertas": len(open_),
        "pct_cierre_crm": round(st.get("won", 0) / total, 4) if total else None,
        "pipeline_abierto_bruto": round(sum(val(o) for o in open_), 2),
        "nota_pipeline": f"suma de monetaryValue de {len(open_)} opps open en GHL",
        "por_etapa": [{"pipeline": k[0], "etapa": k[1], "opps": v["opps"], "valor": round(v["valor"], 2)}
                      for k, v in sorted(etapas.items(), key=lambda kv: (kv[0][0], kv[0][2]))],
        "perdidas_por_motivo": dict(motivos.most_common()),
        "fuentes_opps": dict(fuente_opp.most_common()),
        "tags_contactos_top": dict(tag_cont.most_common(25)),
        "contactos_total": len(contacts),
        "estancadas": {"dias": STALE_DAYS, "opps": len(stale), "valor": round(sum(val(o) for o in stale), 2)},
        "actividades_por_mes": {k: dict(v) for k, v in sorted(act.items())},
        "actividades_por_usuario": {k: dict(v) for k, v in por_usuario.items()},
        "nota_actividades": "GHL desde jul-2026 (migración: jul inicio, ago parcial, sep completa). "
                            "Recompras y recomendados entran fuera del CRM.",
    }


def report(s):
    mx = lambda v: f"${v:,.0f}"
    print(f"Opps: {s['crm_opps_total']} · ganadas {s['crm_ganadas']} · perdidas {s['crm_perdidas']} · "
          f"abandonadas {s['crm_abandonadas']} · abiertas {s['opps_abiertas']}")
    print(f"Pipeline abierto bruto: {mx(s['pipeline_abierto_bruto'])}   "
          f"(estancadas >{s['estancadas']['dias']}d: {s['estancadas']['opps']} / {mx(s['estancadas']['valor'])})")
    print("\nAbiertas por etapa:")
    for e in s["por_etapa"]:
        print(f"  {e['pipeline'][:22]:22} {e['etapa'][:30]:30} {e['opps']:>4}  {mx(e['valor']):>14}")
    print("\nPérdidas por motivo:", s["perdidas_por_motivo"])
    print("\nActividades por mes:")
    for m, v in s["actividades_por_mes"].items():
        print(f"  {m}: " + " · ".join(f"{k} {n}" for k, n in sorted(v.items())))
    print("\nPor usuario:", s["actividades_por_usuario"])


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--raw", help="carpeta de extracción (default: la más reciente)")
    ap.add_argument("--corte", type=dt.date.fromisoformat, default=dt.date.today())
    ap.add_argument("--ventas", type=int)
    ap.add_argument("--monto", type=float)
    ap.add_argument("--write", action="store_true")
    a = ap.parse_args()

    dirs = sorted(p for p in RAW.glob("*") if p.is_dir())
    d = Path(a.raw) if a.raw else (dirs[-1] if dirs else None)
    if not d or not d.exists():
        sys.exit("No hay extracción en ghl/raw/. Corre primero ghl/extract.py")
    s = summarize(d, a.corte)
    if a.ventas and a.monto:
        s["ventas_cerradas"], s["monto_vendido"] = a.ventas, a.monto
        s["ticket_prom"] = round(a.monto / a.ventas, 2)
    report(s)
    if a.write:
        out = SNAP / f"pipeline_{a.corte.strftime('%Y-%m')}.json"
        out.write_text(json.dumps(s, ensure_ascii=False, indent=1))
        print(f"\nEscrito {out}")


if __name__ == "__main__":
    main()
