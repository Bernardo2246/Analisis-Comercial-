"""Extractor de Go High Level → ghl/raw/AAAA-MM-DD/*.json (gitignored).

Uso:
    python3 ghl/extract.py --check                 # 1 llamada: valida token, versión y location
    python3 ghl/extract.py                         # extracción completa
    python3 ghl/extract.py --since 2026-07-01      # fecha inicial para mensajes y citas
    python3 ghl/extract.py --only opportunities,tasks

Cada recurso se guarda en su propio archivo; _manifest.json registra conteos,
errores y número de llamadas. Un recurso que falla no detiene a los demás.
Los mensajes se guardan SIN cuerpo (solo metadatos: tipo, dirección, fecha, usuario).
"""
import argparse
import datetime as dt
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from client import GHLClient, GHLError  # noqa: E402

RAW = Path(__file__).resolve().parent / "raw"
TZ_OFFSET = dt.timezone(dt.timedelta(hours=-6))  # America/Mexico_City (sin horario de verano)
MSG_KEEP = ("id", "type", "messageType", "direction", "status", "dateAdded",
            "contactId", "conversationId", "userId", "source")


def _ms(d):
    return int(dt.datetime.combine(d, dt.time(), TZ_OFFSET).timestamp() * 1000)


# -- recursos ---------------------------------------------------------------

def location(c, a):
    return [c.get(f"/locations/{c.location_id}").get("location", {})]


def users(c, a):
    return c.get("/users/", {"locationId": c.location_id}).get("users", [])


def pipelines(c, a):
    return c.get("/opportunities/pipelines", {"locationId": c.location_id}).get("pipelines", [])


def tags(c, a):
    return c.get(f"/locations/{c.location_id}/tags").get("tags", [])


def custom_fields(c, a):
    return c.get(f"/locations/{c.location_id}/customFields").get("customFields", [])


def opportunities(c, a):
    out, seen = [], set()
    params = {c.loc_param: c.location_id, "limit": 100, "status": "all"}
    while True:
        r = c.get("/opportunities/search", params)
        batch = r.get("opportunities", [])
        new = [o for o in batch if o.get("id") not in seen]
        if not new:
            break
        out += new
        seen.update(o["id"] for o in new)
        meta = r.get("meta", {})
        if not meta.get("startAfterId") or (meta.get("total") and len(out) >= meta["total"]):
            break
        params["startAfter"], params["startAfterId"] = meta.get("startAfter"), meta["startAfterId"]
    return out


def contacts(c, a):
    out, page = [], 1
    while True:
        r = c.search("/contacts/search", {"locationId": c.location_id, "pageLimit": 100, "page": page})
        batch = r.get("contacts", [])
        out += batch
        if not batch or len(out) >= r.get("total", 0):
            break
        page += 1
    return out


def tasks(c, a):
    out, skip = [], 0
    while True:
        r = c.search(f"/locations/{c.location_id}/tasks/search", {"limit": 100, "skip": skip})
        batch = r.get("tasks", [])
        out += batch
        if len(batch) < 100:
            break
        skip += 100
    return out


def notes(c, a, opps=None):
    """Notas solo de contactos con oportunidad (evita miles de llamadas)."""
    ids = sorted({o.get("contactId") or (o.get("contact") or {}).get("id") for o in (opps or [])} - {None})
    out = []
    for cid in ids:
        for n in c.get(f"/contacts/{cid}/notes").get("notes", []):
            n["contactId"] = cid
            out.append(n)
    return out


def calendars(c, a):
    return c.get("/calendars/", {"locationId": c.location_id}).get("calendars", [])


def appointments(c, a, cals=None):
    out = []
    start, end = a.since, dt.date.today() + dt.timedelta(days=1)
    for cal in cals or []:
        d = start
        while d < end:  # ventanas de 31 días
            e = min(d + dt.timedelta(days=31), end)
            r = c.get("/calendars/events", {"locationId": c.location_id, "calendarId": cal["id"],
                                            "startTime": _ms(d), "endTime": _ms(e)})
            out += r.get("events", [])
            d = e
    return out


def conversations(c, a):
    out, params = [], {"locationId": c.location_id, "limit": 100}
    while True:
        r = c.get("/conversations/search", params)
        batch = r.get("conversations", [])
        out += batch
        if len(batch) < 100:
            break
        last = batch[-1]
        params["startAfterDate"] = last.get("sort", [last.get("lastMessageDate")])[0]
    return out


def messages(c, a, convs=None):
    """Mensajes desde --since, leídos conversación por conversación.

    Nota (2026-10-20): el export masivo /conversations/messages/export solo devolvió WhatsApp y
    llamadas de octubre aunque hay conversaciones desde julio; la lectura por conversación sí trae
    el historial completo. Por eso es la vía principal.
    """
    def slim(m):
        s = {k: m.get(k) for k in MSG_KEEP}
        meta = m.get("meta") or {}
        call = meta.get("call") or meta
        s["callDuration"], s["callStatus"] = call.get("duration") or meta.get("callDuration"), call.get("status") or meta.get("callStatus")
        return s

    out, since = [], a.since.isoformat()
    for conv in convs or []:
        if (conv.get("lastMessageDate") or 0) < _ms(a.since) or conv.get("lastMessageType") == "TYPE_NO_SHOW":
            continue
        params = {"limit": 100}
        while True:
            r = c.get(f"/conversations/{conv['id']}/messages", params, version="2021-04-15").get("messages", {})
            batch = r.get("messages", [])
            out += [slim(m) for m in batch if (m.get("dateAdded") or "") >= since]
            if not r.get("nextPage") or not batch or min(m.get("dateAdded") or "" for m in batch) < since:
                break
            params["lastMessageId"] = r.get("lastMessageId")
    return out


ORDER = ["location", "users", "pipelines", "tags", "custom_fields", "opportunities",
         "contacts", "tasks", "notes", "calendars", "appointments", "conversations", "messages"]
DEPENDS = {"notes": "opportunities", "appointments": "calendars", "messages": "conversations"}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--check", action="store_true", help="solo valida credenciales (1 llamada)")
    ap.add_argument("--since", type=dt.date.fromisoformat, default=dt.date(2026, 7, 1))
    ap.add_argument("--only", help="lista separada por comas de recursos")
    a = ap.parse_args()

    c = GHLClient()
    if a.check:
        try:
            loc = location(c, a)[0]
            print(f"OK · versión {c.version} · location '{loc.get('name')}' · tz {loc.get('timezone')}")
        except GHLError as e:
            sys.exit(f"FALLÓ: {e}")
        return

    wanted = a.only.split(",") if a.only else ORDER
    outdir = RAW / dt.date.today().isoformat()
    outdir.mkdir(parents=True, exist_ok=True)
    manifest = {"fecha": dt.datetime.now().isoformat(timespec="seconds"), "since": a.since.isoformat(),
                "version": c.version, "recursos": {}}
    data = {}
    for name in ORDER:
        dep = DEPENDS.get(name)
        if name not in wanted:
            continue
        if dep and dep not in data and (outdir / f"{dep}.json").exists():
            data[dep] = json.loads((outdir / f"{dep}.json").read_text())
        print(f"→ {name} …", flush=True)
        try:
            fn = globals()[name]
            rows = fn(c, a, data.get(dep)) if dep else fn(c, a)
            data[name] = rows
            (outdir / f"{name}.json").write_text(json.dumps(rows, ensure_ascii=False, indent=1))
            manifest["recursos"][name] = {"registros": len(rows)}
            print(f"   {len(rows)} registros")
        except GHLError as e:
            manifest["recursos"][name] = {"error": str(e)}
            print(f"   ERROR {e}", file=sys.stderr)
    manifest["llamadas"] = c.calls
    (outdir / "_manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=1))
    print(f"\nListo: {outdir} · {c.calls} llamadas")


if __name__ == "__main__":
    main()
