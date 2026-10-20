# ghl/ — Conexión de solo lectura con Go High Level

Módulo reutilizable para extraer el pipeline y las actividades comerciales de
Integra Cimentaciones desde Go High Level (API v2, `services.leadconnectorhq.com`).
Sirve para el reporte trimestral y como base de un futuro agente comercial.
Solo usa la librería estándar de Python (3.9+); no hay nada que instalar.

## Seguridad
- Credenciales **solo** en `ghl/.env` (gitignored). Nunca en código, commits ni HTML.
- `ghl/raw/` (gitignored) guarda la extracción cruda con datos personales. No se publica.
- Al sitio público solo llega `snapshots/pipeline_AAAA-MM.json`: conteos y montos agregados,
  sin nombres de contactos.
- El token debe tener **únicamente scopes `.readonly`**. El código no llama endpoints de escritura;
  los dos `POST` que usa son búsquedas.

## Configuración (una sola vez)
1. En GHL: **Settings → Private Integrations → Create new integration** con estos scopes:
   `locations.readonly`, `locations/tags.readonly`, `locations/customFields.readonly`,
   `locations/tasks.readonly`, `opportunities.readonly`, `contacts.readonly`,
   `calendars.readonly`, `calendars/events.readonly`, `conversations.readonly`,
   `conversations/message.readonly`, `users.readonly`.
2. `cp ghl/.env.example ghl/.env` y llenar `GHL_TOKEN` y `GHL_LOCATION_ID`
   (Settings → Business Profile).
3. `python3 ghl/extract.py --check` → debe responder `OK · versión … · location '…'`.
   Si falla por versión, poner `GHL_API_VERSION=v3` en `.env` y repetir.

## Cada trimestre
```bash
python3 ghl/extract.py --since 2026-07-01          # baja todo a ghl/raw/AAAA-MM-DD/
python3 ghl/summarize.py                           # imprime conteos para validar
python3 ghl/summarize.py --write --corte 2026-10-01 --ventas <N> --monto <MXN>
```
`--ventas/--monto` son el acumulado del Excel de ventas a la fecha de corte (mismo criterio
que los snapshots anteriores). La slide "Evolución del pipeline" compara el snapshot nuevo
con el anterior.

## Endpoints
| Recurso (`extract.py`) | Método y ruta | Scope |
|---|---|---|
| location | `GET /locations/{locationId}` | locations.readonly |
| users | `GET /users/?locationId=` | users.readonly |
| pipelines | `GET /opportunities/pipelines?locationId=` | opportunities.readonly |
| tags | `GET /locations/{locationId}/tags` | locations/tags.readonly |
| custom_fields | `GET /locations/{locationId}/customFields` | locations/customFields.readonly |
| opportunities | `GET /opportunities/search` (paginado `startAfter`/`startAfterId`, `status=all`) | opportunities.readonly |
| contacts | `POST /contacts/search` (búsqueda, paginado `page`) | contacts.readonly |
| tasks | `POST /locations/{locationId}/tasks/search` (búsqueda, `limit`/`skip`) | locations/tasks.readonly |
| notes | `GET /contacts/{contactId}/notes` (solo contactos con oportunidad) | contacts.readonly |
| calendars | `GET /calendars/?locationId=` | calendars.readonly |
| appointments | `GET /calendars/events?calendarId=&startTime=&endTime=` (ventanas de 31 días) | calendars/events.readonly |
| conversations | `GET /conversations/search?locationId=` | conversations.readonly |
| messages | `GET /conversations/messages/export` (`Version: v3`, cursor); respaldo `GET /conversations/{id}/messages` | conversations/message.readonly |

Los mensajes se guardan **sin cuerpo**: solo tipo, dirección, fecha, usuario y duración de llamada.

Headers: `Authorization: Bearer <token>`, `Version: 2021-07-28` (o `v3`), `Accept: application/json`.
Límite de GHL: ~100 llamadas / 10 s por location. `client.py` espera 150 ms entre llamadas y reintenta
con backoff en 429/5xx.

## Salidas
- `ghl/raw/AAAA-MM-DD/<recurso>.json` + `_manifest.json` (conteos, errores, número de llamadas).
  Un recurso que falla queda registrado en el manifest y no detiene a los demás.
- `snapshots/pipeline_AAAA-MM.json`: llaves de `pipeline_2026-07.json` + `por_etapa`,
  `perdidas_por_motivo`, `fuentes_opps`, `tags_contactos_top`, `estancadas` (sin cambio de etapa en
  más de 30 días), `actividades_por_mes`, `actividades_por_usuario`.

## Para un agente
- Las funciones de `extract.py` (`opportunities(c, a)`, `tasks(c, a)`, …) son independientes y
  reciben un `GHLClient`. Se pueden importar y llamar por separado.
- `summarize.summarize(Path(raw_dir), corte_date)` devuelve el dict agregado sin escribir nada.
- Contexto de negocio: migración de Clientify a GHL. Inició en jul-2026, uso parcial en ago y
  completo en sep. Las recompras y los recomendados llegan directo y **no** pasan por el CRM: que
  haya menos ventas en GHL que en el Excel es lo esperado.
