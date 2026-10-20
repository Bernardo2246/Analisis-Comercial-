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
python3 ghl/summarize.py --write --corte <AAAA-MM-DD> --ventas <N> --monto <MXN>
python3 data/build.py                              # integra ghl/raw/<fecha>/summary.json a la presentación y al PDF
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
| messages | `GET /conversations/{id}/messages` (`Version: 2021-04-15`, paginado `lastMessageId`) | conversations/message.readonly |

Los mensajes se guardan **sin cuerpo**: solo tipo, dirección, fecha, usuario y duración de llamada.

### Lecciones (primera extracción, 2026-10-20)
- **User-Agent obligatorio:** sin él, Cloudflare responde 403 (error 1010). `client.py` envía uno propio.
- **Mensajes por conversación:** el export masivo `/conversations/messages/export` solo devolvió WhatsApp/llamadas
  de octubre; la lectura por conversación trae el historial completo. Se omiten conversaciones `TYPE_NO_SHOW`.
- **Motivos de pérdida:** la API solo da `lostReasonId`; el catálogo id→nombre vive en `ghl/lost_reasons.json`
  (gitignored: se genera del export CSV de oportunidades). Si aparece un motivo nuevo, actualízalo desde un CSV nuevo.
- **Migración Clientify (jul-2026) parcial:** las opps con `source = Clientify` traen `createdAt` del día de la
  migración y estados no confiables. `summarize.py` las reporta aparte y las excluye de cohortes y tasas.

Headers: `Authorization: Bearer <token>`, `Version: 2021-07-28` (o `v3`), `Accept: application/json`.
Límite de GHL: ~100 llamadas / 10 s por location. `client.py` espera 150 ms entre llamadas y reintenta
con backoff en 429/5xx.

## Salidas
- `ghl/raw/AAAA-MM-DD/<recurso>.json` + `_manifest.json` (conteos, errores, número de llamadas).
  Un recurso que falla queda registrado en el manifest y no detiene a los demás.
- `ghl/raw/<fecha>/summary.json` (interno): todo el resumen — etapas, estancamiento, cohortes por mes,
  motivos de pérdida, migradas, llamadas por estado, salientes por usuario.
- `snapshots/pipeline_AAAA-MM.json` (público): solo agregados sin hallazgos internos. El de 2026-10 es la
  **línea base GHL**; los snapshots de Clientify (2026-05, 2026-07) no son comparables.

## Para un agente
- Las funciones de `extract.py` (`opportunities(c, a)`, `tasks(c, a)`, …) son independientes y
  reciben un `GHLClient`. Se pueden importar y llamar por separado.
- `summarize.summarize(Path(raw_dir), corte_date)` devuelve el dict agregado sin escribir nada.
- Contexto de negocio: migración de Clientify a GHL. Inició en jul-2026, uso parcial en ago y
  completo en sep. Las recompras y los recomendados llegan directo y **no** pasan por el CRM: que
  haya menos ventas en GHL que en el Excel es lo esperado.
