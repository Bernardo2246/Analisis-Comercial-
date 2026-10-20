"""Cliente mínimo (solo stdlib) para la API v2 de Go High Level (LeadConnector).

Solo lectura: el módulo nunca llama endpoints de escritura. Los POST que usa
(/contacts/search, /locations/{id}/tasks/search) son búsquedas.

Credenciales: ghl/.env (gitignored) con GHL_TOKEN y GHL_LOCATION_ID.
Opcional: GHL_API_VERSION (default 2021-07-28; la doc nueva también acepta v3).
"""
import json
import os
import time
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

BASE_URL = "https://services.leadconnectorhq.com"
HERE = Path(__file__).resolve().parent
DEFAULT_VERSION = "2021-07-28"

# GHL permite ~100 llamadas / 10 s por location; dejamos margen.
MIN_INTERVAL_S = 0.15
MAX_RETRIES = 5


class GHLError(Exception):
    def __init__(self, status, path, body):
        super().__init__(f"HTTP {status} en {path}: {body[:300]}")
        self.status = status
        self.path = path
        self.body = body


def load_env(path=HERE / ".env"):
    """Lee KEY=VALUE de ghl/.env sin sobrescribir variables ya definidas."""
    if not path.exists():
        return
    for line in path.read_text().splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        k, v = line.split("=", 1)
        os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))


class GHLClient:
    def __init__(self, token=None, location_id=None, version=None):
        load_env()
        self.token = token or os.environ.get("GHL_TOKEN")
        self.location_id = location_id or os.environ.get("GHL_LOCATION_ID")
        self.version = version or os.environ.get("GHL_API_VERSION") or DEFAULT_VERSION
        if not self.token or not self.location_id:
            raise SystemExit("Faltan GHL_TOKEN / GHL_LOCATION_ID en ghl/.env (ver .env.example)")
        self._last = 0.0
        self.calls = 0

    # -- transporte ---------------------------------------------------------
    def _request(self, method, path, params=None, body=None, version=None):
        url = BASE_URL + path
        if params:
            params = {k: v for k, v in params.items() if v is not None}
            url += "?" + urllib.parse.urlencode(params, doseq=True)
        data = json.dumps(body).encode() if body is not None else None
        headers = {
            "Authorization": f"Bearer {self.token}",
            "Version": version or self.version,
            "Accept": "application/json",
            # Cloudflare (error 1010) bloquea el User-Agent por defecto de urllib
            "User-Agent": "IntegraAnalisisComercial/1.0 (+ghl/client.py)",
        }
        if data is not None:
            headers["Content-Type"] = "application/json"

        for attempt in range(MAX_RETRIES):
            wait = MIN_INTERVAL_S - (time.monotonic() - self._last)
            if wait > 0:
                time.sleep(wait)
            self._last = time.monotonic()
            self.calls += 1
            req = urllib.request.Request(url, data=data, headers=headers, method=method)
            try:
                with urllib.request.urlopen(req, timeout=60) as resp:
                    raw = resp.read().decode()
                    return json.loads(raw) if raw else {}
            except urllib.error.HTTPError as e:
                text = e.read().decode(errors="replace")
                if e.code == 429 or e.code >= 500:
                    retry_after = e.headers.get("Retry-After")
                    time.sleep(float(retry_after) if retry_after else 2 ** attempt)
                    continue
                raise GHLError(e.code, path, text)
            except urllib.error.URLError:
                time.sleep(2 ** attempt)
        raise GHLError(0, path, "reintentos agotados")

    def get(self, path, params=None, **kw):
        return self._request("GET", path, params=params, **kw)

    def search(self, path, body, **kw):
        """POST de búsqueda (solo lectura)."""
        return self._request("POST", path, body=body, **kw)

    @property
    def loc_param(self):
        """Nombre del parámetro de location en /opportunities/search según versión."""
        return "location_id" if self.version == DEFAULT_VERSION else "locationId"
