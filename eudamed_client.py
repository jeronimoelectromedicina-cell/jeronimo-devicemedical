"""
Cliente para la API pública de EUDAMED.

EUDAMED (European Database on Medical Devices) expone un buscador público en
https://ec.europa.eu/tools/eudamed/. El SPA consume una API REST también
pública (sin autenticación) cuyos endpoints estables a fecha 2025 son:

    GET https://ec.europa.eu/tools/eudamed/api/devices/udiDiData
        ?page=0&pageSize=25&size=25
        &languageIso2Code=en&iso2Code=en
        &tradeName=<texto>            # opcional
        &referenceNumber=<texto>      # opcional
        &deviceStatusCode=refdata.device-model-status.on-the-market

La respuesta es un JSON con la forma:

    {
      "content": [
        {
          "tradeName": "...",
          "primaryDi": {"code": "0405...", "issuingEntity": {...}},
          "basicUdi": "...",
          "riskClass": {"code": "refdata.risk-class.iia", "srnCode": "IIa"},
          "manufacturer": {"name": "..."},
          ...
        },
        ...
      ],
      "totalElements": N
    }

La estructura concreta puede cambiar; por eso este cliente:
  * deja la URL base como parámetro,
  * acepta un `transport` callable inyectable (para testear sin red), y
  * extrae los campos de forma tolerante (cualquier campo ausente queda en None).

Uso típico (en local, con red):

    from eudamed_client import EudamedClient
    cli = EudamedClient(cache_path="output/.eudamed_cache.json",
                        request_delay=0.4)
    hit = cli.lookup("MACHIDA", "ENT-30PIII")
    if hit:
        print(hit["trade_name"], hit["udi_di"], hit["risk_class"])
"""
from __future__ import annotations

import json
import logging
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Optional


DEFAULT_BASE_URL = "https://ec.europa.eu/tools/eudamed/api/devices/udiDiData"
DEFAULT_USER_AGENT = (
    "jeronimo-devicemedical-cleaner/1.0 "
    "(+https://github.com/jeronimoelectromedicina-cell/jeronimo-devicemedical)"
)

log = logging.getLogger("eudamed")


# ---------------------------------------------------------------------------
# Tipos y helpers
# ---------------------------------------------------------------------------

Transport = Callable[[str, dict], dict]
"""Callable que recibe (url, params) y devuelve el JSON parseado.

Permite inyectar un mock en los tests sin tocar la red.
"""


def _norm(s: Optional[str]) -> str:
    return (s or "").strip().upper()


def _flatten(rec: dict[str, Any]) -> dict[str, Any]:
    """Extrae los campos que nos interesan de un registro EUDAMED."""
    primary_di = rec.get("primaryDi") or {}
    risk = rec.get("riskClass") or {}
    manuf = rec.get("manufacturer") or {}
    return {
        "trade_name": rec.get("tradeName"),
        "udi_di": primary_di.get("code"),
        "basic_udi_di": rec.get("basicUdi") or rec.get("basicUdiDi"),
        "risk_class": risk.get("srnCode") or risk.get("code"),
        "manufacturer": manuf.get("name"),
        "status": (rec.get("deviceStatus") or {}).get("code"),
    }


# ---------------------------------------------------------------------------
# Transport por defecto (requests)
# ---------------------------------------------------------------------------

def _build_requests_transport(timeout: float, user_agent: str) -> Transport:
    """Crea un Transport basado en `requests`. Import perezoso.

    Reintenta hasta 4 veces con backoff exponencial ante errores de red o 5xx.
    """
    try:
        import requests
    except ImportError as e:  # pragma: no cover
        raise SystemExit(
            "Necesitas instalar 'requests' para usar EudamedClient con red: "
            "pip install requests"
        ) from e

    session = requests.Session()
    session.headers.update({
        "User-Agent": user_agent,
        "Accept": "application/json, text/plain, */*",
        "Accept-Language": "en;q=0.9,es;q=0.8",
    })

    def transport(url: str, params: dict) -> dict:
        delay = 1.0
        last_err: Optional[Exception] = None
        for attempt in range(4):
            try:
                r = session.get(url, params=params, timeout=timeout)
                if r.status_code == 200:
                    return r.json()
                if r.status_code in (429, 500, 502, 503, 504):
                    raise RuntimeError(f"HTTP {r.status_code}")
                # 4xx no recuperables: devolvemos respuesta vacía
                log.warning("EUDAMED %s -> HTTP %s", url, r.status_code)
                return {}
            except Exception as e:
                last_err = e
                if attempt == 3:
                    break
                time.sleep(delay)
                delay *= 2
        log.warning("EUDAMED: agotados reintentos (%s)", last_err)
        return {}

    return transport


# ---------------------------------------------------------------------------
# Cliente con cache
# ---------------------------------------------------------------------------

@dataclass
class EudamedClient:
    cache_path: Optional[Path] = None
    base_url: str = DEFAULT_BASE_URL
    request_delay: float = 0.4
    timeout: float = 20.0
    user_agent: str = DEFAULT_USER_AGENT
    transport: Optional[Transport] = None

    _cache: dict[str, Any] = None  # type: ignore[assignment]
    _last_request: float = 0.0

    def __post_init__(self) -> None:
        self._cache = {}
        if self.cache_path:
            self.cache_path = Path(self.cache_path)
            if self.cache_path.exists():
                try:
                    self._cache = json.loads(self.cache_path.read_text("utf-8"))
                    log.info("Cache EUDAMED cargada: %d entradas", len(self._cache))
                except Exception as e:
                    log.warning("Cache EUDAMED corrupta, se reinicia: %s", e)
                    self._cache = {}
        if self.transport is None:
            self.transport = _build_requests_transport(self.timeout, self.user_agent)

    # ---- API pública ----

    def lookup(self, brand: str, model: str) -> Optional[dict]:
        """Devuelve el mejor match (dict aplanado) o None si no se encuentra.

        Cachea por la pareja (brand, model) normalizada.
        """
        if not brand or not model:
            return None
        key = f"{_norm(brand)}|{_norm(model)}"
        if key in self._cache:
            return self._cache[key]

        result = self._search(brand, model)
        self._cache[key] = result
        self._maybe_persist()
        return result

    def flush(self) -> None:
        """Guarda la cache a disco (idempotente)."""
        if not self.cache_path:
            return
        self.cache_path.parent.mkdir(parents=True, exist_ok=True)
        self.cache_path.write_text(
            json.dumps(self._cache, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    # ---- Internals ----

    def _maybe_persist(self) -> None:
        # Guarda cada 25 entradas nuevas para no perder progreso.
        if self.cache_path and len(self._cache) % 25 == 0:
            self.flush()

    def _throttle(self) -> None:
        if self.request_delay <= 0:
            return
        elapsed = time.monotonic() - self._last_request
        if elapsed < self.request_delay:
            time.sleep(self.request_delay - elapsed)
        self._last_request = time.monotonic()

    def _query(self, **params) -> list[dict]:
        self._throttle()
        defaults = {
            "page": 0,
            "pageSize": 25,
            "size": 25,
            "languageIso2Code": "en",
            "iso2Code": "en",
        }
        defaults.update({k: v for k, v in params.items() if v is not None})
        data = self.transport(self.base_url, defaults) or {}  # type: ignore[misc]
        content = data.get("content") or data.get("results") or []
        return list(content)

    def _search(self, brand: str, model: str) -> Optional[dict]:
        """Busca por `tradeName=model` y filtra por marca; si no, por `referenceNumber=model`."""
        brand_norm = _norm(brand)
        model_norm = _norm(model)

        candidates: list[dict] = []
        for params in (
            {"tradeName": model},
            {"referenceNumber": model},
        ):
            candidates.extend(self._query(**params))
            if candidates:
                break

        if not candidates:
            return None

        # Filtra por marca: el nombre del fabricante debe contener la marca
        # normalizada (subset). Se acepta también igualdad insensible.
        same_brand = [
            r for r in candidates
            if brand_norm in _norm((r.get("manufacturer") or {}).get("name"))
        ]
        # Si ningún registro encaja por marca, devolvemos el primer candidato
        # como "mejor esfuerzo" pero etiquetado.
        chosen = same_brand[0] if same_brand else candidates[0]
        flat = _flatten(chosen)
        flat["brand_matched"] = bool(same_brand)
        flat["model_query"] = model_norm
        return flat


__all__ = ["EudamedClient", "Transport", "DEFAULT_BASE_URL", "DEFAULT_USER_AGENT"]
