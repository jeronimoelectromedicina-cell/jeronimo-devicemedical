"""
Test end-to-end de la integración EUDAMED sin tocar red.

Usa un Transport falso que responde con datos predefinidos para validar:
  * la fixture de cliente lee/escribe cache JSON,
  * `enrich_with_eudamed` propaga las columnas a todas las filas que
    comparten (MARCA_LIMPIA, PROPUESTA_MODELO),
  * el filtro `only_brands` funciona,
  * la cache evita peticiones repetidas.
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from clean_brands_models import OUTPUT_COLUMNS, enrich_with_eudamed  # noqa: E402
from eudamed_client import EudamedClient  # noqa: E402


# ---------------------------------------------------------------------------
# Fake transport
# ---------------------------------------------------------------------------

CATALOG = {
    ("MACHIDA", "ENT-30PIII"): {
        "tradeName": "Flexible Nasopharyngo Laryngoscope ENT-30PIII",
        "primaryDi": {"code": "04580137730255"},
        "basicUdi": "458013773A0013E",
        "riskClass": {"srnCode": "IIa", "code": "refdata.risk-class.iia"},
        "manufacturer": {"name": "Machida Endoscope Co., Ltd."},
    },
    ("MACHIDA", "ENT-30PC"): {
        "tradeName": "Portable Flexible Nasopharyngo Laryngoscope ENT-30PC",
        "primaryDi": {"code": "04580137732372"},
        "basicUdi": "458013773A0013E",
        "riskClass": {"srnCode": "IIa", "code": "refdata.risk-class.iia"},
        "manufacturer": {"name": "Machida Endoscope Co., Ltd."},
    },
}


def make_transport():
    """Devuelve un transport + contador de llamadas para verificar caché."""
    calls: list[tuple[str, dict]] = []

    def transport(url: str, params: dict) -> dict:
        calls.append((url, dict(params)))
        # Buscamos por tradeName o referenceNumber. Hacemos un match laxo:
        # el catálogo contiene el modelo si el query aparece en la clave.
        q = (params.get("tradeName") or params.get("referenceNumber") or "").upper()
        results = [
            rec for (_, model), rec in CATALOG.items() if model.upper() == q
        ]
        return {"content": results, "totalElements": len(results)}

    return transport, calls


# ---------------------------------------------------------------------------
# Helpers para crear un DataFrame "estilo `process_dataframe` ya ejecutado"
# ---------------------------------------------------------------------------

def fake_processed_df() -> pd.DataFrame:
    rows = [
        ("MACHIDA", "ENT-30PIII"),
        ("MACHIDA", "ENT-30PIII"),  # duplicado para verificar propagación
        ("MACHIDA", "ENT-30PC"),
        ("PHILIPS", "TP300"),       # no está en el catálogo
    ]
    df = pd.DataFrame({
        "MARCA_LIMPIA": [b for b, _ in rows],
        "PROPUESTA_MODELO": [m for _, m in rows],
    })
    # Rellenamos el resto de columnas de salida con vacíos.
    for col in OUTPUT_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    return df[OUTPUT_COLUMNS]


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------

def test_enrich_propagates_to_duplicate_rows(tmp_path: Path) -> None:
    transport, calls = make_transport()
    client = EudamedClient(
        cache_path=tmp_path / "cache.json",
        request_delay=0,
        transport=transport,
    )
    out = fake_processed_df()

    n = enrich_with_eudamed(out, client)

    # 3 combos únicos (PHILIPS TP300 también se consulta, aunque devuelva vacío)
    assert n == 3

    # ENT-30PIII aparece 2 veces y ambas deben tener UDI rellenado
    udi_pii = out.loc[out["PROPUESTA_MODELO"] == "ENT-30PIII", "UDI_DI"].tolist()
    assert udi_pii == ["04580137730255", "04580137730255"]
    assert (out.loc[out["PROPUESTA_MODELO"] == "ENT-30PIII",
                    "BASIC_UDI_DI"] == "458013773A0013E").all()
    assert (out.loc[out["PROPUESTA_MODELO"] == "ENT-30PIII",
                    "CLASE_RIESGO"] == "IIa").all()
    assert (out.loc[out["PROPUESTA_MODELO"] == "ENT-30PIII",
                    "FUENTE"] == "EUDAMED").all()

    # PHILIPS TP300 no se encontró: UDI_DI debe quedar vacío
    assert (out.loc[out["PROPUESTA_MODELO"] == "TP300", "UDI_DI"] == "").all()


def test_cache_avoids_repeat_calls(tmp_path: Path) -> None:
    transport, calls = make_transport()
    cache = tmp_path / "cache.json"
    client = EudamedClient(cache_path=cache, request_delay=0, transport=transport)
    out = fake_processed_df()

    enrich_with_eudamed(out, client)
    first_call_count = len(calls)
    assert first_call_count > 0

    # Segundo cliente, misma cache en disco
    client2 = EudamedClient(cache_path=cache, request_delay=0, transport=transport)
    enrich_with_eudamed(fake_processed_df(), client2)
    # No debería haber llamadas nuevas
    assert len(calls) == first_call_count

    # La cache JSON existe y tiene contenido
    cached = json.loads(cache.read_text("utf-8"))
    assert any("ENT-30PIII" in k for k in cached.keys())


def test_only_brands_filter(tmp_path: Path) -> None:
    transport, calls = make_transport()
    client = EudamedClient(
        cache_path=tmp_path / "cache.json",
        request_delay=0,
        transport=transport,
    )
    out = fake_processed_df()

    enrich_with_eudamed(out, client, only_brands={"PHILIPS"})

    # Sólo se debió consultar TP300 (el cliente puede hacer 1-2 peticiones
    # internas por combo —tradeName + referenceNumber si la 1ª no encuentra—
    # pero todas tienen que ser sobre TP300, no sobre los modelos MACHIDA).
    assert calls, "se esperaba al menos una petición"
    queried_models = {
        (c[1].get("tradeName") or c[1].get("referenceNumber")) for c in calls
    }
    assert queried_models == {"TP300"}

    # MACHIDA no debe tener UDI rellenado
    assert (out.loc[out["MARCA_LIMPIA"] == "MACHIDA", "UDI_DI"] == "").all()
