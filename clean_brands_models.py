"""
Limpiador de Marca / Modelo de equipos electromédicos.

Replica la lógica del Excel "Machida_con_propuesta_EUDAMED_MDCLEAN.xlsx" pero
de forma automática y a gran escala:

  1. Normaliza la marca (mayúsculas, sin acentos, sin puntuación, alias).
  2. Normaliza el modelo (mayúsculas, sin acentos, espacios y guiones colapsados).
  3. Agrupa por marca y, dentro de cada marca, agrupa los modelos por una clave
     alfanumérica (clave de comparación) para detectar variantes de formato.
  4. Dentro de cada grupo elige un modelo canónico (el que más usuarios tienen,
     y a igualdad el más "formateado" con guiones).
  5. Para los modelos que no caen en ningún grupo se hace fuzzy matching contra
     el resto de modelos de la misma marca y se propone el más cercano si la
     similitud es >= 92.
  6. Genera una "Nota propuesta", una "CONFIANZA" 0-100 y un STATUS.
  7. Vuelca el resultado a un Excel con las mismas columnas que el ejemplo de
     Machida.

Uso:
    python clean_brands_models.py \
        --input  data/Marcas_y_modelos_en_uso__EQ_ELECTROMEDICO.xlsx \
        --output output/Marcas_y_modelos_LIMPIO.xlsx
"""
from __future__ import annotations

import argparse
import re
import unicodedata
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path
from typing import Optional

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from rapidfuzz import fuzz, process

from eudamed_client import EudamedClient


# ---------------------------------------------------------------------------
# 1. Normalización básica
# ---------------------------------------------------------------------------

def strip_accents(text: str) -> str:
    if not text:
        return ""
    nfkd = unicodedata.normalize("NFKD", text)
    return "".join(c for c in nfkd if not unicodedata.combining(c))


def normalize_brand(raw) -> str:
    """Marca normalizada: mayúsculas, sin acentos, espacios colapsados."""
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return ""
    s = strip_accents(str(raw)).upper().strip()
    s = re.sub(r"[\s_]+", " ", s)
    s = re.sub(r"[.,;:'\"`´¨]", "", s)
    return s.strip()


def normalize_model(raw) -> str:
    """Modelo normalizado: mayúsculas sin acentos, espacios y guiones colapsados."""
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return ""
    s = strip_accents(str(raw)).upper().strip()
    s = s.replace("Ⅰ", "I").replace("Ⅱ", "II").replace("Ⅲ", "III").replace("Ⅳ", "IV")
    s = re.sub(r"\s+", " ", s)
    s = re.sub(r"\s*-\s*", "-", s)         # "ENT - 30" -> "ENT-30"
    s = re.sub(r"-{2,}", "-", s)
    return s.strip(" -")


def model_key(raw) -> str:
    """Clave de comparación: solo alfanuméricos, mayúsculas, normaliza romanos."""
    s = normalize_model(raw)
    # Romanos sueltos por arábigos para juntar 'ENT-30PII' / 'ENT-30P2'
    # OJO: solo en sufijo numérico-romano (T2 <-> TII)
    s = re.sub(r"([A-Z])2\b", r"\1II", s)
    s = re.sub(r"([A-Z])3\b", r"\1III", s)
    s = re.sub(r"([A-Z])4\b", r"\1IV", s)
    s = re.sub(r"[^A-Z0-9]", "", s)
    return s


# ---------------------------------------------------------------------------
# 2. Alias de marca
# ---------------------------------------------------------------------------
#
# Mapa explícito: cuando una marca normalizada coincide con la izquierda, se
# reemplaza por la de la derecha. Permite unificar "STORZ" con "KARL STORZ",
# "GE HEALTHCARE" con "GENERAL ELECTRIC", etc.

BRAND_ALIASES: dict[str, str] = {
    "STORZ": "KARL STORZ",
    "KARLSTORZ": "KARL STORZ",
    "ZEISS": "CARL ZEISS",
    "CARLZEISS": "CARL ZEISS",
    "GE": "GENERAL ELECTRIC",
    "GE HEALTHCARE": "GENERAL ELECTRIC",
    "GE HEALTH CARE": "GENERAL ELECTRIC",
    "GENERAL ELECTRIC HEALTHCARE": "GENERAL ELECTRIC",
    "HP": "HEWLETT PACKARD ENTERPRISE (HPE)",
    "HEWLETT PACKARD": "HEWLETT PACKARD ENTERPRISE (HPE)",
    "HEWLETT-PACKARD": "HEWLETT PACKARD ENTERPRISE (HPE)",
    "HPE": "HEWLETT PACKARD ENTERPRISE (HPE)",
    "DRAGER": "DRÄGER",
    "DRAEGER": "DRÄGER",
    "TOSHIBA": "TOSHIBA MEDICAL SYSTEM",
    "TOSHIBA MEDICAL": "TOSHIBA MEDICAL SYSTEM",
    "TOSHIBA MEDICAL SYSTEMS": "TOSHIBA MEDICAL SYSTEM",
    "PHILLIPS": "PHILIPS",
    "PHYLIPS": "PHILIPS",
    "OLIMPUS": "OLYMPUS",
    "DATEX OHMEDA": "DATEX-OHMEDA",
    "DATEX": "DATEX-OHMEDA",
    "OHMEDA": "DATEX-OHMEDA",
    "WELCH-ALLYN": "WELCH ALLYN",
    "WELCHALLYN": "WELCH ALLYN",
    "FUJI FILM": "FUJIFILM",
    "FUJI": "FUJIFILM",
    "RICHARD-WOLF": "RICHARD WOLF",
    "B BRAUN": "B. BRAUN",
    "BBRAUN": "B. BRAUN",
    "BRAUN": "B. BRAUN",
    "ENRAF-NONIUS": "ENRAF NONIUS",
    "JP-SELECTA": "JP SELECTA",
    "SELECTA": "JP SELECTA",
    "NIHON-KOHDEN": "NIHON KOHDEN",
    "NIHONKOHDEN": "NIHON KOHDEN",
    "SIN MARCA": "SIN MARCA-NO DEFINIDA-DESCONOCIDA",
    "NO DEFINIDA": "SIN MARCA-NO DEFINIDA-DESCONOCIDA",
    "DESCONOCIDA": "SIN MARCA-NO DEFINIDA-DESCONOCIDA",
    "NO DEFINIDO": "SIN MARCA-NO DEFINIDA-DESCONOCIDA",
    "DESCONOCIDO": "SIN MARCA-NO DEFINIDA-DESCONOCIDA",
    "S/M": "SIN MARCA-NO DEFINIDA-DESCONOCIDA",
    "N/D": "SIN MARCA-NO DEFINIDA-DESCONOCIDA",
}


def canonical_brand(raw: str) -> str:
    norm = normalize_brand(raw)
    if not norm:
        return "SIN MARCA-NO DEFINIDA-DESCONOCIDA"
    # match directo
    if norm in BRAND_ALIASES:
        return BRAND_ALIASES[norm]
    # quitar espacios para coincidir con KARLSTORZ, CARLZEISS, etc.
    nospaces = norm.replace(" ", "")
    if nospaces in BRAND_ALIASES:
        return BRAND_ALIASES[nospaces]
    return norm


# ---------------------------------------------------------------------------
# 3. Detección de modelos "no válidos" (basura)
# ---------------------------------------------------------------------------

INVALID_MODEL_PATTERNS = [
    re.compile(r"^[\W_]*$"),              # sólo símbolos: '-', '--', '*', '.'
    re.compile(r"^[¿?]+$"),
    re.compile(r"^N\.?A\.?$", re.I),
    re.compile(r"^NO DEFINIDO$", re.I),
    re.compile(r"^DESCONOCIDO$", re.I),
    re.compile(r"^SIN MODELO$", re.I),
    re.compile(r"^S/M$", re.I),
    re.compile(r"^X+$", re.I),
]


def is_invalid_model(model) -> bool:
    if model is None:
        return True
    s = str(model).strip()
    if not s or s.lower() == "nan":
        return True
    return any(p.match(s) for p in INVALID_MODEL_PATTERNS)


# ---------------------------------------------------------------------------
# 4. Selección de modelo canónico dentro de un grupo
# ---------------------------------------------------------------------------

def pick_canonical(models: list[str]) -> str:
    """
    Dado un grupo de variantes del MISMO modelo (misma `model_key`), elige el
    representante canónico:
      - el que aparece más veces;
      - a igualdad, el que tiene guiones (más "formateado");
      - a igualdad, el más corto.
    """
    if not models:
        return ""
    counts = Counter(models)

    def score(m: str) -> tuple:
        c = counts[m]
        has_dash = "-" in m
        no_spaces = " " not in m
        return (c, has_dash, no_spaces, -len(m), m)

    return max(counts.keys(), key=score)


# ---------------------------------------------------------------------------
# 5. Generación de la nota propuesta
# ---------------------------------------------------------------------------

def diff_only_punctuation(a: str, b: str) -> bool:
    """True si las diferencias entre a y b son sólo espacios/guiones/puntos."""
    clean_a = re.sub(r"[\s\-\.]+", "", a)
    clean_b = re.sub(r"[\s\-\.]+", "", b)
    return clean_a == clean_b


def make_note(original_model: str, proposal: str, match_type: str,
              fuzzy_score: float | None = None) -> str:
    if proposal is None or proposal == "":
        return "No encontrado (sin propuesta automática)"

    norm = normalize_model(original_model)
    if norm == proposal:
        return "Coincide"
    if diff_only_punctuation(norm, proposal):
        return "Normalización de formato"
    if match_type == "FUZZY":
        return f"Posible error tipográfico (similitud {fuzzy_score:.0f}%)"
    if match_type == "GROUP":
        return "Normalización por patrón"
    return "Normalización"


# ---------------------------------------------------------------------------
# 6. Pipeline principal
# ---------------------------------------------------------------------------

OUTPUT_COLUMNS = [
    "Marca", "Modelo",
    "Propuesta modelo oficial / EUDAMED", "Nota propuesta",
    "Encontrado en EUDAMED (según extracto)",
    "Nombre comercial EUDAMED (si aplica)",
    "UDI-DI (completo)", "Basic UDI-DI (si aplica)",
    "Fuente (EUDAMED / web)",
    "Verificado en internet (fabricante/distribuidor)",
    "Link de referencia",
    "ESTADO", "FECHA_REVISION", "SUGERENCIA_WEB", "URL_WEB",
    "MARCA_LIMPIA", "MODELO_LIMPIO", "CATEGORIA",
    "MATCH_TIPO", "CONFIANZA",
    "PROPUESTA_MARCA", "PROPUESTA_MODELO",
    "TRADE_NAME", "BASIC_UDI_DI", "UDI_DI", "CLASE_RIESGO",
    "STATUS", "FUENTE", "OBSERVACIONES", "WEB_SCORE",
]


def enrich_with_eudamed(out: pd.DataFrame,
                        client: EudamedClient,
                        only_missing: bool = False,
                        max_lookups: Optional[int] = None,
                        only_brands: Optional[set[str]] = None) -> int:
    """Enriquece `out` in-place con datos de EUDAMED.

    Sólo consulta combinaciones únicas (MARCA_LIMPIA, PROPUESTA_MODELO) y
    propaga el resultado a todas las filas que coincidan.

    Devuelve el número de combinaciones consultadas.
    """
    mask_target = out["PROPUESTA_MODELO"].notna() & (out["PROPUESTA_MODELO"] != "")
    if only_missing:
        mask_target &= out["UDI_DI"].isna() | (out["UDI_DI"] == "")
    if only_brands:
        mask_target &= out["MARCA_LIMPIA"].isin(only_brands)

    pairs = (out.loc[mask_target, ["MARCA_LIMPIA", "PROPUESTA_MODELO"]]
                .dropna().drop_duplicates())
    if max_lookups is not None:
        pairs = pairs.head(max_lookups)

    print(f"  Consultas únicas a EUDAMED: {len(pairs)}")

    # Cache por (marca, modelo) -> dict aplanado
    results: dict[tuple[str, str], Optional[dict]] = {}
    for i, (_, row) in enumerate(pairs.iterrows(), 1):
        brand = row["MARCA_LIMPIA"]
        model = row["PROPUESTA_MODELO"]
        results[(brand, model)] = client.lookup(brand, model)
        if i % 50 == 0:
            print(f"    {i}/{len(pairs)} ({brand} {model})")
    client.flush()

    # Propagación a las filas
    for (brand, model), hit in results.items():
        if not hit:
            continue
        mask = (out["MARCA_LIMPIA"] == brand) & (out["PROPUESTA_MODELO"] == model)
        out.loc[mask, "TRADE_NAME"] = hit.get("trade_name") or ""
        out.loc[mask, "UDI_DI"] = hit.get("udi_di") or ""
        out.loc[mask, "BASIC_UDI_DI"] = hit.get("basic_udi_di") or ""
        out.loc[mask, "CLASE_RIESGO"] = hit.get("risk_class") or ""
        out.loc[mask, "FUENTE"] = "EUDAMED"
        out.loc[mask, "Encontrado en EUDAMED (según extracto)"] = (
            "Sí" if hit.get("brand_matched") else "Parcial"
        )
        out.loc[mask, "Nombre comercial EUDAMED (si aplica)"] = hit.get("trade_name") or ""
        out.loc[mask, "UDI-DI (completo)"] = hit.get("udi_di") or ""
        out.loc[mask, "Basic UDI-DI (si aplica)"] = hit.get("basic_udi_di") or ""
        out.loc[mask, "Fuente (EUDAMED / web)"] = "EUDAMED"
        out.loc[mask, "Verificado en internet (fabricante/distribuidor)"] = "Sí (EUDAMED)"
        out.loc[mask, "Link de referencia"] = "EUDAMED"
    return len(pairs)


def process_dataframe(df: pd.DataFrame, fuzzy_threshold: int = 92) -> pd.DataFrame:
    """
    Espera un DataFrame con columnas:
        Equipo_codiMarca, refMarca_nombre, Equipo_modelo
    Devuelve un DataFrame con las columnas de OUTPUT_COLUMNS.
    """
    df = df.copy()
    df["MARCA_LIMPIA"] = df["refMarca_nombre"].map(canonical_brand)
    df["MODELO_LIMPIO"] = df["Equipo_modelo"].map(normalize_model)
    df["MODEL_KEY"] = df["Equipo_modelo"].map(model_key)

    # Construye, por marca, el mapa key -> modelo canónico
    print(f"  Marcas distintas tras normalizar: {df['MARCA_LIMPIA'].nunique()}")
    canonical_by_brand: dict[str, dict[str, str]] = defaultdict(dict)
    fuzzy_index_by_brand: dict[str, list[str]] = defaultdict(list)
    key_count_by_brand: dict[str, dict[str, int]] = defaultdict(dict)

    for brand, sub in df.groupby("MARCA_LIMPIA"):
        # Modelos "válidos" no vacíos
        valid = sub[~sub["MODELO_LIMPIO"].map(is_invalid_model)]
        # Por cada model_key elige canónico (entre las versiones MODELO_LIMPIO)
        for key, group in valid.groupby("MODEL_KEY"):
            if not key:
                continue
            canonical = pick_canonical(group["MODELO_LIMPIO"].tolist())
            canonical_by_brand[brand][key] = canonical
            key_count_by_brand[brand][key] = len(group)
        fuzzy_index_by_brand[brand] = sorted(
            set(canonical_by_brand[brand].values())
        )

    proposals: list[str | None] = []
    notes: list[str] = []
    match_types: list[str] = []
    confidences: list[int] = []
    statuses: list[str] = []

    for _, row in df.iterrows():
        brand = row["MARCA_LIMPIA"]
        modelo_norm = row["MODELO_LIMPIO"]
        key = row["MODEL_KEY"]

        # 1) Modelo basura/no válido
        if is_invalid_model(row["Equipo_modelo"]) or not key:
            proposals.append(None)
            notes.append("No encontrado (sin propuesta automática)")
            match_types.append("NONE")
            confidences.append(0)
            statuses.append("REVISAR")
            continue

        # 2) Coincidencia por clave alfanumérica
        canonical = canonical_by_brand.get(brand, {}).get(key)
        if canonical:
            key_count = key_count_by_brand[brand].get(key, 0)
            # 2a) Si la clave es "singleton" (sólo aparece 1 vez), comprobar si
            #     existe otro canónico mucho más popular en la misma marca al
            #     que se parezca mucho — pista de typo.
            typo_hit = None
            if key_count == 1:
                candidates = [
                    c for k, c in canonical_by_brand[brand].items()
                    if k != key and key_count_by_brand[brand].get(k, 0) >= 3
                ]
                if candidates:
                    # Usamos fuzz.ratio (Levenshtein normalizado): penaliza
                    # caracteres añadidos, así una variante "MOD-R" no se
                    # colapsa con "MOD".
                    best = process.extractOne(
                        modelo_norm, candidates, scorer=fuzz.ratio,
                    )
                    if best and best[1] >= 93:
                        typo_hit = best
            if typo_hit:
                proposals.append(typo_hit[0])
                notes.append(
                    f"Posible error tipográfico (similitud {typo_hit[1]:.0f}%)"
                )
                match_types.append("FUZZY")
                confidences.append(int(typo_hit[1]))
                statuses.append("REVISAR")
            elif canonical == modelo_norm:
                proposals.append(canonical)
                notes.append("Coincide")
                match_types.append("EXACT")
                confidences.append(100)
                statuses.append("OK")
            else:
                proposals.append(canonical)
                if diff_only_punctuation(modelo_norm, canonical):
                    notes.append("Normalización de formato")
                else:
                    notes.append("Normalización por patrón")
                match_types.append("GROUP")
                confidences.append(95)
                statuses.append("OK")
            continue

        # 3) Fuzzy matching contra el catálogo de la marca
        candidates = fuzzy_index_by_brand.get(brand, [])
        best = process.extractOne(modelo_norm, candidates,
                                  scorer=fuzz.ratio) if candidates else None
        if best and best[1] >= fuzzy_threshold:
            proposals.append(best[0])
            notes.append(f"Posible error tipográfico (similitud {best[1]:.0f}%)")
            match_types.append("FUZZY")
            confidences.append(int(best[1]))
            statuses.append("REVISAR")
        else:
            proposals.append(None)
            notes.append("No encontrado (sin propuesta automática)")
            match_types.append("NONE")
            confidences.append(0)
            statuses.append("REVISAR")

    fecha = datetime.now().replace(microsecond=0)

    out = pd.DataFrame({
        "Marca": df["refMarca_nombre"],
        "Modelo": df["Equipo_modelo"],
        "Propuesta modelo oficial / EUDAMED": proposals,
        "Nota propuesta": notes,
        "Encontrado en EUDAMED (según extracto)": "",
        "Nombre comercial EUDAMED (si aplica)": "",
        "UDI-DI (completo)": "",
        "Basic UDI-DI (si aplica)": "",
        "Fuente (EUDAMED / web)": "",
        "Verificado en internet (fabricante/distribuidor)": "",
        "Link de referencia": "",
        "ESTADO": statuses,
        "FECHA_REVISION": fecha,
        "SUGERENCIA_WEB": "",
        "URL_WEB": "",
        "MARCA_LIMPIA": df["MARCA_LIMPIA"],
        "MODELO_LIMPIO": df["MODELO_LIMPIO"],
        "CATEGORIA": "GENERAL",
        "MATCH_TIPO": match_types,
        "CONFIANZA": confidences,
        "PROPUESTA_MARCA": df["MARCA_LIMPIA"],
        "PROPUESTA_MODELO": proposals,
        "TRADE_NAME": "",
        "BASIC_UDI_DI": "",
        "UDI_DI": "",
        "CLASE_RIESGO": "",
        "STATUS": statuses,
        "FUENTE": "AUTO",
        "OBSERVACIONES": notes,
        "WEB_SCORE": "",
    }, columns=OUTPUT_COLUMNS)
    return out


# ---------------------------------------------------------------------------
# 7. Escritura del Excel con formato y resumen
# ---------------------------------------------------------------------------

def write_excel(out: pd.DataFrame, output_path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "REVISION"

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="305496")

    ws.append(OUTPUT_COLUMNS)
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill

    fill_ok = PatternFill("solid", fgColor="E2EFDA")
    fill_revisar = PatternFill("solid", fgColor="FFF2CC")
    fill_no = PatternFill("solid", fgColor="FCE4D6")

    # bulk write
    for record in out.itertuples(index=False, name=None):
        ws.append(record)

    estado_col = OUTPUT_COLUMNS.index("ESTADO") + 1
    for row_idx in range(2, ws.max_row + 1):
        estado = ws.cell(row=row_idx, column=estado_col).value
        fill = fill_ok if estado == "OK" else fill_revisar
        if ws.cell(row=row_idx,
                   column=OUTPUT_COLUMNS.index("Propuesta modelo oficial / EUDAMED") + 1
                   ).value in (None, ""):
            fill = fill_no
        for c in (estado_col, OUTPUT_COLUMNS.index("CONFIANZA") + 1):
            ws.cell(row=row_idx, column=c).fill = fill

    # Anchuras razonables
    widths = {
        "A": 28, "B": 32, "C": 28, "D": 38, "L": 10, "M": 20,
        "P": 28, "Q": 28, "R": 14, "S": 12, "T": 11,
        "U": 28, "V": 28, "AB": 12, "AC": 14, "AD": 38,
    }
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    # ---- Hoja resumen ----
    summary = wb.create_sheet("RESUMEN")
    total = len(out)
    n_ok = (out["ESTADO"] == "OK").sum()
    n_rev = (out["ESTADO"] == "REVISAR").sum()
    n_no_prop = out["Propuesta modelo oficial / EUDAMED"].isna().sum()
    n_marcas = out["MARCA_LIMPIA"].nunique()
    summary.append(["Métrica", "Valor"])
    summary.append(["Filas totales", total])
    summary.append(["Marcas distintas (limpias)", n_marcas])
    summary.append(["Propuesta automática OK", int(n_ok)])
    summary.append(["A revisar", int(n_rev)])
    summary.append(["Sin propuesta", int(n_no_prop)])
    for c in summary[1]:
        c.font = header_font
        c.fill = header_fill
    summary.column_dimensions["A"].width = 38
    summary.column_dimensions["B"].width = 14

    # ---- Hoja por marca con conteos ----
    counts = (out.groupby("MARCA_LIMPIA")
                  .agg(filas=("Marca", "size"),
                       sin_propuesta=("Propuesta modelo oficial / EUDAMED",
                                      lambda s: s.isna().sum()),
                       a_revisar=("ESTADO", lambda s: (s == "REVISAR").sum()))
                  .sort_values("filas", ascending=False)
                  .reset_index())
    sh = wb.create_sheet("POR_MARCA")
    sh.append(["MARCA_LIMPIA", "FILAS", "SIN_PROPUESTA", "A_REVISAR"])
    for r in counts.itertuples(index=False, name=None):
        sh.append(list(r))
    for c in sh[1]:
        c.font = header_font
        c.fill = header_fill
    for col, w in {"A": 40, "B": 10, "C": 16, "D": 12}.items():
        sh.column_dimensions[col].width = w

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)


# ---------------------------------------------------------------------------
# 8. CLI
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--input", "-i", required=True, type=Path,
                        help="Excel de entrada (REVISAR sheet con columnas "
                             "Equipo_codiMarca, refMarca_nombre, Equipo_modelo)")
    parser.add_argument("--output", "-o", required=True, type=Path,
                        help="Ruta del Excel de salida")
    parser.add_argument("--sheet", default="REVISAR",
                        help="Nombre de la hoja de entrada (default: REVISAR)")
    parser.add_argument("--fuzzy-threshold", type=int, default=92,
                        help="Umbral 0-100 de similitud para proponer "
                             "corrección por fuzzy matching (default: 92)")
    parser.add_argument("--eudamed", action="store_true",
                        help="Enriquece el resultado consultando EUDAMED. "
                             "Requiere acceso a ec.europa.eu y la librería "
                             "`requests` instalada.")
    parser.add_argument("--eudamed-cache", type=Path,
                        default=Path("output/.eudamed_cache.json"),
                        help="Fichero JSON de cache para las respuestas de "
                             "EUDAMED (default: output/.eudamed_cache.json)")
    parser.add_argument("--eudamed-only-missing", action="store_true",
                        help="Sólo consulta EUDAMED para las filas que aún "
                             "no tienen UDI_DI rellenado.")
    parser.add_argument("--eudamed-max-lookups", type=int, default=None,
                        help="Límite máximo de consultas únicas a EUDAMED "
                             "(útil para pruebas; default: sin límite).")
    parser.add_argument("--eudamed-brands", nargs="+", default=None,
                        help="Restringe el lookup EUDAMED a las marcas "
                             "indicadas (MARCA_LIMPIA exacta).")
    parser.add_argument("--eudamed-delay", type=float, default=0.4,
                        help="Pausa mínima entre peticiones a EUDAMED, en "
                             "segundos (default: 0.4).")
    args = parser.parse_args()

    print(f"Leyendo {args.input} (hoja '{args.sheet}')...")
    df = pd.read_excel(args.input, sheet_name=args.sheet)
    expected = {"Equipo_codiMarca", "refMarca_nombre", "Equipo_modelo"}
    missing = expected - set(df.columns)
    if missing:
        raise SystemExit(f"Faltan columnas en el Excel: {missing}")
    print(f"  Filas leídas: {len(df)}")

    print("Procesando...")
    out = process_dataframe(df, fuzzy_threshold=args.fuzzy_threshold)

    if args.eudamed:
        print("Consultando EUDAMED...")
        client = EudamedClient(
            cache_path=args.eudamed_cache,
            request_delay=args.eudamed_delay,
        )
        only_brands = set(args.eudamed_brands) if args.eudamed_brands else None
        n = enrich_with_eudamed(
            out, client,
            only_missing=args.eudamed_only_missing,
            max_lookups=args.eudamed_max_lookups,
            only_brands=only_brands,
        )
        print(f"  EUDAMED: {n} consultas únicas realizadas")

    print(f"Escribiendo {args.output}...")
    write_excel(out, args.output)

    n_total = len(out)
    n_ok = (out["ESTADO"] == "OK").sum()
    n_rev = (out["ESTADO"] == "REVISAR").sum()
    n_no = out["Propuesta modelo oficial / EUDAMED"].isna().sum()
    print(f"Hecho. Total {n_total} | OK {n_ok} | REVISAR {n_rev} | Sin propuesta {n_no}")


if __name__ == "__main__":
    main()
