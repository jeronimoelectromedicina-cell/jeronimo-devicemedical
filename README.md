# jeronimo-devicemedical

Limpiador automático de Marca / Modelo de equipos electromédicos. Replica la
lógica del Excel de ejemplo `data/Machida_con_propuesta_EUDAMED_MDCLEAN.xlsx`
pero a gran escala (probado con 60 409 filas / 5 035 marcas distintas en
~45 segundos).

## Qué hace

1. **Normaliza la marca**: mayúsculas, sin acentos, espacios colapsados, sin
   puntuación. Aplica un mapa de alias (`STORZ` ↔ `KARL STORZ`,
   `ZEISS` ↔ `CARL ZEISS`, `GE HEALTHCARE` ↔ `GENERAL ELECTRIC`,
   `DRAGER`/`DRAEGER` ↔ `DRÄGER`, etc.).
2. **Normaliza el modelo**: mayúsculas, sin acentos, espacios y guiones
   colapsados, números romanos unificados (`Ⅲ`→`III`, `T2`→`TII`).
3. **Agrupa variantes** por una clave alfanumérica del modelo. Dentro de cada
   grupo elige un *modelo canónico* (el más usado, prefiriendo el formato con
   guiones).
4. **Detecta typos** (segunda pasada): para los modelos que aparecen una sola
   vez (`singletons`), busca en el resto del catálogo de la misma marca un
   modelo mucho más popular con similitud Levenshtein ≥ 93 %. Si lo
   encuentra, lo propone como corrección con `STATUS = REVISAR`.
5. **Marca como `REVISAR`** los modelos basura (`-`, `*`, `?`, `N/A`, etc.) y
   los que no encuentran propuesta automática.

## Salida

Un Excel con las mismas columnas que el ejemplo Machida más dos hojas
auxiliares:

- `REVISION`: una fila por entrada con la `Propuesta modelo oficial / EUDAMED`,
  `Nota propuesta`, `MARCA_LIMPIA`, `MODELO_LIMPIO`, `MATCH_TIPO`
  (`EXACT` / `GROUP` / `FUZZY` / `NONE`), `CONFIANZA` 0-100 y `STATUS`.
- `RESUMEN`: contadores generales.
- `POR_MARCA`: filas / sin propuesta / a revisar por marca limpia.

## Uso

```bash
pip install pandas openpyxl rapidfuzz

python clean_brands_models.py \
    --input  data/Marcas_y_modelos_en_uso__EQ_ELECTROMEDICO.xlsx \
    --output output/Marcas_y_modelos_LIMPIO.xlsx
```

Opciones útiles:

- `--sheet REVISAR` — nombre de la hoja de entrada (por defecto `REVISAR`).
- `--fuzzy-threshold 92` — umbral 0-100 de similitud para proponer corrección
  por fuzzy matching cuando no hay grupo (por defecto `92`). El typo-detector
  de singletons usa internamente un umbral más estricto de 93.

## Formato esperado de entrada

Una hoja (por defecto llamada `REVISAR`) con tres columnas:

| Equipo_codiMarca | refMarca_nombre | Equipo_modelo |
| ---------------- | --------------- | ------------- |
| 12               | PHILIPS         | TP 300        |

## Enriquecimiento con EUDAMED

Opcionalmente puedes consultar la API pública de
[EUDAMED](https://ec.europa.eu/tools/eudamed/) para rellenar `TRADE_NAME`,
`UDI_DI`, `BASIC_UDI_DI`, `CLASE_RIESGO` y las columnas estilo Machida
`Encontrado en EUDAMED`, `Nombre comercial EUDAMED`, `Link de referencia`,
etc. Requiere acceso a `ec.europa.eu` desde la máquina donde lo ejecutes
(esto no funciona en el entorno remoto de Claude Code on the web por la
allowlist de red — hay que correrlo en local).

```bash
pip install requests

# Sólo para las marcas más grandes y como prueba (50 lookups máx):
python clean_brands_models.py \
    --input  data/Marcas_y_modelos_en_uso__EQ_ELECTROMEDICO.xlsx \
    --output output/Marcas_y_modelos_LIMPIO.xlsx \
    --eudamed \
    --eudamed-brands "MACHIDA" "PHILIPS" \
    --eudamed-max-lookups 50
```

Opciones:

- `--eudamed` — activa el lookup. Sólo consulta combinaciones únicas
  `(MARCA_LIMPIA, PROPUESTA_MODELO)`, no fila por fila.
- `--eudamed-cache PATH` — fichero JSON donde guardar/recuperar respuestas
  (default `output/.eudamed_cache.json`). Re-ejecuciones reaprovechan cache.
- `--eudamed-brands PHILIPS MACHIDA ...` — limita a esas marcas.
- `--eudamed-only-missing` — sólo modelos que aún no tienen `UDI_DI`.
- `--eudamed-max-lookups N` — corta tras N consultas (útil para pruebas).
- `--eudamed-delay 0.4` — pausa entre peticiones en segundos.

Si la API de EUDAMED cambia su contrato, el código está pensado para que
sólo haya que tocar la URL base y el extractor en `eudamed_client.py`
(`DEFAULT_BASE_URL` y la función `_flatten`).

## Tests

```bash
pip install pytest
pytest tests/
```

Los tests de EUDAMED no tocan la red: usan un `transport` mockeado.

## Personalización

- **Aliases de marca**: edita el diccionario `BRAND_ALIASES` en
  `clean_brands_models.py`.
- **Patrones de modelo no válido**: edita `INVALID_MODEL_PATTERNS`.
- **Lógica del modelo canónico**: edita `pick_canonical`.

## Resultado real (60 409 filas)

| Métrica                              | Valor  |
| ------------------------------------ | ------ |
| Filas totales                        | 60 409 |
| Marcas distintas (limpias)           | 5 002  |
| Propuesta automática `OK`            | 59 624 |
| `REVISAR` (typos o sin propuesta)    | 785    |
