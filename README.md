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
