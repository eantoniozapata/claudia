# Latinobarómetro – base consolidada

Pipeline para juntar todas las olas anuales de [Latinobarómetro](https://www.latinobarometro.org)
(1995 en adelante) en **una sola base armonizada**, con una fila por entrevista.

## 1. Descargar los microdatos

Los microdatos se descargan gratis pero **con registro** en
<https://www.latinobarometro.org/latContents.jsp> (sección *Análisis online / Descarga de datos*).
Descargue la versión SPSS (`.sav`) o Stata (`.dta`) de cada año que necesite y coloque los archivos,
o los `.zip` tal como vienen, en `data/raw/`. Los términos de uso no permiten redistribuirlos, por eso
`data/raw/` está en `.gitignore`.

## 2. Consolidar

```bash
pip install -r requirements.txt
python scripts/consolidar.py            # base armonizada
python scripts/consolidar.py --todas    # + todas las variables originales apiladas
python scripts/consolidar.py --stata    # + exporta también a .dta
```

Salidas en `data/processed/`:

| Archivo | Contenido |
|---|---|
| `latinobarometro_consolidado.parquet` / `.csv.gz` | Base armonizada: `anio`, `pais_cod`, `pais`, `iso3`, `ponderador` y los conceptos de `config/conceptos.csv` |
| `latinobarometro_todas_variables.parquet` | (con `--todas`) todas las variables de todas las olas, nombres en minúscula |
| `inventario_variables.csv` | Cada variable de cada ola con su etiqueta y número de casos válidos |
| `reporte_mapeo.csv` | Qué variable se usó para cada concepto y año, y qué otras candidatas había |

Para cada concepto hay dos columnas:
- `<concepto>`: el código numérico. Los códigos negativos de Latinobarómetro (no sabe, no responde,
  no preguntado) quedan como `NaN`.
- `<concepto>_txt`: la etiqueta original del valor, incluidos "No sabe" y "No responde".

Use siempre `ponderador` (`wt`) para obtener estimaciones representativas por país.

## 3. Cómo se armoniza

Los nombres de las preguntas cambian de un año a otro (por ejemplo `P12STGBS` en 2018 y `p10stgbs`
en 2020 para el apoyo a la democracia). El script identifica cada concepto por **nombre de la variable**
o por **etiqueta** (sin tildes ni mayúsculas) con las expresiones regulares de `config/conceptos.csv`.

Revise **siempre** `reporte_mapeo.csv` después de correrlo:
- Si un concepto quedó sin variable o tomó la variable equivocada, busque la correcta en
  `inventario_variables.csv` y agréguela en `config/mapeo_manual.csv`:
  ```csv
  anio,concepto,variable
  2018,apoyo_democracia,P12STGBS
  ```
  El mapeo manual siempre tiene prioridad sobre la detección automática.
- Para agregar un concepto nuevo, añada una fila a `config/conceptos.csv`.
- Las categorías de respuesta también pueden cambiar entre olas. Compare las columnas `_txt` antes
  de comparar códigos entre años.

## Pruebas

```bash
pytest tests
```
Las pruebas usan datos sintéticos con el formato de Latinobarómetro, así que no necesitan los microdatos.
