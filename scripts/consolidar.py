#!/usr/bin/env python3
"""Consolida las olas anuales de Latinobarómetro en una sola base.

Lee todos los archivos de microdatos (.sav, .dta o .zip que los contengan) de
``data/raw/`` y genera en ``data/processed/``:

- ``latinobarometro_consolidado.{parquet,csv.gz}``: base armonizada, una fila por
  entrevista, con los conceptos definidos en ``config/conceptos.csv``. Cada
  concepto se entrega como código numérico (los códigos negativos de
  Latinobarómetro —no sabe, no responde, no aplica— se pasan a NaN) y como
  texto de la etiqueta original (``<concepto>_txt``).
- ``inventario_variables.csv``: todas las variables de todas las olas con su
  etiqueta, para revisar o completar el mapeo.
- ``reporte_mapeo.csv``: qué variable se usó para cada concepto y año, y qué
  otras candidatas había.

Con ``--todas`` también genera ``latinobarometro_todas_variables.parquet``:
todas las variables originales apiladas (nombres en minúscula).

La detección automática se corrige con ``config/mapeo_manual.csv``
(anio, concepto, variable), que siempre tiene prioridad.
"""

from __future__ import annotations

import argparse
import re
import sys
import tempfile
import unicodedata
import zipfile
from pathlib import Path

import pandas as pd
import pyreadstat

RAIZ = Path(__file__).resolve().parent.parent
EXTENSIONES = {".sav": pyreadstat.read_sav, ".dta": pyreadstat.read_dta}


def normalizar(texto: str) -> str:
    """Minúsculas y sin tildes, para comparar etiquetas entre olas."""
    texto = unicodedata.normalize("NFKD", str(texto))
    return "".join(c for c in texto if not unicodedata.combining(c)).lower().strip()


def listar_archivos(directorio: Path, tmp: Path) -> list[Path]:
    archivos = []
    for ruta in sorted(directorio.rglob("*")):
        sufijo = ruta.suffix.lower()
        if sufijo in EXTENSIONES:
            archivos.append(ruta)
        elif sufijo == ".zip":
            destino = tmp / ruta.stem
            with zipfile.ZipFile(ruta) as zf:
                for nombre in zf.namelist():
                    if Path(nombre).suffix.lower() in EXTENSIONES and not nombre.startswith("__MACOSX"):
                        zf.extract(nombre, destino)
                        archivos.append(destino / nombre)
    return archivos


def leer(ruta: Path) -> tuple[pd.DataFrame, pyreadstat.metadata_container]:
    lector = EXTENSIONES[ruta.suffix.lower()]
    try:
        return lector(str(ruta), apply_value_formats=False)
    except Exception:
        # Algunos .sav antiguos vienen en latin-1 sin declararlo.
        return lector(str(ruta), apply_value_formats=False, encoding="LATIN1")


def detectar_anio(df: pd.DataFrame, ruta: Path) -> int:
    if "numinves" in df.columns:
        valores = pd.to_numeric(df["numinves"], errors="coerce").dropna()
        valores = valores[(valores >= 1995) & (valores <= 2100)]
        if not valores.empty:
            return int(valores.mode().iloc[0])
    encontrados = re.findall(r"(?<!\d)(19[89]\d|20\d\d)(?!\d)", ruta.name)
    if encontrados:
        return int(encontrados[-1])
    raise ValueError(f"No se pudo determinar el año de {ruta.name}: agregue el año al nombre del archivo")


def candidatas(concepto: pd.Series, etiquetas: dict[str, str]) -> list[str]:
    """Variables que coinciden por nombre (prioridad) o por etiqueta."""
    por_nombre, por_etiqueta = [], []
    if isinstance(concepto.regex_nombre, str) and concepto.regex_nombre:
        patron = re.compile(concepto.regex_nombre, re.I)
        por_nombre = [v for v in etiquetas if patron.search(v)]
    if isinstance(concepto.regex_etiqueta, str) and concepto.regex_etiqueta:
        patron = re.compile(concepto.regex_etiqueta, re.I)
        por_etiqueta = [v for v, e in etiquetas.items() if patron.search(e) and v not in por_nombre]
    return por_nombre + por_etiqueta


def texto_de(serie: pd.Series, etiquetas_valor: dict | None) -> pd.Series:
    if not etiquetas_valor:
        return serie.astype("string")
    mapa = {}
    for k, v in etiquetas_valor.items():
        mapa[k] = v
        if isinstance(k, float) and k.is_integer():
            mapa[int(k)] = v
    return serie.map(mapa).astype("string").fillna(serie.astype("string"))


def armonizar_ola(df, meta, anio, conceptos, manual, reporte) -> pd.DataFrame:
    etiquetas = {v: normalizar(meta.column_names_to_labels.get(orig) or "")
                 for v, orig in zip(df.columns, meta.column_names)}
    etiquetas_valor = {v: meta.variable_value_labels.get(orig)
                       for v, orig in zip(df.columns, meta.column_names)}
    salida = pd.DataFrame(index=df.index)
    for c in conceptos.itertuples(index=False):
        opciones = candidatas(c, etiquetas)
        elegida = manual.get((anio, c.concepto))
        origen = "manual"
        if elegida is not None and elegida not in df.columns:
            print(f"  ! {anio}: la variable manual '{elegida}' para {c.concepto} no existe", file=sys.stderr)
            elegida = None
        if elegida is None:
            elegida, origen = (opciones[0], "auto") if opciones else (None, "sin_dato")
        reporte.append({
            "anio": anio, "concepto": c.concepto, "variable": elegida, "origen": origen,
            "etiqueta": etiquetas.get(elegida, "") if elegida else "",
            "otras_candidatas": "; ".join(o for o in opciones if o != elegida),
        })
        if elegida is None:
            salida[c.concepto] = pd.NA
            continue
        numerica = pd.to_numeric(df[elegida], errors="coerce")
        if c.concepto not in {"pais_cod", "anio", "ponderador"}:
            salida[f"{c.concepto}_txt"] = texto_de(df[elegida], etiquetas_valor[elegida])
            numerica = numerica.mask(numerica < 0)
        salida[c.concepto] = numerica
    salida["anio"] = anio  # el año del archivo manda aunque numinves falte
    return salida


def main(argv=None) -> int:
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--entrada", type=Path, default=RAIZ / "data" / "raw")
    p.add_argument("--salida", type=Path, default=RAIZ / "data" / "processed")
    p.add_argument("--config", type=Path, default=RAIZ / "config")
    p.add_argument("--todas", action="store_true", help="también apila todas las variables originales")
    p.add_argument("--stata", action="store_true", help="también exporta la base armonizada a .dta")
    args = p.parse_args(argv)

    conceptos = pd.read_csv(args.config / "conceptos.csv", dtype=str, keep_default_na=False)
    paises = pd.read_csv(args.config / "paises.csv")
    manual_df = pd.read_csv(args.config / "mapeo_manual.csv", dtype={"anio": int, "concepto": str, "variable": str})
    manual = {(r.anio, r.concepto): r.variable.lower() for r in manual_df.itertuples()}

    args.salida.mkdir(parents=True, exist_ok=True)
    armonizadas, completas, inventario, reporte = [], [], [], []
    with tempfile.TemporaryDirectory() as tmp:
        archivos = listar_archivos(args.entrada, Path(tmp))
        if not archivos:
            print(f"No hay archivos .sav/.dta/.zip en {args.entrada}", file=sys.stderr)
            return 1
        for ruta in archivos:
            df, meta = leer(ruta)
            df.columns = [c.lower() for c in df.columns]
            anio = detectar_anio(df, ruta)
            print(f"{anio}: {ruta.name} ({len(df):,} casos, {df.shape[1]} variables)")
            for v, orig in zip(df.columns, meta.column_names):
                inventario.append({"anio": anio, "archivo": ruta.name, "variable": v,
                                   "etiqueta": meta.column_names_to_labels.get(orig),
                                   "n_validos": int(df[v].notna().sum())})
            ola = armonizar_ola(df, meta, anio, conceptos, manual, reporte)
            ola.insert(0, "archivo_origen", ruta.name)
            armonizadas.append(ola)
            if args.todas:
                completas.append(df.assign(anio_ola=anio, archivo_origen=ruta.name))

    base = pd.concat(armonizadas, ignore_index=True)
    base = base.merge(paises, how="left", left_on="pais_cod", right_on="idenpa").drop(columns="idenpa")
    base.insert(0, "id_consolidado", range(1, len(base) + 1))
    orden = ["id_consolidado", "anio", "pais_cod", "pais", "iso3", "ponderador", "archivo_origen"]
    base = base[orden + [c for c in base.columns if c not in orden]].sort_values(["anio", "pais_cod"], kind="stable")

    base.to_parquet(args.salida / "latinobarometro_consolidado.parquet", index=False)
    base.to_csv(args.salida / "latinobarometro_consolidado.csv.gz", index=False)
    if args.stata:
        exportable = base.copy()
        for col in exportable.select_dtypes(include=["string", "object"]).columns:
            exportable[col] = exportable[col].astype(object).where(exportable[col].notna(), "")
        pyreadstat.write_dta(exportable, str(args.salida / "latinobarometro_consolidado.dta"),
                             column_labels=[None] * exportable.shape[1])
    pd.DataFrame(inventario).to_csv(args.salida / "inventario_variables.csv", index=False)
    rep = pd.DataFrame(reporte)
    rep.to_csv(args.salida / "reporte_mapeo.csv", index=False)
    if completas:
        todas = pd.concat(completas, ignore_index=True)
        # Una misma variable puede cambiar de tipo entre olas; parquet exige uno solo.
        for col in todas.columns[todas.dtypes == object]:
            todas[col] = todas[col].astype("string")
        todas.to_parquet(args.salida / "latinobarometro_todas_variables.parquet", index=False)

    faltantes = rep[rep.variable.isna()]
    print(f"\nBase consolidada: {len(base):,} casos, {base.anio.nunique()} olas, "
          f"{base.pais.nunique()} países -> {args.salida}")
    if not faltantes.empty:
        print(f"{len(faltantes)} combinaciones año-concepto sin variable; revise reporte_mapeo.csv "
              "y complete config/mapeo_manual.csv")
    return 0


if __name__ == "__main__":
    sys.exit(main())
