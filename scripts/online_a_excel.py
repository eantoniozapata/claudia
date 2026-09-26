#!/usr/bin/env python3
"""Convierte exportaciones de la herramienta de análisis online de Latinobarómetro
(tablas país x año) en un libro de Excel con un panel país-año para regresión,
con las variables calculadas con el criterio A (NS/NR en el denominador) y B
(solo respuestas válidas).

Uso: python scripts/online_a_excel.py
Lee data/online/*.xlsx según VARIABLES y escribe excel/latinobarometro_panel.xlsx.
"""

import csv
import re
import unicodedata
from pathlib import Path

import openpyxl
import pandas as pd
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

RAIZ = Path(__file__).resolve().parent.parent
SALIDA = RAIZ / "excel" / "latinobarometro_panel.xlsx"

# nombre: archivo, categorías, categoría de interés, categorías válidas (denominador B)
VARIABLES = {
    "Apoyo democracia": dict(
        # Serie histórica (país x año) y olas sueltas (un año, países en columnas). Las olas sueltas
        # reemplazan a la serie en los país-año que ambas traen (2015: la serie solo trae 4 países).
        archivos=["data/online/apoyo_democracia.xlsx", "data/online/apoyo_democracia_2013.xlsx",
                  "data/online/apoyo_democracia_2015.xlsx"],
        categorias=["Democracia", "Gobierno Autoritario", "Da lo mismo", "No sabe", "No contesta"],
        interes="Democracia",
        validas=["Democracia", "Gobierno Autoritario", "Da lo mismo"],
    ),
    "Grupos poderosos": dict(
        archivos=["data/online/grupos_poderosos.xlsx", "data/online/grupos_poderosos_2013.xlsx",
                  "data/online/grupos_poderosos_2015.xlsx"],
        categorias=["Grupos poderosos en su propio beneficio", "Para el bien de todo el pueblo",
                    "No sabe; no responde"],
        interes="Grupos poderosos en su propio beneficio",
        validas=["Grupos poderosos en su propio beneficio", "Para el bien de todo el pueblo"],
    ),
}
DEPENDIENTE = "Apoyo democracia"

# Índice de Percepción de la Corrupción (Transparency International)
CPI_DATAHUB = "data/corrupcion/cpi_datahub.csv"      # datahub.io, escala 0-10, 1995-2017
CPI_GLOBAL = "data/corrupcion/global_cpi_all_hdx.csv"  # data.humdata.org, escala 0-100, 2012-2025
ULTIMO_ANIO_DATAHUB = 2011
ISO = {"Argentina": "ARG", "Bolivia": "BOL", "Brasil": "BRA", "Chile": "CHL", "Colombia": "COL",
       "Costa Rica": "CRI", "Rep. Dominicana": "DOM", "Ecuador": "ECU", "El Salvador": "SLV",
       "Guatemala": "GTM", "Honduras": "HND", "México": "MEX", "Nicaragua": "NIC", "Panamá": "PAN",
       "Paraguay": "PRY", "Perú": "PER", "España": "ESP", "Uruguay": "URY", "Venezuela": "VEN"}
NOMBRES_DATAHUB = {"Brasil": ["Brazil"], "México": ["Mexico"], "Panamá": ["Panama"], "Perú": ["Peru"],
                   "España": ["Spain"],
                   "Rep. Dominicana": ["Dominican Republic", "Dominican Rep", "Dominican Rep."]}
# Controles país-año
GINI_BM = "data/controles/gini_bm.xls"  # Banco Mundial SI.POV.GINI, 0-100
# nombre de columna: (archivo, fuente, formato)
CONTROLES = {
    "Homicidios por 100.000 hab.": ("data/controles/homicidios_bm.xls",
                                    "Banco Mundial, WDI VC.IHR.PSRC.P5 (UNODC)", "0.0"),
    "Crecimiento del PIB (% anual)": ("data/controles/crecimiento_pib_bm.xlsx",
                                      "Banco Mundial, WDI NY.GDP.MKTP.KD.ZG", "0.0"),
    "Gasto social (% del PIB, gob. central)": ("data/controles/gasto_social_cepal.xlsx",
                                               "CEPALSTAT, gasto público social del gobierno central", "0.0"),
}
NOMBRES_CEPAL = {"Bolivia (Estado Plurinacional de)": "Bolivia", "República Dominicana": "Rep. Dominicana",
                 "Venezuela (República Bolivariana de)": "Venezuela"}
# Empalme por regresión: CPI nuevo (2012) = a + b * CPI antiguo (2011), con todos los países
ANIO_PUENTE_NUEVO = 2012
# Nombres de datahub (2011) que no coinciden con los de HDX
NOMBRES_DATAHUB_ISO = {"Czech Republic": "CZE", "Korea (North)": "PRK", "Korea (South)": "KOR",
                       "United States": "USA", "Brunei": "BRN", "Laos": "LAO", "Syria": "SYR", "Iran": "IRN",
                       "Russia": "RUS", "Cape Verde": "CPV", "Swaziland": "SWZ", "Vietnam": "VNM",
                       "Congo  Republic": "COG", "Democratic Republic of the Congo": "COD",
                       "Sao Tome & Principe": "STP", "FYR Macedonia": "MKD", "Macau": "MAC",
                       "Cote d´Ivoire": "CIV", "Côte d´Ivoire": "CIV"}

ARIAL = "Arial"
F_NORMAL = Font(name=ARIAL, size=10)
F_TITULO = Font(name=ARIAL, size=10, bold=True, color="FFFFFF")
F_INPUT = Font(name=ARIAL, size=10, color="0000FF")
F_LINK = Font(name=ARIAL, size=10, color="008000")
RELLENO_TITULO = PatternFill("solid", fgColor="1F4E78")
RELLENO_VAR = PatternFill("solid", fgColor="DDEBF7")
RELLENO_INPUT = PatternFill("solid", fgColor="FFFF00")
BORDE = Border(bottom=Side(style="thin", color="BFBFBF"))
MAXFILA = 5000  # rangos acotados para que SUMIFS sea rápido


def a_numero(v):
    return float(v.rstrip("%")) / 100


def leer_ola(filas, enc):
    """Exportación de un solo estudio: categorías en filas, países en columnas."""
    anio = int(re.search(r"\d{4}", next(str(v) for v in filas[enc - 4] if v)).group())
    paises = filas[enc][2:]  # la columna B es "Total"
    fila_n = next(f for f in filas[enc + 1:] if f[0] and str(f[0]).startswith("(N)"))
    datos = []
    for f in filas[enc + 1:]:
        if not f[0] or str(f[0]).startswith("(N)"):
            continue
        for pais, v, n in zip(paises, f[2:], fila_n[2:]):
            if pais and v not in (None, "", "-"):
                datos.append((pais, anio, f[0], a_numero(v), int(str(n).split(" ")[0].replace(",", ""))))
    return datos


def leer_exportacion(ruta: Path):
    """Devuelve [(país, año, categoría, proporción, n)] omitiendo '-' y el bloque Total."""
    filas = list(openpyxl.load_workbook(ruta).active.iter_rows(values_only=True))
    enc = next(i for i, f in enumerate(filas) if f[0] == "Restaurar vista")
    if not str(filas[enc][2]).isdigit():  # columnas = países: estudio de un solo año
        return leer_ola(filas, enc)
    anios = filas[enc][2:]  # la columna B es "Total" (todas las olas juntas)
    datos, pais, bloque = [], None, []
    for f in filas[enc + 1:]:
        if not f[0]:
            continue
        if all(v in (None, "") for v in f[1:]):
            pais, bloque = f[0], []
        elif str(f[0]).startswith("(N)"):
            if pais != "Total":
                for j, anio in enumerate(anios):
                    n = str(f[2 + j]).split(" ")[0].replace(",", "")
                    for cat, valores in bloque:
                        v = valores[j]
                        if v not in (None, "", "-"):
                            datos.append((pais, int(anio), cat, a_numero(v), int(n)))
        else:
            bloque.append((f[0], f[2:]))
    return datos


def leer_cpi():
    """{(país, año): valor} de datahub (0-10, <= 2011; 0.0 = sin dato) y de HDX (0-100)."""
    datahub, glob = {}, {}
    with open(RAIZ / CPI_DATAHUB, encoding="utf-8") as f:
        filas = list(csv.reader(f))
    anios = [int(a) for a in filas[0][1:]]
    for pais in ISO:
        nombres = NOMBRES_DATAHUB.get(pais, [pais])
        for fila in filas[1:]:
            if fila[0] in nombres:
                for a, v in zip(anios, fila[1:]):
                    if a <= ULTIMO_ANIO_DATAHUB and float(v) > 0:
                        datahub[(pais, a)] = float(v)
    pais_de = {i: p for p, i in ISO.items()}
    with open(RAIZ / CPI_GLOBAL, encoding="utf-8") as f:
        for r in csv.DictReader(f):
            if r["iso3"] in pais_de and r["score"]:
                glob[(pais_de[r["iso3"]], int(r["year"]))] = float(r["score"])
    return datahub, glob


def leer_banco_mundial(ruta):
    """{(país, año): valor} de un archivo de Indicadores del Desarrollo Mundial (hoja Data)."""
    datos = pd.read_excel(RAIZ / ruta, sheet_name="Data", header=3)
    pais_de = {i: p for p, i in ISO.items()}
    datos = datos[datos["Country Code"].isin(pais_de)]
    largo = datos.melt(id_vars="Country Code", value_vars=[c for c in datos.columns if str(c)[:2] in ("19", "20")],
                       var_name="anio", value_name="valor").dropna()
    return {(pais_de[r["Country Code"]], int(float(r["anio"]))): float(r["valor"]) for _, r in largo.iterrows()}


def leer_cepal(ruta):
    """{(país, año): valor} de una descarga de CEPALSTAT (hoja datos)."""
    datos = pd.read_excel(RAIZ / ruta, sheet_name="datos")
    datos["pais"] = datos["País__ESTANDAR"].replace(NOMBRES_CEPAL)
    datos = datos[datos["pais"].isin(ISO)]
    return {(r["pais"], int(r["Años__ESTANDAR"])): float(r["value"]) for _, r in datos.iterrows()}


def normalizar_nombre(nombre: str) -> str:
    nombre = unicodedata.normalize("NFKD", nombre).encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^a-z]", "", nombre)


def leer_puente():
    """Países con CPI 2011 (datahub, 0-10) y CPI 2012 (HDX, 0-100): [(país, iso3, cpi2011, cpi2012)]."""
    iso_de, cpi2012 = {}, {}
    with open(RAIZ / CPI_GLOBAL, encoding="utf-8") as f:
        for r in csv.DictReader(f):
            iso_de[normalizar_nombre(r["country"])] = r["iso3"]
            if int(r["year"]) == ANIO_PUENTE_NUEVO and r["score"]:
                cpi2012[r["iso3"]] = (r["country"], float(r["score"]))
    iso_de.update({normalizar_nombre(n): i for n, i in NOMBRES_DATAHUB_ISO.items()})
    cpi2011 = {}
    with open(RAIZ / CPI_DATAHUB, encoding="utf-8") as f:
        filas = list(csv.reader(f))
    col = filas[0].index(str(ANIO_PUENTE_NUEVO - 1))
    for fila in filas[1:]:
        iso = iso_de.get(normalizar_nombre(fila[0]))
        if iso and float(fila[col]) > 0:
            cpi2011.setdefault(iso, float(fila[col]))
    return sorted((cpi2012[i][0], i, cpi2011[i], cpi2012[i][1]) for i in cpi2011.keys() & cpi2012.keys())


def leer_variable(cfg):
    """Une los archivos de una variable; un archivo posterior reemplaza los país-año de los anteriores."""
    por_clave = {}
    for ruta in cfg["archivos"]:
        filas = leer_exportacion(RAIZ / ruta)
        for clave in {(p, a) for p, a, *_ in filas}:
            por_clave[clave] = [f for f in filas if (f[0], f[1]) == clave]
    return [f for clave in sorted(por_clave) for f in por_clave[clave]]


def encabezado(ws, fila, textos, relleno=RELLENO_TITULO, fuente=F_TITULO):
    for j, t in enumerate(textos, 1):
        c = ws.cell(fila, j, t)
        c.font, c.fill = fuente, relleno
        c.alignment = Alignment(wrap_text=True, vertical="center")


def anchos(ws, anchos_col):
    for j, a in enumerate(anchos_col, 1):
        ws.column_dimensions[get_column_letter(j)].width = a


def main():
    datos = {v: leer_variable(cfg) for v, cfg in VARIABLES.items()}
    paises = list(dict.fromkeys(p for p, *_ in datos[DEPENDIENTE]))
    panel = sorted({(p, a) for p, a, *_ in datos[DEPENDIENTE]}, key=lambda x: (paises.index(x[0]), x[1]))

    wb = openpyxl.Workbook()
    leeme = wb.active
    leeme.title = "Leeme"

    # --- Datos: todas las exportaciones en formato largo -----------------------
    ws = wb.create_sheet("Datos")
    encabezado(ws, 1, ["Variable", "País", "Año", "Categoría", "Porcentaje", "N (casos)"])
    fila = 2
    for var, filas in datos.items():
        for p, a, cat, pct, n in filas:
            for j, v in enumerate([var, p, a, cat, pct, n], 1):
                c = ws.cell(fila, j, v)
                c.font = F_INPUT
            ws.cell(fila, 5).number_format = "0.0%"
            ws.cell(fila, 6).number_format = "#,##0"
            fila += 1
    ultima_datos = fila - 1
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:F{ultima_datos}"
    anchos(ws, [20, 18, 8, 40, 12, 11])

    # --- Categorias: % de cada categoría por país-año (fórmulas) --------------
    wc = wb.create_sheet("Categorias")
    cols = [(v, c) for v, cfg in VARIABLES.items() for c in cfg["categorias"]]
    encabezado(wc, 1, ["Variable →", ""] + [v for v, _ in cols])
    encabezado(wc, 2, ["País", "Año"] + [c for _, c in cols])
    rng = lambda col: f"Datos!${col}$2:${col}${MAXFILA}"
    for i, (p, a) in enumerate(panel, 3):
        wc.cell(i, 1, p).font = F_NORMAL
        wc.cell(i, 2, a).font = F_NORMAL
        for j in range(3, 3 + len(cols)):
            L = get_column_letter(j)
            crit = (f"{rng('A')},{L}$1,{rng('B')},$A{i},{rng('C')},$B{i},{rng('D')},{L}$2")
            c = wc.cell(i, j, f'=IF(COUNTIFS({crit})=0,"",SUMIFS({rng("E")},{crit}))')
            c.font, c.number_format = F_NORMAL, "0.0%"
    wc.freeze_panes = "C3"
    anchos(wc, [16, 7] + [14] * len(cols))
    col_de = {vc: get_column_letter(3 + k) for k, vc in enumerate(cols)}

    # --- Gini: insumo a completar -------------------------------------------
    wg = wb.create_sheet("Gini")
    gini = leer_banco_mundial(GINI_BM)
    encabezado(wg, 1, ["País", "Año", "Gini (0-100)", "Fuente", "Clave", "Año anterior con dato",
                       "Gini anterior", "Año siguiente con dato", "Gini siguiente", "Gini interpolado",
                       "Método", "", "Usar Gini interpolado en los paneles (1 = sí, 0 = no)"])
    for i, (p, a) in enumerate(panel, 2):
        wg.cell(i, 1, p).font = F_NORMAL
        wg.cell(i, 2, a).font = F_NORMAL
        if (p, a) in gini:
            wg.cell(i, 3, gini[(p, a)])
            wg.cell(i, 4, "Banco Mundial, WDI SI.POV.GINI")
        else:  # sin dato del Banco Mundial: se puede completar con otra fuente
            for j in (3, 4):
                wg.cell(i, j).fill = RELLENO_INPUT
        for j in (3, 4):
            wg.cell(i, j).font = F_INPUT
        wg.cell(i, 3).number_format = "0.0"
        wg.cell(i, 5, f'=A{i}&"|"&B{i}').font = F_NORMAL
        # Interpolación lineal entre el último año observado antes y el primero después (sin extrapolar)
        if (p, a) not in gini:
            antes = [y for (q, y) in gini if q == p and y < a]
            despues = [y for (q, y) in gini if q == p and y > a]
            if antes and despues:
                y0, y1 = max(antes), min(despues)
                for j, v in ((6, y0), (7, gini[(p, y0)]), (8, y1), (9, gini[(p, y1)])):
                    wg.cell(i, j, v).font = F_INPUT
                wg.cell(i, 7).number_format = wg.cell(i, 9).number_format = "0.0"
        wg.cell(i, 10, f'=IF(C{i}<>"",C{i},IF(OR(F{i}="",H{i}=""),"",G{i}+(I{i}-G{i})*(B{i}-F{i})/(H{i}-F{i})))')
        wg.cell(i, 10).number_format = "0.0"
        wg.cell(i, 11, f'=IF(C{i}<>"","Observado",IF(J{i}="","Sin dato","Interpolado "&F{i}&"-"&H{i}))')
        for j in (10, 11):
            wg.cell(i, j).font = F_NORMAL
    wg["M2"] = 0
    wg["M2"].fill, wg["M2"].font = RELLENO_INPUT, Font(name=ARIAL, size=12, bold=True, color="0000FF")
    wg["M3"] = "0 = Gini observado (columna C). 1 = Gini interpolado (columna J)."
    wg["M3"].font = F_NORMAL
    wg["F1"].comment = Comment("Años y valores observados del Banco Mundial (serie anual completa, no solo los "
                               "años del panel) usados para interpolar. Sin valor antes o después = no se "
                               "extrapola.", "Claude")
    wg["C1"].comment = Comment("Banco Mundial, Indicadores del Desarrollo Mundial, SI.POV.GINI (0-100). "
                               "Celdas amarillas = sin dato para ese año; puede completarlas con otra fuente "
                               "(p. ej. SEDLAC) o dejarlas vacías (la fila queda fuera de la regresión).", "Claude")
    wg.freeze_panes = "A2"
    anchos(wg, [16, 7, 11, 30, 20, 11, 10, 11, 10, 11, 20, 3, 34])
    wg.row_dimensions[1].height = 42
    ultima_gini = len(panel) + 1

    # --- Puente_CPI: regresión 2012 (método nuevo) sobre 2011 (método antiguo) --
    puente = leer_puente()
    wq = wb.create_sheet("Puente_CPI")
    encabezado(wq, 1, ["País (HDX)", "ISO3", f"CPI {ANIO_PUENTE_NUEVO - 1} datahub (0-10)",
                       f"CPI {ANIO_PUENTE_NUEVO} HDX (0-100)", "", "Coeficientes del empalme", "Valor"])
    for i, fila in enumerate(puente, 2):
        for j, v in enumerate(fila, 1):
            wq.cell(i, j, v).font = F_INPUT if j > 2 else F_NORMAL
    uq = len(puente) + 1
    X, Y = f"$C$2:$C${uq}", f"$D$2:$D${uq}"
    coefs = [("Intercepto (a)", f"=INTERCEPT({Y},{X})", "0.000"),
             ("Pendiente (b)", f"=SLOPE({Y},{X})", "0.000"),
             ("R cuadrado", f"=RSQ({Y},{X})", "0.000"),
             ("N países", f"=COUNT({X})", "0")]
    for k, (nombre, formula, fmt) in enumerate(coefs, 2):
        wq.cell(k, 6, nombre).font = Font(name=ARIAL, size=10, bold=True)
        c = wq.cell(k, 7, formula)
        c.font, c.number_format = F_NORMAL, fmt
    wq.cell(7, 6, f"CPI empalmado (años <= {ANIO_PUENTE_NUEVO - 1}) = a + b x CPI datahub (0-10). "
                  f"Desde {ANIO_PUENTE_NUEVO}: CPI HDX sin cambios.").font = F_NORMAL
    wq.cell(8, 6, f"Supuesto: la corrupción percibida no cambió realmente entre {ANIO_PUENTE_NUEVO - 1} y "
                  f"{ANIO_PUENTE_NUEVO}; la diferencia se atribuye al cambio de metodología.").font = F_NORMAL
    wq.freeze_panes = "A2"
    anchos(wq, [30, 7, 14, 14, 3, 24, 10])
    wq.row_dimensions[1].height = 30

    # --- Corrupcion: CPI de Transparency International -----------------------
    datahub, glob = leer_cpi()
    wk = wb.create_sheet("Corrupcion")
    encabezado(wk, 1, ["País", "Año", "CPI datahub (0-10, hasta 2011)", "CPI global HDX (0-100, desde 2012)",
                       "CPI 0-100 (solo desde 2012)",
                       "CPI 0-100 empalmado por regresión", "Clave"])
    for i, (p, a) in enumerate(panel, 2):
        wk.cell(i, 1, p).font = F_NORMAL
        wk.cell(i, 2, a).font = F_NORMAL
        if (p, a) in datahub:
            wk.cell(i, 3, datahub[(p, a)])
        if a > ULTIMO_ANIO_DATAHUB and (p, a) in glob:
            wk.cell(i, 4, glob[(p, a)])
        for j in (3, 4):
            wk.cell(i, j).font = F_INPUT
        wk.cell(i, 5, f'=IF(OR(B{i}<={ULTIMO_ANIO_DATAHUB},D{i}=""),"",D{i})')
        wk.cell(i, 6, f'=IF(B{i}<={ULTIMO_ANIO_DATAHUB},IF(C{i}="","",Puente_CPI!$G$2+Puente_CPI!$G$3*C{i}),'
                      f'IF(D{i}="","",D{i}))')
        wk.cell(i, 7, f'=A{i}&"|"&B{i}')
        for j in (5, 7):
            wk.cell(i, j).font = F_NORMAL
        wk.cell(i, 6).font = F_LINK
        wk.cell(i, 3).number_format = "0.00"
        wk.cell(i, 4).number_format = wk.cell(i, 5).number_format = "0"
        wk.cell(i, 6).number_format = "0.0"
    wk["C1"].comment = Comment("Fuente: datahub.io/core/corruption-perceptions-index (Transparency "
                               "International). 0.0 en el archivo original = sin dato (queda vacío).", "Claude")
    wk["D1"].comment = Comment("Fuente: data.humdata.org/dataset/global-corruption-perceptions-index "
                               "(Transparency International), columna score.", "Claude")
    wk.freeze_panes = "C2"
    wk.auto_filter.ref = f"A1:G{len(panel) + 1}"
    anchos(wk, [16, 7, 16, 18, 16, 18, 22])
    wk.row_dimensions[1].height = 42
    ultima_cpi = len(panel) + 1

    # --- Controles: homicidios, crecimiento y gasto social ---------------------
    controles = {n: (leer_banco_mundial(r) if "bm" in r else leer_cepal(r)) for n, (r, _, _) in CONTROLES.items()}
    wt = wb.create_sheet("Controles")
    encabezado(wt, 1, ["País", "Año"] + list(CONTROLES) + ["Clave"])
    for j, (_, fuente, _) in enumerate(CONTROLES.values(), 3):
        wt.cell(1, j).comment = Comment(f"Fuente: {fuente}. Vacío = sin dato para ese país-año.", "Claude")
    for i, (p, a) in enumerate(panel, 2):
        wt.cell(i, 1, p).font = F_NORMAL
        wt.cell(i, 2, a).font = F_NORMAL
        for j, (n, (_, _, fmt)) in enumerate(CONTROLES.items(), 3):
            c = wt.cell(i, j, controles[n].get((p, a)))
            c.font, c.number_format = F_INPUT, fmt
        wt.cell(i, 3 + len(CONTROLES), f'=A{i}&"|"&B{i}').font = F_NORMAL
    col_clave_ctrl = get_column_letter(3 + len(CONTROLES))
    wt.freeze_panes = "C2"
    wt.auto_filter.ref = f"A1:{col_clave_ctrl}{len(panel) + 1}"
    anchos(wt, [16, 7] + [18] * len(CONTROLES) + [22])
    wt.row_dimensions[1].height = 42

    # --- Panel_A y Panel_B: listos para regresión -----------------------------
    for criterio in ("A", "B"):
        wp = wb.create_sheet(f"Panel_{criterio}")
        otras = [v for v in VARIABLES if v != DEPENDIENTE]
        titulos = (["País", "Año", f"Y: {DEPENDIENTE} ({criterio})"]
                   + [f"X: {v} ({criterio})" for v in otras]
                   + ["X: Gini (observado o interpolado según Gini!M2)", "X: CPI 0-100 (desde 2012)",
                      "X: CPI 0-100 empalmado por regresión"]
                   + [f"X: {n}" for n in CONTROLES]
                   + ["Completo con CPI 0-100 (1 = usar)",
                      "Completo con CPI regresión (1 = usar)", "Controles completos (1 = sí)"])
        encabezado(wp, 1, titulos)
        for i, (p, a) in enumerate(panel, 2):
            r = i + 1  # fila correspondiente en Categorias
            wp.cell(i, 1, p).font = F_NORMAL
            wp.cell(i, 2, a).font = F_NORMAL
            for j, v in enumerate([DEPENDIENTE] + otras, 3):
                cfg = VARIABLES[v]
                interes = f"Categorias!{col_de[(v, cfg['interes'])]}{r}"
                if criterio == "A":
                    f = f'=IF({interes}="","",{interes})'
                else:
                    denom = "+".join(f"N(Categorias!{col_de[(v, c)]}{r})" for c in cfg["validas"])
                    f = f'=IF({interes}="","",{interes}/({denom}))'
                c = wp.cell(i, j, f)
                c.font, c.number_format = F_LINK, "0.0%"
            jg = 3 + 1 + len(otras)
            idx = (f"INDEX(Gini!$A$2:$J${ultima_gini},MATCH($A{i}&\"|\"&$B{i},Gini!$E$2:$E${ultima_gini},0),"
                   f"IF(Gini!$M$2=1,10,3))")
            wp.cell(i, jg, f'=IFERROR(IF({idx}="","",{idx}),"")').font = F_LINK
            wp.cell(i, jg).number_format = "0.0"
            for k, col, fmt in ((1, "E", "0"), (2, "F", "0.0")):
                idx = (f"INDEX(Corrupcion!${col}$2:${col}${ultima_cpi},"
                       f"MATCH($A{i}&\"|\"&$B{i},Corrupcion!$G$2:$G${ultima_cpi},0))")
                c = wp.cell(i, jg + k, f'=IFERROR(IF({idx}="","",{idx}),"")')
                c.font, c.number_format = F_LINK, fmt
            uc = len(panel) + 1
            for k, (n, (_, _, fmt)) in enumerate(CONTROLES.items(), 3):
                col = get_column_letter(3 + list(CONTROLES).index(n))
                idx = (f"INDEX(Controles!${col}$2:${col}${uc},"
                       f"MATCH($A{i}&\"|\"&$B{i},Controles!${col_clave_ctrl}$2:${col_clave_ctrl}${uc},0))")
                c = wp.cell(i, jg + k, f'=IFERROR(IF({idx}="","",{idx}),"")')
                c.font, c.number_format = F_LINK, fmt
            jc = jg + 2 + len(CONTROLES)  # última columna de controles
            base = f"COUNTBLANK(C{i}:{get_column_letter(jg)}{i})"
            for k in (1, 2):
                L = get_column_letter(jg + k)
                wp.cell(i, jc + k, f"=IF({base}+COUNTBLANK({L}{i})=0,1,0)").font = F_NORMAL
            ctrl = f"{get_column_letter(jg + 3)}{i}:{get_column_letter(jc)}{i}"
            wp.cell(i, jc + 3, f"=IF(COUNTBLANK({ctrl})=0,1,0)").font = F_NORMAL
        ult = len(panel) + 1
        wp.freeze_panes = "C2"
        wp.auto_filter.ref = f"A1:{get_column_letter(len(titulos))}{ult}"
        anchos(wp, [16, 7] + [22] * (len(titulos) - 2))
        wp.row_dimensions[1].height = 42

    # --- Leeme ---------------------------------------------------------------
    txt = [
        ("Panel país-año Latinobarómetro para regresión (MCO)", True),
        ("", False),
        ("Fuente: exportaciones de la herramienta de análisis online de Latinobarómetro "
         "(porcentajes ya ponderados). Archivos: data/online/apoyo_democracia*.xlsx y "
         "data/online/grupos_poderosos*.xlsx (serie histórica + estudios 2013 y 2015 exportados aparte, porque "
         "la serie histórica no trae 2013 y en 2015 solo trae 4 países).", False),
        ("", False),
        ("HOJAS", True),
        ("Datos: todas las tablas exportadas en formato largo (una fila por variable, país, año y categoría). "
         "Valores copiados de la herramienta online (texto azul).", False),
        ("Categorias: % de cada categoría por país-año, calculado con SUMIFS sobre Datos.", False),
        ("Gini: Banco Mundial (SI.POV.GINI, 0-100). Las celdas amarillas no tienen dato del Banco Mundial para ese "
         "año; puede completarlas con otra fuente (p. ej. SEDLAC) o dejarlas vacías. 'Gini interpolado' (col. J) "
         "llena los huecos con una recta entre el año observado anterior y el siguiente (no extrapola). "
         "La celda Gini!M2 decide qué versión usan los paneles: 0 = observado, 1 = interpolado.", False),
        ("Controles: homicidios por 100.000 hab. (Banco Mundial/UNODC), crecimiento del PIB % anual (Banco "
         "Mundial) y gasto público social del gobierno central en % del PIB (CEPALSTAT, 2010-2023).", False),
        ("Corrupcion: Índice de Percepción de la Corrupción (Transparency International). Dos versiones: "
         "'CPI 0-100 (desde 2012)' = global HDX sin transformar; 'CPI 0-100 empalmado por regresión' = hasta 2011 "
         "a + b x datahub, con a y b estimados en la hoja Puente_CPI (176 países, CPI 2012 sobre CPI 2011), "
         "y desde 2012 global HDX.", False),
        ("Puente_CPI: datos de los 176 países con CPI 2011 (método antiguo) y 2012 (método nuevo) y los "
         "coeficientes del empalme (INTERSECCION.EJE, PENDIENTE, R2).", False),
        ("Panel_A: variables con el criterio A = % de la categoría sobre el total (incluye NS/NR). Robustez.", False),
        ("Panel_B: variables con el criterio B = % sobre respuestas válidas (excluye NS/NR). Modelo principal.", False),
        ("", False),
        ("DEFINICIONES", True),
        ("Y = % 'La democracia es preferible a cualquier otra forma de gobierno'. "
         "B: Democracia / (Democracia + Gobierno autoritario + Da lo mismo).", False),
        ("X grupos poderosos = % 'gobernado por unos cuantos grupos poderosos en su propio beneficio'. "
         "B: Grupos / (Grupos + Bien de todo el pueblo).", False),
        ("Ejemplo Argentina 2024: A = 74,6 %; B = 74,6 / (74,6 + 11,6 + 10,1) = 77,5 %.", False),
        ("CPI: más alto = MENOS corrupción percibida. Transparency International cambió la metodología en 2012 y "
         "advierte que los puntajes anteriores no son comparables con los posteriores: la empalmada por regresión corrige la escala (no el hecho de que antes "
         "de 2012 los cambios entre años no son confiables: use efectos fijos de año); la versión 0-100 no lo cruza.", False),
        ("", False),
        ("CÓMO CORRER LA REGRESIÓN", True),
        ("1. Revise la hoja Gini. 2. En Panel_A o Panel_B filtre la columna 'Completo' de la versión de CPI "
         "que va a usar (= 1) y, si incluye controles, también 'Controles completos' (= 1). 3. Copie las filas visibles y péguelas como valores en una hoja nueva; "
         "elimine las columnas de CPI que no usa para que las X queden contiguas. "
         "4. Datos > Análisis de datos > Regresión: Rango Y = columna Y; Rango X = columnas X (contiguas).", False),
        ("", False),
        ("CÓMO AGREGAR OTRA VARIABLE DE LATINOBARÓMETRO", True),
        ("1. Pegue sus filas al final de Datos con un nombre en 'Variable' (mismo formato: país, año, categoría, "
         "proporción). 2. En Categorias agregue una columna por categoría: fila 1 = nombre de la variable, "
         "fila 2 = categoría exacta; copie la fórmula de una columna vecina. 3. En Panel_A y Panel_B inserte "
         "una columna ANTES de 'X: Gini' (así las columnas 'Completo' la incluyen) con A = categoría de interés y "
         "B = categoría de interés / suma de categorías válidas.", False),
        ("", False),
        ("NOTAS", True),
        ("'-' en la herramienta = pregunta no aplicada en ese país-año; queda vacío (nunca 0).", False),
        ("2023: la serie histórica no entrega datos por país (solo el total regional); falta exportar el estudio "
         "2023 por separado. 2013 y 2015 vienen de los estudios de cada año.", False),
        ("España solo tiene 'Apoyo democracia' (no la pregunta de grupos poderosos): sus filas quedan con "
         "Completo = 0. 'Grupos poderosos' existe desde 2004 y falta en algunos país-año (p. ej. 2015).", False),
        ("CPI México 2015: global HDX (y el Excel oficial de TI) dice 31; datahub dice 3,5. Se usa HDX según "
         "la regla (desde 2012). Otros 10 países difieren en 2015, ninguno de Latinobarómetro.", False),
        ("Gasto social: solo gobierno central (subestima el gasto en países federales como Argentina, Brasil y "
         "México); no hay dato para España y Venezuela solo llega a 2014. Con este control el panel empieza en 2010.", False),
        ("Colores: azul = dato de origen; verde = vínculo a otra hoja; negro = cálculo; amarillo = completar.", False),
    ]
    for i, (t, negrita) in enumerate(txt, 1):
        c = leeme.cell(i, 1, t)
        c.font = Font(name=ARIAL, size=12 if i == 1 else 10, bold=negrita)
        c.alignment = Alignment(wrap_text=True, vertical="top")
    leeme.column_dimensions["A"].width = 120

    wb.calculation.fullCalcOnLoad = True  # Excel calcula todo al abrir
    SALIDA.parent.mkdir(exist_ok=True)
    wb.save(SALIDA)
    print(f"{SALIDA} | {len(panel)} país-año | filas en Datos: {ultima_datos - 1}")


if __name__ == "__main__":
    main()
