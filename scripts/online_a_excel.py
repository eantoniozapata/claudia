#!/usr/bin/env python3
"""Convierte exportaciones de la herramienta de análisis online de Latinobarómetro
(tablas país x año) en un libro de Excel con un panel país-año para regresión,
con las variables calculadas con el criterio A (NS/NR en el denominador) y B
(solo respuestas válidas).

Uso: python scripts/online_a_excel.py
Lee data/online/*.xlsx según VARIABLES y escribe excel/latinobarometro_panel.xlsx.
"""

import csv
from pathlib import Path

import openpyxl
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

RAIZ = Path(__file__).resolve().parent.parent
SALIDA = RAIZ / "excel" / "latinobarometro_panel.xlsx"

# nombre: archivo, categorías, categoría de interés, categorías válidas (denominador B)
VARIABLES = {
    "Apoyo democracia": dict(
        archivo="data/online/apoyo_democracia.xlsx",
        categorias=["Democracia", "Gobierno Autoritario", "Da lo mismo", "No sabe", "No contesta"],
        interes="Democracia",
        validas=["Democracia", "Gobierno Autoritario", "Da lo mismo"],
    ),
    "Grupos poderosos": dict(
        archivo="data/online/grupos_poderosos.xlsx",
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


def leer_exportacion(ruta: Path):
    """Devuelve [(país, año, categoría, proporción, n)] omitiendo '-' y el bloque Total."""
    filas = list(openpyxl.load_workbook(ruta).active.iter_rows(values_only=True))
    enc = next(i for i, f in enumerate(filas) if f[0] == "Restaurar vista")
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
                            datos.append((pais, int(anio), cat, float(v.rstrip("%")) / 100, int(n)))
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


def encabezado(ws, fila, textos, relleno=RELLENO_TITULO, fuente=F_TITULO):
    for j, t in enumerate(textos, 1):
        c = ws.cell(fila, j, t)
        c.font, c.fill = fuente, relleno
        c.alignment = Alignment(wrap_text=True, vertical="center")


def anchos(ws, anchos_col):
    for j, a in enumerate(anchos_col, 1):
        ws.column_dimensions[get_column_letter(j)].width = a


def main():
    datos = {v: leer_exportacion(RAIZ / cfg["archivo"]) for v, cfg in VARIABLES.items()}
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
    encabezado(wg, 1, ["País", "Año", "Gini (0-100)", "Fuente", "Clave"])
    for i, (p, a) in enumerate(panel, 2):
        wg.cell(i, 1, p).font = F_NORMAL
        wg.cell(i, 2, a).font = F_NORMAL
        for j in (3, 4):
            wg.cell(i, j).fill, wg.cell(i, j).font = RELLENO_INPUT, F_INPUT
        wg.cell(i, 5, f'=A{i}&"|"&B{i}').font = F_NORMAL
    wg["C1"].comment = Comment("Escriba el Gini del país-año en escala 0-100 (ej.: 42.4). "
                               "Deje vacío si no hay dato; esa fila quedará fuera de la regresión.", "Claude")
    wg.freeze_panes = "A2"
    anchos(wg, [16, 7, 13, 40, 22])
    ultima_gini = len(panel) + 1

    # --- Corrupcion: CPI de Transparency International -----------------------
    datahub, glob = leer_cpi()
    wk = wb.create_sheet("Corrupcion")
    encabezado(wk, 1, ["País", "Año", "CPI datahub (0-10, hasta 2011)", "CPI global HDX (0-100, desde 2012)",
                       "CPI escala 0-10 (empalmado)", "CPI 0-100 (solo desde 2012)", "Clave"])
    for i, (p, a) in enumerate(panel, 2):
        wk.cell(i, 1, p).font = F_NORMAL
        wk.cell(i, 2, a).font = F_NORMAL
        if (p, a) in datahub:
            wk.cell(i, 3, datahub[(p, a)])
        if a > ULTIMO_ANIO_DATAHUB and (p, a) in glob:
            wk.cell(i, 4, glob[(p, a)])
        for j in (3, 4):
            wk.cell(i, j).font = F_INPUT
        wk.cell(i, 5, f'=IF(B{i}<={ULTIMO_ANIO_DATAHUB},IF(C{i}="","",C{i}),IF(D{i}="","",D{i}/10))')
        wk.cell(i, 6, f'=IF(OR(B{i}<={ULTIMO_ANIO_DATAHUB},D{i}=""),"",D{i})')
        wk.cell(i, 7, f'=A{i}&"|"&B{i}')
        for j in (5, 6, 7):
            wk.cell(i, j).font = F_NORMAL
        wk.cell(i, 3).number_format = wk.cell(i, 5).number_format = "0.00"
        wk.cell(i, 4).number_format = wk.cell(i, 6).number_format = "0"
    wk["C1"].comment = Comment("Fuente: datahub.io/core/corruption-perceptions-index (Transparency "
                               "International). 0.0 en el archivo original = sin dato (queda vacío).", "Claude")
    wk["D1"].comment = Comment("Fuente: data.humdata.org/dataset/global-corruption-perceptions-index "
                               "(Transparency International), columna score.", "Claude")
    wk.freeze_panes = "C2"
    wk.auto_filter.ref = f"A1:G{len(panel) + 1}"
    anchos(wk, [16, 7, 16, 18, 16, 16, 22])
    wk.row_dimensions[1].height = 42
    ultima_cpi = len(panel) + 1

    # --- Panel_A y Panel_B: listos para regresión -----------------------------
    for criterio in ("A", "B"):
        wp = wb.create_sheet(f"Panel_{criterio}")
        otras = [v for v in VARIABLES if v != DEPENDIENTE]
        titulos = (["País", "Año", f"Y: {DEPENDIENTE} ({criterio})"]
                   + [f"X: {v} ({criterio})" for v in otras]
                   + ["X: Gini", "X: CPI escala 0-10 (empalmado)", "X: CPI 0-100 (desde 2012)",
                      "Completo con CPI 0-10 (1 = usar)", "Completo con CPI 0-100 (1 = usar)"])
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
            idx = f"INDEX(Gini!$C$2:$C${ultima_gini},MATCH($A{i}&\"|\"&$B{i},Gini!$E$2:$E${ultima_gini},0))"
            wp.cell(i, jg, f'=IFERROR(IF({idx}="","",{idx}),"")').font = F_LINK
            wp.cell(i, jg).number_format = "0.0"
            for k, col in ((1, "E"), (2, "F")):
                idx = (f"INDEX(Corrupcion!${col}$2:${col}${ultima_cpi},"
                       f"MATCH($A{i}&\"|\"&$B{i},Corrupcion!$G$2:$G${ultima_cpi},0))")
                c = wp.cell(i, jg + k, f'=IFERROR(IF({idx}="","",{idx}),"")')
                c.font, c.number_format = F_LINK, "0.00" if k == 1 else "0"
            base = f"COUNTBLANK(C{i}:{get_column_letter(jg)}{i})"
            for k in (1, 2):
                L = get_column_letter(jg + k)
                wp.cell(i, jg + 2 + k, f"=IF({base}+COUNTBLANK({L}{i})=0,1,0)").font = F_NORMAL
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
         "(porcentajes ya ponderados). Archivos: data/online/apoyo_democracia.xlsx y "
         "data/online/grupos_poderosos.xlsx.", False),
        ("", False),
        ("HOJAS", True),
        ("Datos: todas las tablas exportadas en formato largo (una fila por variable, país, año y categoría). "
         "Valores copiados de la herramienta online (texto azul).", False),
        ("Categorias: % de cada categoría por país-año, calculado con SUMIFS sobre Datos.", False),
        ("Gini: COMPLETAR las celdas amarillas (Gini 0-100 y fuente, p. ej. Banco Mundial / SEDLAC).", False),
        ("Corrupcion: Índice de Percepción de la Corrupción (Transparency International). Dos versiones: "
         "'CPI escala 0-10 (empalmado)' = datahub.io hasta 2011 + global HDX / 10 desde 2012; "
         "'CPI 0-100 (desde 2012)' = global HDX sin transformar.", False),
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
         "advierte que los puntajes anteriores no son comparables con los posteriores: la versión empalmada "
         "cruza ese quiebre; la versión 0-100 no.", False),
        ("", False),
        ("CÓMO CORRER LA REGRESIÓN", True),
        ("1. Complete la hoja Gini. 2. En Panel_A o Panel_B filtre la columna 'Completo' de la versión de CPI "
         "que va a usar (= 1). 3. Copie las filas visibles y péguelas como valores en una hoja nueva; "
         "elimine la columna de CPI que no usa para que las X queden contiguas. "
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
        ("2023: la herramienta online no entrega datos por país (solo el total regional), por eso no hay filas 2023.", False),
        ("España solo tiene 'Apoyo democracia' (no la pregunta de grupos poderosos): sus filas quedan con "
         "Completo = 0. 'Grupos poderosos' existe desde 2004 y falta en algunos país-año (p. ej. 2015).", False),
        ("CPI México 2015: global HDX (y el Excel oficial de TI) dice 31; datahub dice 3,5. Se usa HDX según "
         "la regla (desde 2012). Otros 10 países difieren en 2015, ninguno de Latinobarómetro.", False),
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
