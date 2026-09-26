"""Prueba de punta a punta con dos olas sintéticas de formato Latinobarómetro."""

import sys
import zipfile
from pathlib import Path

import pandas as pd
import pyreadstat

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))
import consolidar  # noqa: E402

CONFIG = Path(__file__).resolve().parent.parent / "config"


def ola(ruta, anio, var_apoyo, etiqueta_apoyo):
    df = pd.DataFrame({
        "IDENPA": [32, 76, 152, 862],
        "NUMINVES": [anio] * 4,
        "WT": [1.0, 0.9, 1.1, 1.0],
        "SEXO": [1, 2, 2, 1],
        "EDAD": [30, 45, 61, 22],
        var_apoyo: [1, 2, -1, 3],
        "P99X": [5, 6, 7, 8],
    })
    pyreadstat.write_sav(
        df, str(ruta),
        column_labels=["País", "Año", "Ponderador", "Sexo", "Edad",
                       etiqueta_apoyo, "Otra pregunta"],
        variable_value_labels={var_apoyo: {1.0: "La democracia es preferible", 2.0: "Autoritario",
                                           3.0: "Da lo mismo", -1.0: "No sabe"}},
    )


def test_consolidacion(tmp_path):
    raw, out = tmp_path / "raw", tmp_path / "out"
    raw.mkdir()
    ola(raw / "Latinobarometro_2018.sav", 2018, "P12STGBS", "P12STGBS Apoyo a la democracia")
    ola(tmp_path / "F00011.sav", 2020, "p10stgbs", "Apoyo a la Democracia")
    with zipfile.ZipFile(raw / "Latinobarometro_2020_Esp_Spss.zip", "w") as zf:
        zf.write(tmp_path / "F00011.sav", "F00011.sav")

    assert consolidar.main(["--entrada", str(raw), "--salida", str(out),
                            "--config", str(CONFIG), "--todas", "--stata"]) == 0

    base = pd.read_parquet(out / "latinobarometro_consolidado.parquet")
    assert len(base) == 8
    assert sorted(base.anio.unique()) == [2018, 2020]
    assert set(base.pais) == {"Argentina", "Brasil", "Chile", "Venezuela"}
    # El código -1 (no sabe) pasa a NaN pero se conserva en el texto.
    assert base.apoyo_democracia.isna().sum() == 2
    assert (base.apoyo_democracia_txt == "No sabe").sum() == 2
    assert (base.apoyo_democracia_txt == "La democracia es preferible").sum() == 2
    assert base.edad.tolist()[:4] == [30, 45, 61, 22]

    rep = pd.read_csv(out / "reporte_mapeo.csv")
    fila = rep[(rep.anio == 2018) & (rep.concepto == "apoyo_democracia")].iloc[0]
    assert fila.variable == "p12stgbs" and fila.origen == "auto"
    assert (out / "latinobarometro_todas_variables.parquet").exists()
    assert (out / "latinobarometro_consolidado.dta").exists()
    assert len(pd.read_csv(out / "inventario_variables.csv")) == 14


def test_mapeo_manual_tiene_prioridad(tmp_path):
    raw, out, cfg = tmp_path / "raw", tmp_path / "out", tmp_path / "cfg"
    raw.mkdir(); cfg.mkdir()
    ola(raw / "lb_2018.sav", 2018, "P12STGBS", "Apoyo a la democracia")
    for f in ("conceptos.csv", "paises.csv"):
        (cfg / f).write_text((CONFIG / f).read_text())
    (cfg / "mapeo_manual.csv").write_text("anio,concepto,variable\n2018,apoyo_democracia,P99X\n")
    consolidar.main(["--entrada", str(raw), "--salida", str(out), "--config", str(cfg)])
    base = pd.read_parquet(out / "latinobarometro_consolidado.parquet")
    assert base.apoyo_democracia.tolist() == [5, 6, 7, 8]
