# -*- coding: utf-8 -*-
"""
HTML interactivo autocontenido (sin dependencias externas) generado con los MISMOS datos del Excel maestro.
Los valores de score, clase, motivo, trazabilidad, controles, pendientes y diferencias se leen del Excel
recalculado; si no se pudo recalcular, se usan los valores equivalentes calculados en Python.
"""
from __future__ import annotations

import datetime as dt
import json
import math
from pathlib import Path

import openpyxl
import pandas as pd

import pk16_datos as D
import pk16_oficial as O

TPL = Path(__file__).resolve().parent / "plantilla_html.html"


def _v(x):
    if x is None:
        return None
    if isinstance(x, float):
        if math.isnan(x):
            return None
        return round(x, 4)
    if isinstance(x, (pd.Timestamp, dt.date, dt.datetime)):
        return str(x)[:10]
    if hasattr(x, "item"):
        return _v(x.item())
    return x


def _tabla(ws, header_row, first_col=1, last_col=None, stop_blank_col=1):
    """Lee una tabla de una hoja (valores en caché) a partir de su fila de cabecera."""
    last_col = last_col or ws.max_column
    heads = [ws.cell(header_row, c).value for c in range(first_col, last_col + 1)]
    rows = []
    r = header_row + 1
    while r <= ws.max_row:
        if ws.cell(r, stop_blank_col).value in (None, ""):
            break
        rows.append([_v(ws.cell(r, c).value) for c in range(first_col, last_col + 1)])
        r += 1
    return {"cols": heads, "rows": rows}


def construir_html(M, S, xlsx: Path, out: Path):
    wbv = openpyxl.load_workbook(xlsx, data_only=True)
    recalculado = wbv["04_FTE_CRITICIDAD"]["D11"].value is not None

    # --- Score por activo desde 03 (Excel recalculado) o desde Python
    sc = {}
    if recalculado:
        ws = wbv["03_SCORE_CRITICIDAD"]
        for r in range(6, ws.max_row + 1):
            i = ws.cell(r, 1).value
            if not i:
                continue
            g = lambda col: ws[f"{col}{r}"].value
            sc[i] = dict(c1=g("V"), c2=g("W"), c3=g("X"), c4=g("Y"), c5=g("Z"), sc=g("AA"), cs=g("AB"), cl=g("AE"),
                         pv=g("AF"), mo=g("AH"), pt=g("AI"), fte=g("AK"), su=g("AD"))
    else:
        for _, x in S.iterrows():
            sc[x["ID"]] = dict(c1=x["C1"], c2=x["C2"], c3=x["C3"], c4=x["C4"], c5=x["C5"], sc=x["SCORE"], cs=x["CLASE_SCORE"],
                               cl=x["CLASE"], pv=x["PROVISIONAL"], mo="", pt=x["PUNTOS"], fte=x["FTE"], su=x["DATOS_SUF"])

    assets = []
    for _, x in M["inv"].iterrows():
        s = sc.get(x["ID"], {})
        assets.append({
            "id": x["ID"], "eq": x["EQUIPO"], "inv": x["INVENTARIO"], "de": x["DESC"], "ub": x["UBIC"], "se": x["DESC_UBIC"],
            "tp": x["TP"], "ta": x["TIPO_ACTIVO"], "te": x["TECNOLOGIA"], "fa": x["FABRICANTE"], "mo": x["MODELO"],
            "an": _v(x["ANIO"]), "es": x["ESTADO"], "pq": 1 if x["EN_PARQUE"] == "S" else 0, "me": x["MOTIVO_ESTADO"],
            "pg": x["PROGRAMA"], "pr": x["PRODUCTO"], "li": x["LINEA"], "ps": x["PROCESO"], "gr": x["GRUPO_MANT"],
            "cb": x["CB"], "cbd": x["CB_DETALLE"], "mt": x["MONTAJE_PPAL"], "mb": x["MEB_CRIT"], "abc": x["SAP_ABC"],
            "avn": _v(x["AV_N"]), "avh": _v(x["AV_H"]), "avx": _v(x["AV_MAX"]), "ti": x["TIPOLOGIA"], "tn": x["TIPOLOGIA_NOMBRE"],
            "rf": x["REGLA_FORZADA"], "va": x["VALIDACION"], "fu": x["FUENTE"], "ho": x["HOJA"], "fi": _v(x["FILA"]),
            "dp": x["DEP_PREVIA"], "cpm": x["CRIT_PREV_MC"], "cpi": x["CRIT_PREV_INT"], "pa": x["EQUIPO_PADRE"],
            "c1": _v(s.get("c1")), "c2": _v(s.get("c2")), "c3": _v(s.get("c3")), "c4": _v(s.get("c4")), "c5": _v(s.get("c5")),
            "sc": _v(s.get("sc")), "cl": s.get("cl"), "pv": s.get("pv"), "mv": s.get("mo"), "pt": _v(s.get("pt")), "ft": _v(s.get("fte")),
        })

    # --- Averías por mes y top
    av = M["av"]
    avm = av.assign(MES=av["FECHA"].dt.strftime("%Y-%m")).groupby("MES").agg(n=("MIN", "size"), h=("MIN", lambda s: s.sum() / 60)).reset_index()

    # --- Tablas del Excel recalculado
    def tabla(sheet, hr, last_col=None, stop=1, first=1):
        return _tabla(wbv[sheet], hr, first, last_col, stop) if recalculado else {"cols": [], "rows": []}

    data = {
        "meta": {"generado": dt.datetime.now().strftime("%d/%m/%Y %H:%M"), "ppt": O.PPT, "recalculado": recalculado,
                 "xlsx": xlsx.name, "periodo_ini": "2025-01-03", "periodo_fin": str(av["FECHA"].max())[:10]},
        "of": {k: v[0] for k, v in O.OF.items()},
        "params": {"N_A": 418, "N_B": 275, "N_C": 1200, "N_D": 150, "RATIO_A": 12, "RATIO_B": 20, "RATIO_C": 35, "RATIO_D": 60,
                   "RMIN_A": 8, "RMIN_B": 12, "RMIN_C": 20, "RMIN_D": 35, "PT_A": 4, "PT_B": 2.5, "PT_C": 1.5, "PT_D": 1,
                   "PLANTILLA": 80, "HORAS_PROD": 1450, "FORM_H": 55, "PERS_SAB": 10, "SABADOS": 45, "FESTIVOS": 0, "JORNADAS": 220,
                   "PUNTOS_PM": 50000, "AVISOS": 10000, "PRODUCT": 6300, "COB_MIN": 0.15, "COB_MAX": 0.25, "COB_BASE": 1},
        "grupos_ppt": O.GRUPOS_S7, "secciones": O.SECCIONES_DIM, "familias": O.FAMILIAS_PPT,
        "assets": assets,
        "averias_mes": avm.to_dict("records"),
        "averias_total": {"n": int(len(av)), "h": round(float(av["MIN"].sum() / 60), 1)},
        "meb_of": [[_v(x) for x in r] for r in M["meb_of"][["Producto", "Proceso", "Tipo activo crítico", "Total crítico confirmado",
                                                               "Identificados en ficheros", "Pendientes inventario exacto"]].values.tolist()],
        "traz": tabla("11_TRAZABILIDAD", 5, 13),
        "dif": tabla("17_DIFERENCIAS", 5, 12),
        "pend": tabla("15_PENDIENTES_VALIDACION", 5, 9),
        "ctrl": tabla("16_CONTROLES", 10, 6),
        "fuentes": tabla("14_FUENTES", 5, 11),
        "tipologias": [list(t) for t in D.TIPOLOGIAS],
        "refs": O.REFERENCIAS_EXTERNAS,
    }
    html = TPL.read_text(encoding="utf-8").replace("/*__DATA__*/null", json.dumps(data, ensure_ascii=False, separators=(",", ":"), default=str).replace("</", "<\\/"))
    out.write_text(html, encoding="utf-8")
    print(f"    HTML: {out.stat().st_size / 1e6:.1f} MB · {len(assets)} registros")
