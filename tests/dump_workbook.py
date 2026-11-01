"""Exporta las tablas de un .xlsx a JSON para el libro simulado de las pruebas.

Valores: del archivo recalculado (LibreOffice). Fórmulas: del archivo original.
Las fechas se convierten a número de serie de Excel, como las ve Office Scripts.
Uso: python3 tests/dump_workbook.py original.xlsx recalculado.xlsx salida.json
"""
import datetime as dt
import json
import sys

from openpyxl import load_workbook

BASE = dt.datetime(1899, 12, 30)


def serial(v):
    if isinstance(v, dt.datetime):
        return round((v - BASE).total_seconds() / 86400, 5)
    if isinstance(v, dt.date):
        return (dt.datetime(v.year, v.month, v.day) - BASE).days
    return v


def main(orig, recalc, out):
    wf = load_workbook(orig)
    wv = load_workbook(recalc, data_only=True)
    tablas = {}
    for ws in wf.worksheets:
        for t in ws.tables.values():
            name = t.displayName
            cells = ws[t.ref]
            vcells = wv[ws.title][t.ref]
            headers = [c.value for c in cells[0]]
            rows, forms = [], []
            for rf, rv in zip(cells[1:], vcells[1:]):
                fr, vr = [], []
                for cf, cv in zip(rf, rv):
                    f = cf.value if isinstance(cf.value, str) and cf.value.startswith("=") else ""
                    fr.append(f)
                    v = serial(cv.value)
                    vr.append("" if v is None else v)
                rows.append(vr)
                forms.append(fr)
            tablas[name] = {"hoja": ws.title, "headers": headers, "rows": rows, "formulas": forms}
    json.dump({"tablas": tablas, "hojas": wf.sheetnames}, open(out, "w", encoding="utf-8"), ensure_ascii=False)


if __name__ == "__main__":
    main(*sys.argv[1:4])
