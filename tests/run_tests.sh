#!/usr/bin/env bash
# Ejecuta las pruebas de los Office Scripts con un libro simulado.
# Requisitos: python3 + openpyxl, LibreOffice (para recalcular), node y tsc.
set -euo pipefail
RAIZ="$(cd "$(dirname "$0")/.." && pwd)"
TMP="${TMPDIR:-/tmp}/pruebas_postulaciones"
rm -rf "$TMP" && mkdir -p "$TMP"
python3 "$RAIZ/src/build_workbook.py" "$TMP" >/dev/null
python3 "$RAIZ/tools/bundle_office_scripts.py" >/dev/null
RECALC="${RECALC:-}"
for f in Seguimiento_Postulaciones_2027 Seguimiento_Postulaciones_2027_DEMO; do
  cp "$TMP/$f.xlsx" "$TMP/$f.recalc.xlsx"
  if [ -n "$RECALC" ]; then python3 "$RECALC" "$TMP/$f.recalc.xlsx" 120 >/dev/null; \
  else soffice --headless --calc --convert-to xlsx --outdir "$TMP/lo" "$TMP/$f.xlsx" >/dev/null 2>&1 && cp "$TMP/lo/$f.xlsx" "$TMP/$f.recalc.xlsx"; fi
  python3 "$RAIZ/tests/dump_workbook.py" "$TMP/$f.xlsx" "$TMP/$f.recalc.xlsx" "$TMP/$f.json"
done
cp "$RAIZ/tests/correos_prueba.json" "$TMP/"
estado=0
for t in RegistrarCorreo EvaluarRelevancia ProcesarPendientes LimpiarDatosFicticios ActualizarDashboard; do
  {
    echo "const LIBRO_PROD = '$TMP/Seguimiento_Postulaciones_2027.json';"
    echo "const LIBRO_DEMO = '$TMP/Seguimiento_Postulaciones_2027_DEMO.json';"
    cat "$RAIZ/office-scripts/$t.ts" "$RAIZ/tests/mock.ts"
    grep -v '^declare const LIBRO_' "$RAIZ/tests/t_$t.ts"
  } > "$TMP/$t.ts"
  tsc --strict --target es2019 --lib es2019 --outDir "$TMP/js" "$RAIZ/tests/excelscript.d.ts" "$TMP/$t.ts" || { estado=1; continue; }
  node "$TMP/js/$t.js" || estado=1
done
# Verificación de tipos de los scripts sin pruebas de ejecución
tsc --noEmit --strict --target es2019 --lib es2019 "$RAIZ/tests/excelscript.d.ts" "$RAIZ/office-scripts/CrearAnalisisDinamico.ts" && echo "CrearAnalisisDinamico: tipos OK" || estado=1
exit $estado
