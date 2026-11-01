/**
 * LimpiarDatosFicticios
 * ---------------------------------------------------------------------------
 * Elimina todas las filas de prueba (ID que empieza con «DEMO-», empresas que
 * empiezan con «[DEMO]» u Origen = «Datos ficticios (prueba)»).
 * No toca ningún registro real.
 */

function main(workbook: ExcelScript.Workbook): string {
  const reglas: [string, (tb: Tabla, i: number) => boolean][] = [
    ["tblPostulaciones", (tb, i) => texto(tb, i, "ID de postulación").startsWith("DEMO-") ||
      texto(tb, i, "Origen del registro") === "Datos ficticios (prueba)" || texto(tb, i, "Empresa").startsWith("[DEMO]")],
    ["tblComunicaciones", (tb, i) => texto(tb, i, "ID del mensaje").startsWith("DEMO-") || texto(tb, i, "ID de postulación").startsWith("DEMO-")],
    ["tblPendientes", (tb, i) => texto(tb, i, "ID del mensaje").startsWith("DEMO-")],
    ["tblEmpresas", (tb, i) => texto(tb, i, "Empresa").startsWith("[DEMO]")],
  ];
  const resumen: string[] = [];
  for (const [nombre, esDemo] of reglas) {
    const tb = leerTabla(workbook, nombre);
    let borradas = 0;
    // de abajo hacia arriba para no desplazar índices
    for (let i = tb.v.length - 1; i >= 0; i--) {
      if (esDemo(tb, i)) {
        if (tb.t.getRowCount() > 1) {
          tb.t.deleteRowsAt(i, 1);
        } else {
          // Excel exige al menos una fila: se vacía en lugar de borrarla
          tb.h.forEach((_, c) => {
            if (!tb.f[c]) tb.t.getRangeBetweenHeaderAndTotal().getCell(i, c).setValue("");
          });
        }
        borradas++;
      }
    }
    resumen.push(`${nombre}: ${borradas}`);
  }
  return `Filas de prueba eliminadas → ${resumen.join(", ")}`;
}
