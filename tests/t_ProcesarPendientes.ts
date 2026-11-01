declare const LIBRO_DEMO: string;
const wb = cargarLibro(LIBRO_DEMO);
const pend = wb.tablas.get("tblPendientes")!;
const set = (id: string, col: string, v: string) => {
  const i = pend.d.rows.findIndex((r) => r[0] === id);
  pend.d.rows[i][pend.d.headers.indexOf(col)] = v;
};
set("DEMO-MSG-017", "ID de postulación confirmada", "DEMO-0001");
set("DEMO-MSG-017", "Decisión", "Registrar sin cambiar estado");
set("DEMO-MSG-018", "Empresa confirmada", "[DEMO] Empresa Desconocida");
set("DEMO-MSG-018", "Cargo confirmado", "Analista Junior");
set("DEMO-MSG-018", "Estado confirmado", "Evaluación online");
set("DEMO-MSG-018", "Decisión", "Crear nueva postulación");
const r = main(wbx(wb), "2026-11-02T09:00:00");
igual([r.procesados, r.errores], [2, 0], "procesa dos pendientes");
const p17 = wb.fila("tblPendientes", "ID del mensaje", "DEMO-MSG-017")!;
igual(p17["Procesado"], "Sí", "017 procesado");
const c17 = wb.fila("tblComunicaciones", "ID del mensaje", "DEMO-MSG-017")!;
igual([c17["ID de postulación"], c17["Estado confirmado"], c17["Confirmación manual"]],
  ["DEMO-0001", "Entrevista con jefatura", "Confirmada"], "017 vinculado sin cambiar estado");
igual(wb.fila("tblPostulaciones", "ID de postulación", "DEMO-0001")!["Estado actual"], "Entrevista con jefatura", "017 estado intacto");
const nueva = wb.fila("tblPostulaciones", "Empresa", "[DEMO] Empresa Desconocida")!;
igual([nueva["ID de postulación"], nueva["Estado actual"], nueva["Cargo o programa"]], ["POST-0016", "Evaluación online", "Analista Junior"], "018 crea postulación");
verificar(!!wb.fila("tblEmpresas", "Empresa", "[DEMO] Empresa Desconocida"), "018 agrega la empresa al catálogo");
const c18 = wb.fila("tblComunicaciones", "ID del mensaje", "DEMO-MSG-018")!;
igual([c18["ID de postulación"], c18["Confirmación manual"]], ["POST-0016", "Confirmada"], "018 historial vinculado");
// Segunda ejecución: no reprocesa
const r2 = main(wbx(wb), "2026-11-02T09:10:00");
igual(r2.procesados, 0, "idempotente");
// Error de validación: aplicar sin ID
pend.d.rows.push(pend.d.headers.map((h) => (h === "ID del mensaje" ? "DEMO-MSG-005" : h === "Decisión" ? "Aplicar" : h === "Tipo sugerido" ? "Invitación a entrevista" : "")));
pend.d.formulas.push(pend.d.headers.map(() => ""));
const r3 = main(wbx(wb), "2026-11-02T09:20:00");
igual(r3.errores, 1, "aplicar sin ID → error informado");
verificar(String(wb.fila("tblPendientes", "ID del mensaje", "DEMO-MSG-005")!["Resultado del procesamiento"]).startsWith("ERROR"), "mensaje de error visible");
finPruebas("ProcesarPendientes");
