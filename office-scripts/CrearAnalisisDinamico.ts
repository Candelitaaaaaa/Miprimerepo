// CrearAnalisisDinamico.ts — ARCHIVO GENERADO desde office-scripts/src/CrearAnalisisDinamico.ts.
// Copie TODO este archivo en Excel para la Web > Automatizar > Nuevo script.

/**
 * CrearAnalisisDinamico
 * ---------------------------------------------------------------------------
 * Ejecútelo UNA vez desde Excel para la Web (Automatizar > CrearAnalisisDinamico).
 * Crea la hoja «Análisis» con tablas dinámicas sobre tblPostulaciones y
 * tblComunicaciones, más segmentadores (Industria, Tipo de oportunidad,
 * Especialidad, Estado). Si ya existen, no las duplica: solo las refresca.
 *
 * Nota: los gráficos del Dashboard NO dependen de estas tablas dinámicas (usan
 * fórmulas que se actualizan solas). Esta hoja es para explorar con filtros.
 */

function main(workbook: ExcelScript.Workbook): string {
  const nombreHoja = "Análisis";
  const existente = workbook.getWorksheet(nombreHoja);
  const hoja: ExcelScript.Worksheet = existente ? existente : workbook.addWorksheet(nombreHoja);
  if (!existente) {
    hoja.getRange("A1").setValue("Análisis dinámico");
    hoja.getRange("A1").getFormat().getFont().setSize(16);
    hoja.getRange("A1").getFormat().getFont().setBold(true);
    hoja.getRange("A2").setValue("Tablas dinámicas y segmentadores. Se refrescan con el script ActualizarDashboard o Datos > Actualizar todo.");
  }
  const post = workbook.getTable("tblPostulaciones");
  const com = workbook.getTable("tblComunicaciones");
  if (!post || !com) {
    throw new Error("No se encontraron tblPostulaciones / tblComunicaciones.");
  }

  const crear = (nombre: string, origen: ExcelScript.Table, destino: string, filas: string[], columnas: string[], valor: string) => {
    const previa = workbook.getPivotTable(nombre);
    if (previa) return previa;
    const pt = hoja.addPivotTable(nombre, origen, destino);
    const jerarquia = (campo: string): ExcelScript.PivotHierarchy => {
      const h = pt.getHierarchy(campo);
      if (!h) throw new Error(`La tabla de origen no tiene la columna «${campo}».`);
      return h;
    };
    filas.forEach((f) => pt.addRowHierarchy(jerarquia(f)));
    columnas.forEach((c) => pt.addColumnHierarchy(jerarquia(c)));
    const dh = pt.addDataHierarchy(jerarquia(valor));
    dh.setSummarizeBy(ExcelScript.AggregationFunction.count);
    return pt;
  };

  const ptEstado = crear("pvEstado", post, "A6", ["Estado actual"], [], "ID de postulación");
  crear("pvEmpresaGrupo", post, "E6", ["Empresa"], ["Grupo de estado"], "ID de postulación");
  crear("pvIndustriaTipo", post, "A40", ["Industria"], ["Tipo de oportunidad"], "ID de postulación");
  crear("pvMensual", post, "A70", ["Mes de postulación"], ["Grupo de estado"], "ID de postulación");
  crear("pvComunicaciones", com, "E70", ["Tipo de comunicación"], ["Confirmación manual"], "ID del mensaje");

  // Segmentadores conectados a la tabla dinámica principal y a la de empresas
  const existentes = workbook.getSlicers().map((s) => s.getCaption());
  const seg = (fuente: ExcelScript.PivotTable, campo: string, titulo: string, izq: number, arriba: number) => {
    if (existentes.indexOf(titulo) >= 0) return;
    const s = workbook.addSlicer(fuente, campo, hoja);
    s.setCaption(titulo);
    s.setLeft(izq);
    s.setTop(arriba);
    s.setWidth(170);
    s.setHeight(180);
  };
  const ptEmp = workbook.getPivotTable("pvEmpresaGrupo") as ExcelScript.PivotTable;
  seg(ptEmp, "Industria", "Industria", 780, 20);
  seg(ptEmp, "Tipo de oportunidad", "Tipo de oportunidad", 960, 20);
  seg(ptEmp, "Especialidad", "Especialidad", 780, 210);
  seg(ptEstado, "Origen del registro", "Origen del registro", 960, 210);

  workbook.refreshAllPivotTables();
  return "Hoja «Análisis» lista: 5 tablas dinámicas y 4 segmentadores.";
}
