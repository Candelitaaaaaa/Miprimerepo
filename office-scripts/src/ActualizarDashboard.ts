/**
 * ActualizarDashboard
 * ---------------------------------------------------------------------------
 * Puede ejecutarse a mano (Automatizar > ActualizarDashboard) o desde un flujo
 * programado diario de Power Automate. Hace lo siguiente:
 *   1. (Opcional, parámetro MarcarSinRespuestaAutomatico = Sí) cambia a
 *      «Sin respuesta» los procesos Enviada/En revisión que superan DiasSinRespuesta,
 *      dejando constancia en Observaciones.
 *   2. Ajusta los rangos de los gráficos del Dashboard al tamaño real de los
 *      catálogos (si agregó estados, industrias o especialidades).
 *   3. Recalcula el libro y refresca todas las tablas dinámicas (hoja Análisis).
 *   4. Escribe la fecha/hora de ejecución en el Dashboard.
 */

interface ResumenDashboard { marcadosSinRespuesta: number; graficosAjustados: number; tablasDinamicas: number; }

function main(workbook: ExcelScript.Workbook, ahoraIso?: string): ResumenDashboard {
  const ctx = cargarContexto(workbook);
  const hoy = ahoraSerial(ahoraIso);
  const res: ResumenDashboard = { marcadosSinRespuesta: 0, graficosAjustados: 0, tablasDinamicas: 0 };

  // 1. Sin respuesta (opcional)
  if (ctx.marcarSinRespuesta) {
    workbook.getApplication().calculate(ExcelScript.CalculationType.full);
    const post = leerTabla(workbook, "tblPostulaciones");
    for (let i = 0; i < post.v.length; i++) {
      const alerta = texto(post, i, "Alerta sin respuesta");
      const estado = texto(post, i, "Estado actual");
      const grupo = texto(post, i, "Grupo de estado");
      if (alerta === "Sí" && estado !== "Sin respuesta" && (grupo === "Enviada" || grupo === "En revisión")) {
        const dias = texto(post, i, "Días desde última comunicación");
        escribir(post, i, "Estado actual", "Sin respuesta");
        escribir(post, i, "Fecha de última actualización", Math.floor(hoy));
        const obs = texto(post, i, "Observaciones");
        escribir(post, i, "Observaciones", `${obs ? obs + " | " : ""}[${serialATexto(hoy)}] Marcado «Sin respuesta» ` +
          `automáticamente (${dias} días sin comunicación; estado previo «${estado}»).`);
        res.marcadosSinRespuesta++;
      }
    }
  }

  // 2. Rangos de gráficos según tamaño de catálogos
  const calc = workbook.getWorksheet("Cálculos");
  const dash = workbook.getWorksheet("Dashboard");
  if (calc && dash) {
    const n = (tabla: string) => Math.max(1, leerTabla(workbook, tabla).v.filter((r) => String(r[0] ?? "").trim() !== "").length);
    const nEst = Math.min(40, n("tblEstados"));
    const nInd = Math.min(40, n("tblIndustrias"));
    const nEsp = Math.min(20, n("tblEspecialidades"));
    const rangos: { [prefijo: string]: [string, string] } = {
      "1.": [`J5:J${4 + nEst}`, `K5:K${4 + nEst}`],
      "3.": [`U5:U${4 + nInd}`, `V5:V${4 + nInd}`],
      "5.": [`AB5:AB${4 + nEsp}`, `AC5:AC${4 + nEsp}`],
      "7.": [`AH5:AH${4 + nEst}`, `AI5:AI${4 + nEst}`],
    };
    for (const ch of dash.getCharts()) {
      let titulo = "";
      try {
        titulo = ch.getTitle().getText();
      } catch (e) {
        titulo = "";
      }
      const clave = Object.keys(rangos).find((k) => titulo.startsWith(k));
      if (!clave) continue;
      const serie = ch.getSeries()[0];
      if (!serie) continue;
      serie.setXAxisValues(calc.getRange(rangos[clave][0]));
      serie.setValues(calc.getRange(rangos[clave][1]));
      res.graficosAjustados++;
    }
    calc.getRange("G9").setValue(hoy);
  }

  // 3. Recalcular y refrescar tablas dinámicas
  workbook.getApplication().calculate(ExcelScript.CalculationType.full);
  const pivotes = workbook.getPivotTables();
  if (pivotes.length > 0) {
    workbook.refreshAllPivotTables();
  }
  res.tablasDinamicas = pivotes.length;
  console.log(JSON.stringify(res));
  return res;
}
