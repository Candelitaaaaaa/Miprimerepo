/**
 * ProcesarPendientes
 * ---------------------------------------------------------------------------
 * Ejecútelo desde Excel para la Web (Automatizar > ProcesarPendientes) después
 * de completar las columnas amarillas de la hoja «Pendientes».
 *
 * Decisión:
 *   Aplicar                      vincula el correo a «ID de postulación confirmada»
 *                                y aplica «Estado confirmado» (si está vacío, no cambia el estado).
 *   Crear nueva postulación      crea la postulación con Empresa/Cargo confirmados
 *                                (o sugeridos) y el Estado confirmado.
 *   Registrar sin cambiar estado vincula el correo, sin cambiar el estado.
 *   Descartar (irrelevante)      marca el evento como Descartado en el historial.
 *
 * La confirmación manual es la autoridad: puede asignar cualquier estado,
 * incluido «Postulación enviada» si usted confirma que envió la postulación.
 */

interface ResumenPendientes { procesados: number; errores: number; detalle: string[]; }

function main(workbook: ExcelScript.Workbook, ahoraIso?: string): ResumenPendientes {
  const ctx = cargarContexto(workbook);
  const p = ctx.pend;
  const hoy = ahoraSerial(ahoraIso);
  const res: ResumenPendientes = { procesados: 0, errores: 0, detalle: [] };

  for (let i = 0; i < p.v.length; i++) {
    const idMsg = texto(p, i, "ID del mensaje");
    if (!idMsg || texto(p, i, "Procesado") === "Sí") continue;
    const decision = texto(p, i, "Decisión");
    if (!decision || decision === "Pendiente") continue;

    const fc = buscarFila(ctx.com, "ID del mensaje", idMsg);
    const fallo = (msg: string) => {
      escribir(p, i, "Resultado del procesamiento", `ERROR: ${msg}`);
      res.errores++;
      res.detalle.push(`${idMsg}: ${msg}`);
    };
    if (fc < 0) {
      fallo("no se encontró el mensaje en Comunicaciones.");
      continue;
    }
    const tipoConf = texto(p, i, "Tipo de comunicación confirmado") || texto(p, i, "Tipo sugerido");
    if (!ctx.tipos.has(tipoConf)) {
      fallo(`tipo de comunicación no válido: «${tipoConf}».`);
      continue;
    }
    const estadoConf = texto(p, i, "Estado confirmado");
    if (estadoConf && !ctx.estados.has(estadoConf)) {
      fallo(`estado no válido: «${estadoConf}».`);
      continue;
    }
    const fechaSerial = numero(leer(ctx.com, fc, "Fecha de recepción"), hoy);
    const correoRem = texto(ctx.com, fc, "Correo del remitente");
    const contacto = texto(ctx.com, fc, "Dirección") === "Recibido" && !esNoReply(correoRem) ? correoRem : "";

    if (decision === "Descartar (irrelevante)") {
      escribir(ctx.com, fc, "Confirmación manual", "Descartada");
      cerrar(p, i, hoy, "Descartado: no se modificó ninguna postulación.");
      res.procesados++;
      continue;
    }

    let filaPost = -1;
    let creado = false;
    if (decision === "Crear nueva postulación") {
      const empresa = texto(p, i, "Empresa confirmada") || texto(p, i, "Empresa sugerida");
      const cargo = texto(p, i, "Cargo confirmado") || texto(p, i, "Cargo sugerido");
      if (!empresa || !cargo) {
        fallo("para crear una postulación indique Empresa confirmada y Cargo confirmado.");
        continue;
      }
      if (!estadoConf) {
        fallo("para crear una postulación indique el Estado confirmado.");
        continue;
      }
      asegurarEmpresa(ctx, empresa, dominioDe(correoRem));
      filaPost = crearPostulacion(ctx, empresa, cargo, estadoConf, fechaSerial, hoy, "", contacto,
        `Creada por confirmación manual del correo ${idMsg}.`);
      creado = true;
    } else {
      const idConf = texto(p, i, "ID de postulación confirmada") ||
        (/^[A-Z]+-\d+$/.test(texto(p, i, "ID de postulación sugerida")) ? texto(p, i, "ID de postulación sugerida") : "");
      filaPost = buscarFila(ctx.post, "ID de postulación", idConf);
      if (filaPost < 0) {
        fallo(`indique un «ID de postulación confirmada» existente (recibido: «${idConf}»).`);
        continue;
      }
    }

    const idPost = texto(ctx.post, filaPost, "ID de postulación");
    const estadoAntes = creado ? "" : texto(ctx.post, filaPost, "Estado actual");
    const cambia = decision === "Aplicar" && estadoConf !== "" && estadoConf !== estadoAntes;
    const estadoFinal = creado ? estadoConf : cambia ? estadoConf : estadoAntes;
    const cargoConf = texto(p, i, "Cargo confirmado");
    actualizarPostulacion(ctx, filaPost, {
      fechaSerial, hoy, tipo: tipoConf, estadoNuevo: cambia ? estadoConf : null, cargo: cargoConf,
      accion: texto(ctx.com, fc, "Acción pendiente detectada"), fechaLimite: null, contacto, evidenciaEnvio: "",
    });

    const corregido = tipoConf !== texto(p, i, "Tipo sugerido") ||
      (estadoConf !== "" && estadoConf !== texto(p, i, "Estado sugerido")) ||
      (texto(p, i, "Empresa confirmada") !== "" && texto(p, i, "Empresa confirmada") !== texto(p, i, "Empresa sugerida"));
    // Se completan campos vacíos del evento (no se sobrescribe la clasificación original).
    escribir(ctx.com, fc, "ID de postulación", idPost);
    escribir(ctx.com, fc, "Empresa", texto(ctx.post, filaPost, "Empresa"));
    if (!texto(ctx.com, fc, "Estado anterior")) escribir(ctx.com, fc, "Estado anterior", estadoAntes);
    escribir(ctx.com, fc, "Estado confirmado", estadoFinal);
    escribir(ctx.com, fc, "Confirmación manual", corregido ? "Corregida" : "Confirmada");
    if (tipoConf !== texto(ctx.com, fc, "Tipo de comunicación")) {
      const j = texto(ctx.com, fc, "Justificación de la clasificación");
      escribir(ctx.com, fc, "Justificación de la clasificación",
        `${j} [Manual ${serialATexto(hoy)}: tipo corregido de «${texto(ctx.com, fc, "Tipo de comunicación")}» a «${tipoConf}».]`);
      escribir(ctx.com, fc, "Tipo de comunicación", tipoConf);
    }
    cerrar(p, i, hoy, creado ? `Creada ${idPost} con estado «${estadoFinal}».` :
      cambia ? `${idPost}: «${estadoAntes}» → «${estadoFinal}».` : `${idPost}: registrado sin cambio de estado.`);
    res.procesados++;
  }
  res.detalle.unshift(`Procesados: ${res.procesados}. Errores: ${res.errores}.`);
  console.log(res.detalle.join("\n"));
  return res;
}

function cerrar(p: Tabla, i: number, hoy: number, msg: string): void {
  escribir(p, i, "Procesado", "Sí");
  escribir(p, i, "Fecha de procesamiento", hoy);
  escribir(p, i, "Resultado del procesamiento", msg);
}
