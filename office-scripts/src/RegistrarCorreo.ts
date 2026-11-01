/**
 * RegistrarCorreo
 * ---------------------------------------------------------------------------
 * Lo ejecuta Power Automate (acción «Ejecutar script» de Excel Online Business)
 * por cada correo que pasó el prefiltro.
 *
 * Parámetros (texto JSON):
 *   correoJson        datos del correo (ver interfaz Correo).
 *   clasificacionJson respuesta JSON de Claude. Si viene vacía o inválida, el
 *                     script clasifica con reglas (modo respaldo, más conservador).
 *
 * Garantías:
 *   - No duplica: si el «ID del mensaje» ya existe, no escribe nada.
 *   - No sobrescribe el historial: siempre agrega una fila nueva en Comunicaciones.
 *   - «Postulación enviada» solo con un correo ENVIADO por la usuaria.
 *   - Avance / Oferta / Rechazo se aplican solo con confianza ≥ UmbralConfianzaSensible
 *     y señales textuales explícitas; si no, van a Pendientes.
 *   - Nunca hace retroceder un proceso; cambios sobre procesos finalizados → Pendientes.
 */

interface Resultado {
  accion: string;          // registrado | pendiente | creado | duplicado | irrelevante | error
  idMensaje: string;
  idPostulacion: string;
  estadoAnterior: string;
  estadoNuevo: string;
  metodo: string;
  detalle: string;
}

function main(workbook: ExcelScript.Workbook, correoJson: string, clasificacionJson: string): Resultado {
  let correo: Correo;
  try {
    correo = JSON.parse(correoJson) as Correo;
  } catch (e) {
    return { accion: "error", idMensaje: "", idPostulacion: "", estadoAnterior: "", estadoNuevo: "", metodo: "",
      detalle: "correoJson no es JSON válido" };
  }
  const idMensaje = String(correo.idMensaje ?? "").trim();
  const base: Resultado = { accion: "error", idMensaje, idPostulacion: "", estadoAnterior: "", estadoNuevo: "", metodo: "", detalle: "" };
  if (!idMensaje) {
    base.detalle = "Falta idMensaje (use internetMessageId).";
    return base;
  }
  const ctx = cargarContexto(workbook);
  if (mensajeYaRegistrado(ctx, idMensaje)) {
    base.accion = "duplicado";
    base.detalle = "El mensaje ya estaba registrado; no se modificó nada.";
    return base;
  }

  const direccion = correo.direccion === "Enviado" ? "Enviado" : "Recibido";
  const textoNorm = normalizar(`${correo.asunto ?? ""}\n${correo.extracto ?? ""}`);
  const reglas = clasificarPorReglas(textoNorm);
  const motivos: string[] = [];

  // ---------------------------------------------------------- clasificación
  let cls: Clasificacion | null = null;
  let metodo = "Claude API";
  if (clasificacionJson && clasificacionJson.trim()) {
    try {
      cls = JSON.parse(clasificacionJson) as Clasificacion;
    } catch (e) {
      cls = null;
      motivos.push("La respuesta de la IA no era JSON válido.");
    }
  }
  if (cls && !ctx.tipos.has(String(cls.tipo_comunicacion))) {
    motivos.push(`Tipo de la IA no reconocido: «${cls.tipo_comunicacion}».`);
    cls = null;
  }
  if (!cls) {
    metodo = "Reglas";
    cls = {
      es_relevante: reglas.tipo !== "Correo irrelevante",
      empresa: null, cargo: null, tipo_comunicacion: reglas.tipo, etapa_proceso: null, fecha_evento: null,
      accion_pendiente: false, descripcion_accion: null, fecha_limite_accion: null, confianza: reglas.confianza,
      justificacion: reglas.evidencia ? `Coincidencia textual: «${reglas.evidencia}»` : "Sin patrones reconocidos.",
      evidencia_textual: reglas.evidencia, resumen: String(correo.asunto ?? ""),
      es_postulacion_enviada: direccion === "Enviado" && reglas.tipo === "Confirmación de postulación",
    };
    if (reglas.tipo === "Correo irrelevante") {
      // Pasó el prefiltro (dominio o palabra clave) pero no hay patrón claro: revisar.
      cls.confianza = 0.5;
    }
  }
  let confianza = Math.max(0, Math.min(1, numero(cls.confianza, 0)));
  const tipo = String(cls.tipo_comunicacion);
  const infoTipo = ctx.tipos.get(tipo) as InfoTipo;

  if (metodo === "Claude API" && infoTipo.sensible && reglas.encontrados.indexOf(tipo) < 0) {
    confianza = Math.min(confianza, ctx.umbralSensible - 0.01);
    motivos.push("Tipo sensible sin frase explícita reconocida por las reglas: requiere confirmación.");
  }

  if (tipo === "Correo irrelevante" || cls.es_relevante === false) {
    if (confianza >= ctx.umbralAuto) {
      base.accion = "irrelevante";
      base.metodo = metodo;
      base.detalle = "Clasificado como irrelevante con confianza suficiente; no se registra.";
      return base;
    }
    motivos.push("Posible correo irrelevante, pero con baja confianza.");
  }

  // --------------------------------------------------- empresa y postulación
  const empresa = resolverEmpresa(ctx, correo, cls.empresa, textoNorm);
  const cargo = String(cls.cargo ?? "").trim();
  const coincidencia = buscarPostulacion(ctx, empresa, cargo || null, String(correo.idConversacion ?? ""));
  const fechaSerial = isoASerial(correo.fecha) ?? ahoraSerial(correo.ahora);
  const hoy = ahoraSerial(correo.ahora);
  const enviado = direccion === "Enviado" && (cls.es_postulacion_enviada === true || tipo === "Confirmación de postulación");
  const estadoActual = coincidencia.fila >= 0 ? texto(ctx.post, coincidencia.fila, "Estado actual") : "";
  const propuesto = estadoPropuesto(ctx, tipo, cls.etapa_proceso, estadoActual, direccion, enviado);
  const umbral = infoTipo.sensible ? ctx.umbralSensible : ctx.umbralAuto;
  const contacto = direccion === "Recibido" && !esNoReply(correo.remitenteCorreo) ? correo.remitenteCorreo : "";
  const evidencia = enviado ? (correo.enlace ? correo.enlace : `Correo enviado ${serialATexto(fechaSerial)}: ${correo.asunto}`) : "";
  const accion = cls.accion_pendiente && cls.descripcion_accion ? String(cls.descripcion_accion) : "";

  let modo: "aplicar" | "crear" | "pendiente" = "pendiente";
  let estadoFinal = estadoActual;
  let aplicarEstado = false;

  if (motivos.some((m) => m.startsWith("Posible correo irrelevante"))) {
    modo = "pendiente";
  } else if (coincidencia.fila >= 0) {
    const ev = evaluarCambio(ctx, estadoActual, propuesto, direccion);
    if (ev.nota) motivos.push(ev.nota);
    if (ev.revisar) {
      modo = "pendiente";
    } else if (confianza < umbral && ev.aplicar) {
      motivos.push(`Confianza ${Math.round(confianza * 100)}% bajo el umbral ${Math.round(umbral * 100)}%.`);
      modo = "pendiente";
    } else if (confianza < ctx.umbralAuto) {
      motivos.push(`Confianza ${Math.round(confianza * 100)}% bajo el umbral ${Math.round(ctx.umbralAuto * 100)}%.`);
      modo = "pendiente";
    } else {
      modo = "aplicar";
      aplicarEstado = ev.aplicar;
      estadoFinal = ev.estadoFinal;
    }
  } else if (coincidencia.ambiguo) {
    motivos.push(`No se pudo elegir la postulación: ${coincidencia.motivo}.`);
  } else {
    const esConfirmacion = tipo === "Confirmación de postulación" || tipo === "Confirmación de recepción" || enviado;
    if (!empresa) {
      motivos.push("No se identificó la empresa.");
    } else if (!cargo) {
      motivos.push(`No existe postulación de ${empresa.nombre} y el correo no indica el cargo.`);
    } else if (!esConfirmacion) {
      motivos.push(`No existe postulación de ${empresa.nombre} para «${cargo}» y el correo no es una confirmación.`);
    } else if (!ctx.crearAuto) {
      motivos.push("CrearPostulacionesAutomaticas = No.");
    } else if (confianza < ctx.umbralAuto) {
      motivos.push(`Confianza ${Math.round(confianza * 100)}% bajo el umbral para crear una postulación.`);
    } else {
      modo = "crear";
    }
  }

  // ------------------------------------------------------------- escritura
  let idPost = "";
  let estadoAnterior = estadoActual;
  if (modo === "crear" && empresa) {
    asegurarEmpresa(ctx, empresa.nombre, dominioContraparte(correo));
    estadoAnterior = "";
    estadoFinal = enviado ? "Postulación enviada" : "Postulación recibida";
    const fila = crearPostulacion(ctx, empresa.nombre, cargo, estadoFinal, fechaSerial, hoy, evidencia, contacto,
      `Creada automáticamente desde el correo «${correo.asunto}» (${serialATexto(fechaSerial)}). ` +
      (enviado ? "Evidencia: correo enviado." : "La fecha de postulación es la del correo de confirmación."));
    idPost = texto(ctx.post, fila, "ID de postulación");
    actualizarPostulacion(ctx, fila, { fechaSerial, hoy, tipo, estadoNuevo: null, cargo, accion, fechaLimite: cls.fecha_limite_accion,
      contacto, evidenciaEnvio: evidencia });
    if (!accion) {
      escribir(ctx.post, fila, "Próxima acción", "Hacer seguimiento si no hay respuesta");
      escribir(ctx.post, fila, "Fecha de seguimiento", Math.floor(fechaSerial + ctx.diasSinRespuesta));
    }
  } else if (modo === "aplicar") {
    idPost = texto(ctx.post, coincidencia.fila, "ID de postulación");
    actualizarPostulacion(ctx, coincidencia.fila, {
      fechaSerial, hoy, tipo, estadoNuevo: aplicarEstado ? estadoFinal : null, cargo, accion,
      fechaLimite: cls.fecha_limite_accion, contacto, evidenciaEnvio: evidencia,
    });
  }

  const pendiente = modo === "pendiente";
  const justificacion = [cls.justificacion, cls.evidencia_textual ? `Evidencia: «${cls.evidencia_textual}»` : "",
    coincidencia.motivo ? `Vínculo: ${coincidencia.motivo}.` : "", empresa ? `Empresa por ${empresa.fuente}.` : ""]
    .concat(motivos).filter((x) => x).join(" ");

  agregarFila(ctx.com, new Map<string, Valor>([
    ["ID del mensaje", idMensaje],
    ["ID de postulación", pendiente ? "" : idPost],
    ["Empresa", empresa ? empresa.nombre : ""],
    ["Fecha de recepción", fechaSerial],
    ["Remitente", String(correo.remitenteNombre ?? "")],
    ["Correo del remitente", String(correo.remitenteCorreo ?? "")],
    ["Asunto", String(correo.asunto ?? "")],
    ["Tipo de comunicación", tipo],
    ["Estado anterior", estadoAnterior],
    ["Estado propuesto", propuesto],
    ["Estado confirmado", pendiente ? "" : estadoFinal],
    ["Resumen del mensaje", String(cls.resumen ?? "")],
    ["Enlace al correo original en Outlook", String(correo.enlace ?? "")],
    ["Nivel de confianza de clasificación", Math.round(confianza * 100) / 100],
    ["Confirmación manual", pendiente ? CONF_PEND : CONF_AUTO],
    ["Dirección", direccion],
    ["ID de conversación", String(correo.idConversacion ?? "")],
    ["Cargo detectado", cargo],
    ["Justificación de la clasificación", justificacion],
    ["Acción pendiente detectada", accion],
    ["Fecha del evento", String(cls.fecha_evento ?? "")],
    ["Método de clasificación", metodo],
    ["Fecha de registro", hoy],
  ]));

  if (pendiente) {
    agregarFila(ctx.pend, new Map<string, Valor>([
      ["ID del mensaje", idMensaje],
      ["Fecha de recepción", fechaSerial],
      ["Remitente", String(correo.remitenteNombre ?? "")],
      ["Correo del remitente", String(correo.remitenteCorreo ?? "")],
      ["Asunto", String(correo.asunto ?? "")],
      ["Resumen del mensaje", String(cls.resumen ?? "")],
      ["Enlace al correo", String(correo.enlace ?? "")],
      ["Empresa sugerida", empresa ? empresa.nombre : String(cls.empresa ?? "")],
      ["Cargo sugerido", cargo],
      ["Tipo sugerido", tipo],
      ["Estado sugerido", propuesto],
      ["ID de postulación sugerida", coincidencia.fila >= 0 ? texto(ctx.post, coincidencia.fila, "ID de postulación") :
        coincidencia.candidatos.join("; ")],
      ["Confianza", Math.round(confianza * 100) / 100],
      ["Motivo de revisión", motivos.join(" ") || "Revisión manual requerida."],
      ["Decisión", "Pendiente"],
      ["Procesado", "No"],
    ]));
  }

  return {
    accion: pendiente ? "pendiente" : modo === "crear" ? "creado" : "registrado",
    idMensaje,
    idPostulacion: pendiente ? "" : idPost,
    estadoAnterior,
    estadoNuevo: pendiente ? "" : estadoFinal,
    metodo,
    detalle: motivos.join(" ") || (coincidencia.motivo ?? ""),
  };
}
