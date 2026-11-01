// Escenarios de RegistrarCorreo sobre el archivo de PRODUCCIÓN (15 registros «Por verificar»).
declare const LIBRO_PROD: string;
const wb = cargarLibro(LIBRO_PROD);

function correo(id: string, de: string, asunto: string, extracto: string, extra: Partial<Correo> = {}): string {
  return JSON.stringify(Object.assign({ idMensaje: id, idConversacion: "", fecha: "2026-11-02T10:00:00", ahora: "2026-11-02T10:05:00",
    remitenteNombre: "Reclutamiento", remitenteCorreo: de, asunto, extracto, enlace: `https://outlook.office.com/mail/deeplink/read/${id}`,
    direccion: "Recibido" }, extra));
}
function ia(tipo: string, conf: number, extra: Partial<Clasificacion> = {}): string {
  return JSON.stringify(Object.assign({ es_relevante: true, empresa: null, cargo: null, tipo_comunicacion: tipo, etapa_proceso: null,
    fecha_evento: null, accion_pendiente: false, descripcion_accion: null, fecha_limite_accion: null, confianza: conf,
    justificacion: "prueba", evidencia_textual: "", resumen: "resumen" }, extra));
}
const nCom = () => wb.filas("tblComunicaciones").filter((r) => r["ID del mensaje"] !== "").length;
const nPend = () => wb.filas("tblPendientes").filter((r) => r["ID del mensaje"] !== "").length;
const post = (empresa: string) => wb.filas("tblPostulaciones").filter((r) => r["Empresa"] === empresa);

// 1. Confirmación de recepción de BHP → vincula POST-0001 (cargo por verificar), estado Postulación recibida
let r = main(wbx(wb), correo("m1", "talento@bhp.com", "Hemos recibido tu postulación",
  "Hola, hemos recibido tu postulación a Práctica Profesional Verano 2027. Saludos."),
  ia("Confirmación de recepción", 0.96, { empresa: "BHP", cargo: "Práctica Profesional Verano 2027" }));
igual(r.accion, "registrado", "1 acción");
igual(r.idPostulacion, "POST-0001", "1 vincula BHP");
igual(r.estadoNuevo, "Postulación recibida", "1 estado");
let b = wb.fila("tblPostulaciones", "ID de postulación", "POST-0001")!;
igual(b["Cargo o programa"], "Práctica Profesional Verano 2027", "1 completa cargo");
igual(b["Tipo de oportunidad"], "Práctica profesional", "1 infiere tipo de oportunidad desde el cargo");
igual(b["Fecha de postulación"], 46328, "1 fecha de postulación = fecha del correo (02-11-2026)");
igual(b["Evidencia de envío"], "", "1 confirmación NO es evidencia de envío");
igual(b["Correo de contacto"], "talento@bhp.com", "1 contacto");
igual(nCom(), 1, "1 una comunicación (reutiliza la fila vacía)");
igual(wb.filas("tblComunicaciones").length, 1, "1 no agrega fila extra");

// 2. Duplicado
r = main(wbx(wb), correo("m1", "talento@bhp.com", "Hemos recibido tu postulación", "x"), ia("Confirmación de recepción", 0.96));
igual(r.accion, "duplicado", "2 duplicado");
igual(nCom(), 1, "2 no agrega filas");

// 3. Rechazo con alta confianza pero SIN frase explícita → pendiente
r = main(wbx(wb), correo("m3", "talento@bhp.com", "Novedades de tu proceso", "Te escribimos para darte novedades sobre tu proceso."),
  ia("Rechazo", 0.97, { cargo: "Práctica Profesional Verano 2027" }));
igual(r.accion, "pendiente", "3 rechazo sin señales → pendiente");
igual(wb.fila("tblPostulaciones", "ID de postulación", "POST-0001")!["Estado actual"], "Postulación recibida", "3 estado no cambia");
igual(nPend(), 1, "3 agrega a Pendientes");
igual(wb.fila("tblComunicaciones", "ID del mensaje", "m3")!["Confirmación manual"], "Pendiente", "3 historial marcado pendiente");

// 4. Rechazo explícito → Rechazada
r = main(wbx(wb), correo("m4", "talento@bhp.com", "Resultado de tu postulación",
  "Lamentamos informarte que no continuarás en el proceso de Práctica Profesional Verano 2027."),
  ia("Rechazo", 0.97, { cargo: "Práctica Profesional Verano 2027" }));
igual([r.accion, r.estadoNuevo], ["registrado", "Rechazada"], "4 rechazo explícito aplicado");
b = wb.fila("tblPostulaciones", "ID de postulación", "POST-0001")!;
igual([b["Próxima acción"], b["Fecha de seguimiento"]], ["", ""], "4 limpia próxima acción");

// 5. Invitación sobre proceso finalizado → pendiente
r = main(wbx(wb), correo("m5", "talento@bhp.com", "Invitación a entrevista", "Te invitamos a una entrevista."),
  ia("Invitación a entrevista", 0.95, { cargo: "Práctica Profesional Verano 2027" }));
igual(r.accion, "pendiente", "5 proceso finalizado → pendiente");

// 6. Confirmación Codelco: la IA sugiere «Postulación enviada» pero solo se permite «recibida»
r = main(wbx(wb), correo("m6", "no-reply@codelco.cl", "Gracias por postular", "Gracias por postular a Memoria Planificación Minera 2027.",
  { idConversacion: "conv-codelco-1" }),
  ia("Confirmación de postulación", 0.93, { empresa: "Codelco", cargo: "Memoria Planificación Minera 2027", etapa_proceso: "Postulación enviada" }));
igual([r.accion, r.idPostulacion, r.estadoNuevo], ["registrado", "POST-0009", "Postulación recibida"], "6 nunca «enviada» desde un correo recibido");
igual(wb.fila("tblPostulaciones", "ID de postulación", "POST-0009")!["Correo de contacto"], "", "6 no guarda no-reply como contacto");

// 7. Correo ENVIADO a Codelco con otra postulación → crea POST-0016 «Postulación enviada» con evidencia
r = main(wbx(wb), correo("m7", "yo@alumnos.uai.cl", "Postulación Práctica Profesional Verano", "Adjunto mi CV para la Práctica Profesional Verano.",
  { direccion: "Enviado", destinatarios: "practicas@codelco.cl", idConversacion: "conv-codelco-2" }),
  ia("Confirmación de postulación", 0.95, { empresa: "Codelco", cargo: "Práctica Profesional Verano", es_postulacion_enviada: true }));
igual([r.accion, r.idPostulacion, r.estadoNuevo], ["creado", "POST-0016", "Postulación enviada"], "7 crea postulación enviada");
b = wb.fila("tblPostulaciones", "ID de postulación", "POST-0016")!;
verificar(String(b["Evidencia de envío"]).startsWith("https://outlook.office.com"), "7 evidencia = enlace al correo enviado");
igual(b["Origen del registro"], "Automático (correo)", "7 origen");
igual(b["Industria"], "Minería", "7 industria desde catálogo");

// 8. Avance explícito sobre la memoria Codelco
r = main(wbx(wb), correo("m8", "seleccion@codelco.cl", "Siguiente etapa",
  "Has sido seleccionada para continuar a la siguiente etapa de Memoria Planificación Minera 2027."),
  ia("Avance en el proceso", 0.97, { cargo: "Memoria Planificación Minera 2027" }));
igual([r.idPostulacion, r.estadoNuevo], ["POST-0009", "En selección – etapa 1"], "8 avance aplicado");

// 9. Correo de revisión posterior: no retrocede
r = main(wbx(wb), correo("m9", "seleccion@codelco.cl", "Tu postulación está en revisión", "Tu postulación está siendo revisada."),
  ia("Revisión curricular", 0.95, { cargo: "Memoria Planificación Minera 2027" }));
igual([r.accion, r.estadoNuevo], ["registrado", "En selección – etapa 1"], "9 no retrocede");
igual(wb.fila("tblComunicaciones", "ID del mensaje", "m9")!["Estado propuesto"], "En revisión", "9 historial guarda lo propuesto");

// 10. Invitación a evaluación (movimiento dentro de selección)
r = main(wbx(wb), correo("m10", "seleccion@codelco.cl", "Evaluación online", "Te invitamos a rendir una evaluación online.",
  { idConversacion: "conv-codelco-1" }), ia("Invitación a evaluación", 0.9, { accion_pendiente: true,
  descripcion_accion: "Completar test online", fecha_limite_accion: "2026-11-06" }));
igual([r.idPostulacion, r.estadoNuevo], ["POST-0009", "Evaluación online"], "10 vincula por conversación y aplica");
b = wb.fila("tblPostulaciones", "ID de postulación", "POST-0009")!;
igual([b["Próxima acción"], b["Fecha de seguimiento"]], ["Completar test online", 46332], "10 próxima acción y fecha límite");

// 11. Modo reglas (sin IA): confirmación de Teck → Postulación recibida
r = main(wbx(wb), correo("m11", "careers@teck.com", "Gracias por postular", "Gracias por postular a Teck."), "");
igual([r.accion, r.metodo, r.idPostulacion, r.estadoNuevo], ["registrado", "Reglas", "POST-0005", "Postulación recibida"], "11 respaldo por reglas");

// 12. Modo reglas: avance → siempre pendiente (umbral sensible)
r = main(wbx(wb), correo("m12", "careers@teck.com", "Novedades", "Has sido seleccionada para continuar en el proceso."), "");
igual(r.accion, "pendiente", "12 reglas + tipo sensible → pendiente");

// 13. Empresa desconocida con baja confianza → pendiente
r = main(wbx(wb), correo("m13", "jobs@empresa-x.com", "Invitación a proceso", "Te invitamos a una entrevista."),
  ia("Invitación a entrevista", 0.6, { empresa: "Empresa X", cargo: "Analista" }));
igual(r.accion, "pendiente", "13 empresa desconocida → pendiente");

// 14. Irrelevante con alta confianza → no se registra
const antes = nCom();
r = main(wbx(wb), correo("m14", "news@portal.com", "Newsletter", "Ofertas de la semana"), ia("Correo irrelevante", 0.95, { es_relevante: false }));
igual([r.accion, nCom()], ["irrelevante", antes], "14 irrelevante no se registra");

// 15. JSON de IA inválido → respaldo por reglas
r = main(wbx(wb), correo("m15", "rrhh@sqm.com", "Hemos recibido tu postulación", "Hemos recibido tu postulación. Saludos."), "{no es json");
igual([r.metodo, r.idPostulacion, r.estadoNuevo], ["Reglas", "POST-0015", "Postulación recibida"], "15 IA inválida → reglas");

// 16. Asunto que parece fórmula se guarda como texto
r = main(wbx(wb), correo("m16", "rrhh@sqm.com", "=HYPERLINK(\"x\")", "Recordatorio de tu proceso."), ia("Seguimiento administrativo", 0.9));
igual(wb.fila("tblComunicaciones", "ID del mensaje", "m16")!["Asunto"], "=HYPERLINK(\"x\")", "16 asunto conservado (escrito como texto)");

// 17. Oferta explícita sobre Codelco memoria
r = main(wbx(wb), correo("m17", "seleccion@codelco.cl", "Carta oferta", "Nos complace ofrecerte la memoria. Adjuntamos carta oferta.",
  { idConversacion: "conv-codelco-1" }), ia("Oferta laboral", 0.97));
igual(r.estadoNuevo, "Oferta recibida", "17 oferta aplicada");

// 18. Correo sin cargo a Codelco: la memoria está finalizada (oferta) → vincula la única abierta (práctica)
r = main(wbx(wb), correo("m18", "practicas@codelco.cl", "Recordatorio", "Recordatorio: actualiza tus datos en el portal."),
  ia("Seguimiento administrativo", 0.9));
igual([r.idPostulacion, r.estadoNuevo], ["POST-0016", "Postulación enviada"], "18 única abierta, sin cambio de estado");

// 19. Gold Fields: confirmación de una empresa con registro «Por verificar» pero cargo distinto en otra postulación
r = main(wbx(wb), correo("m19", "careers@goldfields.com", "Application received",
  "We have received your application for Graduate Program 2027."),
  ia("Confirmación de recepción", 0.95, { cargo: "Graduate Program 2027" }));
igual([r.idPostulacion, r.estadoNuevo], ["POST-0003", "Postulación recibida"], "19 usa el registro inicial de la empresa");
igual(wb.fila("tblPostulaciones", "ID de postulación", "POST-0003")!["Tipo de oportunidad"], "Programa graduate", "19 tipo inferido");

// Integridad: nunca se escribió sobre columnas calculadas (el mock lanza error si ocurre) y no hay IDs repetidos
const ids = wb.filas("tblComunicaciones").map((x) => x["ID del mensaje"]).filter((x) => x !== "");
igual(ids.length, new Set(ids).size, "sin mensajes duplicados en el historial");
const idsP = wb.filas("tblPostulaciones").map((x) => x["ID de postulación"]);
igual(idsP.length, new Set(idsP).size, "sin IDs de postulación duplicados");
finPruebas("RegistrarCorreo");
