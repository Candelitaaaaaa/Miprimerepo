declare const LIBRO_DEMO: string;
const wb = cargarLibro(LIBRO_DEMO);
const c = (id: string, de: string, asunto: string, extracto: string, extra: object = {}) =>
  JSON.stringify(Object.assign({ idMensaje: id, fecha: "2026-11-02T10:00:00", remitenteCorreo: de, asunto, extracto, direccion: "Recibido" }, extra));
let r = main(wbx(wb), c("a1", "talento@bhp.com", "Hola", "Texto cualquiera"));
igual([r.procesar, r.empresaPorDominio], [true, "BHP"], "dominio de empresa del catálogo");
verificar(r.contextoIA.indexOf("DEMO-0001 | [DEMO] Minera Los Andes | Memoria Planificación Mina 2027") >= 0, "contexto incluye postulaciones abiertas");
verificar(r.contextoIA.indexOf("DEMO-0002") < 0, "contexto excluye procesos finalizados");
r = main(wbx(wb), c("a2", "no-reply@myworkday.com", "Update", "Your profile"));
igual(r.procesar, true, "plataforma ATS relevante");
r = main(wbx(wb), c("a3", "jobs-noreply@linkedin.com", "Nuevos empleos para ti", "Mira estos empleos"));
igual(r.procesar, false, "alertas de LinkedIn sin palabras clave se ignoran");
r = main(wbx(wb), c("a4", "jobs-noreply@linkedin.com", "Tu solicitud se envió a Epiroc", "Tu solicitud se envió a Epiroc"));
igual(r.procesar, true, "confirmación de LinkedIn Easy Apply sí pasa");
r = main(wbx(wb), c("a5", "promo@tienda.cl", "50% de descuento", "Solo hoy"));
igual(r.procesar, false, "promoción se ignora");
r = main(wbx(wb), c("a6", "x@y.com", "Invitación a entrevista", "Te invitamos a una entrevista para la práctica"));
igual(r.procesar, true, "palabra clave");
r = main(wbx(wb), c("DEMO-MSG-001", "x@y.com", "a", "b"));
igual([r.procesar, r.duplicado], [false, true], "duplicado no se reprocesa");
r = main(wbx(wb), c("a7", "yo@alumnos.uai.cl", "CV", "Adjunto CV", { direccion: "Enviado", destinatarios: "rrhh@sqm.com" }));
igual(r.procesar, true, "correo enviado a dominio de empresa");
r = main(wbx(wb), c("a8", "x@y.com", "Tu cuenta", "Tu cuenta CVS Pharmacy")); // «cv» debe ser palabra completa
igual(r.procesar, false, "palabra clave como palabra completa");

// --- Clasificador por reglas (respaldo) contra los correos ficticios de prueba ---
{
  const fs = require("fs");
  const casos = JSON.parse(fs.readFileSync(require("path").join(LIBRO_DEMO, "..", "correos_prueba.json"), "utf-8"));
  let aciertos = 0;
  for (const c of casos) {
    const rr = clasificarPorReglas(normalizar(`${c.asunto}\n${c.texto}`));
    if (rr.tipo === c.esperado) aciertos++;
    else console.log(`  (reglas) ${c.id}: ${rr.tipo} ≠ ${c.esperado}`);
  }
  console.log(`  Reglas: ${aciertos}/${casos.length} correos ficticios clasificados igual que lo esperado`);
  verificar(aciertos >= casos.length - 1, "el respaldo por reglas acierta al menos n-1 casos");
}
finPruebas("EvaluarRelevancia + reglas");
