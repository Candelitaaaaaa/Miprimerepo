/**
 * EvaluarRelevancia
 * ---------------------------------------------------------------------------
 * Prefiltro barato que Power Automate ejecuta ANTES de llamar a la IA.
 * Devuelve procesar = true solo si el correo:
 *   - no está registrado aún (anti-duplicados), y
 *   - viene de (o va a) un dominio de una empresa del catálogo o de una
 *     plataforma de reclutamiento marcada como relevante, o
 *   - contiene alguna palabra clave de Catálogos > PalabrasClave.
 * Así se evita enviar a la IA (y pagar por) correos ajenos al proceso.
 */

interface Relevancia {
  procesar: boolean;
  duplicado: boolean;
  motivo: string;
  empresaPorDominio: string;
  contextoIA: string;   // empresas del catálogo y postulaciones abiertas, para el mensaje a la IA
}

function contextoParaIA(ctx: Contexto): string {
  const lineas: string[] = [];
  for (let i = 0; i < ctx.post.v.length && lineas.length < 60; i++) {
    const id = texto(ctx.post, i, "ID de postulación");
    const est = ctx.estados.get(texto(ctx.post, i, "Estado actual"));
    if (id && (!est || est.categoria !== "Final")) {
      lineas.push(`${id} | ${texto(ctx.post, i, "Empresa")} | ${texto(ctx.post, i, "Cargo o programa")} | ${texto(ctx.post, i, "Estado actual")}`);
    }
  }
  const empresas = ctx.empresas.map((e) => e.alias.length ? `${e.nombre} (${e.alias.join(", ")})` : e.nombre).join("; ");
  return `Empresas registradas: ${empresas}\nPostulaciones abiertas (ID | Empresa | Cargo | Estado):\n${lineas.join("\n")}`;
}

function main(workbook: ExcelScript.Workbook, correoJson: string): Relevancia {
  let correo: Correo;
  try {
    correo = JSON.parse(correoJson) as Correo;
  } catch (e) {
    return { procesar: false, duplicado: false, motivo: "correoJson inválido", empresaPorDominio: "", contextoIA: "" };
  }
  const ctx = cargarContexto(workbook);
  const id = String(correo.idMensaje ?? "").trim();
  if (!id) {
    return { procesar: false, duplicado: false, motivo: "Falta idMensaje", empresaPorDominio: "", contextoIA: "" };
  }
  if (mensajeYaRegistrado(ctx, id)) {
    return { procesar: false, duplicado: true, motivo: "Ya registrado", empresaPorDominio: "", contextoIA: "" };
  }
  const dom = dominioContraparte(correo);
  for (const e of ctx.empresas) {
    if (dom && e.dominios.some((d) => coincideDominio(dom, d))) {
      return { procesar: true, duplicado: false, motivo: `Dominio de ${e.nombre}`, empresaPorDominio: e.nombre, contextoIA: contextoParaIA(ctx) };
    }
  }
  const ats = ctx.ats.find((a) => coincideDominio(dom, a.dominio));
  if (ats && ats.relevante) {
    return { procesar: true, duplicado: false, motivo: `Plataforma de reclutamiento (${ats.dominio})`, empresaPorDominio: "", contextoIA: contextoParaIA(ctx) };
  }
  const textoNorm = normalizar(`${correo.asunto ?? ""}\n${correo.extracto ?? ""}`);
  const palabra = contienePalabraClave(ctx, textoNorm);
  if (palabra) {
    return { procesar: true, duplicado: false, motivo: `Palabra clave «${palabra}»`, empresaPorDominio: "", contextoIA: contextoParaIA(ctx) };
  }
  const reglas = clasificarPorReglas(textoNorm);
  if (reglas.tipo !== "Correo irrelevante") {
    return { procesar: true, duplicado: false, motivo: `Patrón «${reglas.evidencia}»`, empresaPorDominio: "", contextoIA: contextoParaIA(ctx) };
  }
  return { procesar: false, duplicado: false, motivo: "Sin señales de proceso de selección", empresaPorDominio: "", contextoIA: "" };
}
