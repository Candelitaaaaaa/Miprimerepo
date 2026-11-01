// RegistrarCorreo.ts — ARCHIVO GENERADO desde office-scripts/src/RegistrarCorreo.ts + _comun.ts.
// Copie TODO este archivo en Excel para la Web > Automatizar > Nuevo script.

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

// ============================================================================
// BLOQUE COMÚN (se copia automáticamente en cada script por tools/bundle).
// Office Scripts no permite importar módulos: por eso cada script publicado en
// office-scripts/*.ts es autónomo. Edite este archivo y ejecute
//   python3 tools/bundle_office_scripts.py
// ============================================================================

type Valor = string | number | boolean;

const SIN_CAMBIO = "(sin cambio)";
const ORIGEN_AUTO = "Automático (correo)";
const CONF_AUTO = "No requerida (automática)";
const CONF_PEND = "Pendiente";

interface Tabla {
  nombre: string;
  t: ExcelScript.Table;
  h: string[];
  v: Valor[][];
  f: string[]; // fórmula de referencia por columna ("" si no es calculada)
}

interface Correo {
  idMensaje: string;
  idGraph?: string;
  idConversacion?: string;
  fecha: string;            // ISO local, p. ej. 2026-10-29T12:00:00
  ahora?: string;           // ISO local de la ejecución
  remitenteNombre?: string;
  remitenteCorreo: string;
  destinatarios?: string;   // correos separados por ;
  asunto: string;
  extracto?: string;        // texto plano (recortado)
  enlace?: string;
  direccion?: string;       // Recibido | Enviado
}

interface Clasificacion {
  es_relevante: boolean;
  empresa: string | null;
  cargo: string | null;
  tipo_comunicacion: string;
  etapa_proceso: string | null;
  fecha_evento: string | null;
  accion_pendiente: boolean;
  descripcion_accion: string | null;
  fecha_limite_accion: string | null;
  confianza: number;
  justificacion: string;
  evidencia_textual: string;
  resumen: string;
  es_postulacion_enviada?: boolean;
}

interface InfoEstado { orden: number; grupo: string; categoria: string; }
interface InfoTipo { defecto: string; permitidos: string[]; cambia: boolean; sensible: boolean; }
interface InfoEmpresa { nombre: string; industria: string; dominios: string[]; alias: string[]; fila: number; }
interface InfoAts { dominio: string; relevante: boolean; }

interface Contexto {
  wb: ExcelScript.Workbook;
  post: Tabla;
  com: Tabla;
  pend: Tabla;
  emp: Tabla;
  estados: Map<string, InfoEstado>;
  tipos: Map<string, InfoTipo>;
  empresas: InfoEmpresa[];
  ats: InfoAts[];
  palabras: string[];
  umbralAuto: number;
  umbralSensible: number;
  diasSeguimiento: number;
  diasSinRespuesta: number;
  crearAuto: boolean;
  marcarSinRespuesta: boolean;
}

// ---------------------------------------------------------------- tablas ---
function leerTabla(wb: ExcelScript.Workbook, nombre: string): Tabla {
  const t = wb.getTable(nombre);
  if (!t) {
    throw new Error(`No existe la tabla ${nombre}. Verifique que el archivo sea Seguimiento_Postulaciones_2027.xlsx.`);
  }
  const h = t.getHeaderRowRange().getValues()[0].map((x) => String(x));
  const n = t.getRowCount();
  let v: Valor[][] = [];
  let f: string[] = h.map(() => "");
  if (n > 0) {
    const cuerpo = t.getRangeBetweenHeaderAndTotal();
    v = cuerpo.getValues() as Valor[][];
    const fr = cuerpo.getRow(0).getFormulas()[0];
    f = fr.map((x) => (typeof x === "string" && x.startsWith("=") ? x : ""));
  }
  return { nombre, t, h, v, f };
}

function col(tb: Tabla, nombre: string): number {
  const i = tb.h.indexOf(nombre);
  if (i < 0) {
    throw new Error(`La tabla ${tb.nombre} no tiene la columna «${nombre}».`);
  }
  return i;
}

function leer(tb: Tabla, fila: number, nombre: string): Valor {
  return tb.v[fila][col(tb, nombre)];
}

function texto(tb: Tabla, fila: number, nombre: string): string {
  const x = leer(tb, fila, nombre);
  return x === null || x === undefined ? "" : String(x).trim();
}

function limpiarTexto(v: Valor): Valor {
  // Evita que un asunto como "=..." o "+56..." se interprete como fórmula.
  if (typeof v === "string" && /^[=+\-@]/.test(v)) {
    return "'" + v;
  }
  return v;
}

function escribir(tb: Tabla, fila: number, nombre: string, valor: Valor): void {
  const c = col(tb, nombre);
  if (tb.f[c]) {
    return; // nunca sobrescribir columnas calculadas
  }
  tb.v[fila][c] = valor;
  tb.t.getRangeBetweenHeaderAndTotal().getCell(fila, c).setValue(limpiarTexto(valor));
}

function filaVacia(tb: Tabla): number {
  for (let i = 0; i < tb.v.length; i++) {
    if (String(tb.v[i][0] ?? "").trim() === "") {
      return i;
    }
  }
  return -1;
}

/** Agrega una fila (reutiliza una fila vacía si existe). Devuelve el índice. */
function agregarFila(tb: Tabla, datos: Map<string, Valor>): number {
  datos.forEach((_, k) => col(tb, k)); // valida nombres
  const libre = filaVacia(tb);
  if (libre >= 0) {
    datos.forEach((valor, k) => escribir(tb, libre, k, valor));
    return libre;
  }
  const fila: Valor[] = tb.h.map((nombre, i) => {
    if (tb.f[i]) {
      return tb.f[i];
    }
    return datos.has(nombre) ? limpiarTexto(datos.get(nombre) as Valor) : "";
  });
  tb.t.addRow(-1, fila);
  tb.v.push(tb.h.map((nombre, i) => (tb.f[i] ? "" : datos.has(nombre) ? (datos.get(nombre) as Valor) : "")));
  return tb.v.length - 1;
}

function buscarFila(tb: Tabla, nombreCol: string, valor: string): number {
  const c = col(tb, nombreCol);
  const objetivo = valor.trim();
  if (!objetivo) {
    return -1;
  }
  for (let i = 0; i < tb.v.length; i++) {
    if (String(tb.v[i][c] ?? "").trim() === objetivo) {
      return i;
    }
  }
  return -1;
}

// ---------------------------------------------------------------- texto ----
function normalizar(s: string | null | undefined): string {
  return String(s ?? "")
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .toLowerCase()
    .replace(/[–—]/g, "-")
    .replace(/\s+/g, " ")
    .trim();
}

const VACIAS = new Set(["de", "del", "la", "el", "los", "las", "y", "en", "para", "por", "a", "al", "the", "of", "and",
  "for", "in", "to", "con", "un", "una", "2026", "2027", "2028"]);

function tokens(s: string): string[] {
  return normalizar(s)
    .replace(/[^a-z0-9 ]/g, " ")
    .split(" ")
    .filter((x) => x.length >= 3 && !VACIAS.has(x));
}

function similitud(a: string, b: string): number {
  const ta = new Set(tokens(a));
  const tb = new Set(tokens(b));
  if (ta.size === 0 || tb.size === 0) {
    return 0;
  }
  let inter = 0;
  ta.forEach((x) => {
    if (tb.has(x)) {
      inter++;
    }
  });
  return inter / Math.min(ta.size, tb.size);
}

function esPorVerificar(s: string): boolean {
  const n = normalizar(s);
  return n === "" || n === "por verificar";
}

function dominioDe(correo: string): string {
  const m = /@([a-z0-9.\-]+)/i.exec(correo || "");
  return m ? m[1].toLowerCase() : "";
}

function coincideDominio(dominio: string, base: string): boolean {
  const b = base.trim().toLowerCase();
  return b !== "" && (dominio === b || dominio.endsWith("." + b));
}

function esNoReply(correo: string): boolean {
  return /no-?reply|noreply|do-?not-?reply|notification|notificacion/i.test(correo || "");
}

// ---------------------------------------------------------------- fechas ---
const BASE_SERIAL = Date.UTC(1899, 11, 30);

function isoASerial(iso: string | null | undefined): number | null {
  const m = /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2}))?)?/.exec(String(iso ?? "").trim());
  if (!m) {
    return null;
  }
  const ms = Date.UTC(+m[1], +m[2] - 1, +m[3], +(m[4] ?? 0), +(m[5] ?? 0), +(m[6] ?? 0));
  return Math.round(((ms - BASE_SERIAL) / 86400000) * 100000) / 100000;
}

function ahoraSerial(iso?: string): number {
  const s = isoASerial(iso);
  if (s !== null) {
    return s;
  }
  const d = new Date(); // hora local del navegador o UTC en Power Automate
  return isoASerial(
    `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}T` +
      `${String(d.getHours()).padStart(2, "0")}:${String(d.getMinutes()).padStart(2, "0")}:00`
  ) as number;
}

function serialATexto(serial: number): string {
  const d = new Date(BASE_SERIAL + Math.floor(serial) * 86400000);
  return `${String(d.getUTCDate()).padStart(2, "0")}-${String(d.getUTCMonth() + 1).padStart(2, "0")}-${d.getUTCFullYear()}`;
}

// ------------------------------------------------------------- contexto ----
function parametro(wb: ExcelScript.Workbook, nombre: string, porDefecto: Valor): Valor {
  const tp = leerTabla(wb, "tblParametros");
  const i = buscarFila(tp, "Parámetro", nombre);
  if (i < 0) {
    return porDefecto;
  }
  const v = leer(tp, i, "Valor");
  return v === "" || v === null || v === undefined ? porDefecto : v;
}

function numero(v: Valor, d: number): number {
  const n = typeof v === "number" ? v : parseFloat(String(v).replace(",", "."));
  return isFinite(n) ? n : d;
}

function cargarContexto(wb: ExcelScript.Workbook): Contexto {
  const post = leerTabla(wb, "tblPostulaciones");
  const com = leerTabla(wb, "tblComunicaciones");
  const pend = leerTabla(wb, "tblPendientes");
  const emp = leerTabla(wb, "tblEmpresas");
  const te = leerTabla(wb, "tblEstados");
  const tt = leerTabla(wb, "tblTiposComunicacion");
  const ta = leerTabla(wb, "tblDominiosATS");
  const tk = leerTabla(wb, "tblPalabrasClave");

  const estados = new Map<string, InfoEstado>();
  te.v.forEach((_, i) => {
    const n = texto(te, i, "Estado");
    if (n) {
      estados.set(n, {
        orden: numero(leer(te, i, "Orden"), 0),
        grupo: texto(te, i, "Grupo"),
        categoria: texto(te, i, "Categoría de regla") || "Selección",
      });
    }
  });
  const tipos = new Map<string, InfoTipo>();
  tt.v.forEach((_, i) => {
    const n = texto(tt, i, "Tipo de comunicación");
    if (n) {
      tipos.set(n, {
        defecto: texto(tt, i, "Estado propuesto por defecto") || SIN_CAMBIO,
        permitidos: texto(tt, i, "Estados permitidos").split(";").map((x) => x.trim()).filter((x) => x),
        cambia: normalizar(texto(tt, i, "Cambia estado")) === "si",
        sensible: normalizar(texto(tt, i, "Sensible")) === "si",
      });
    }
  });
  const empresas: InfoEmpresa[] = [];
  emp.v.forEach((_, i) => {
    const n = texto(emp, i, "Empresa");
    if (n) {
      empresas.push({
        nombre: n,
        industria: texto(emp, i, "Industria"),
        dominios: texto(emp, i, "Dominios de correo").split(/[;,\s]+/).map((x) => x.trim().toLowerCase()).filter((x) => x),
        alias: texto(emp, i, "Alias o nombres alternativos").split(";").map((x) => x.trim()).filter((x) => x),
        fila: i,
      });
    }
  });
  const ats: InfoAts[] = [];
  ta.v.forEach((_, i) => {
    const d = texto(ta, i, "Dominio").toLowerCase();
    if (d) {
      ats.push({ dominio: d, relevante: normalizar(texto(ta, i, "Relevante por dominio")) !== "no" });
    }
  });
  const palabras = tk.v.map((r) => normalizar(String(r[0] ?? ""))).filter((x) => x);

  return {
    wb, post, com, pend, emp, estados, tipos, empresas, ats, palabras,
    umbralAuto: numero(parametro(wb, "UmbralConfianzaAutomatica", 0.85), 0.85),
    umbralSensible: numero(parametro(wb, "UmbralConfianzaSensible", 0.95), 0.95),
    diasSeguimiento: numero(parametro(wb, "DiasSeguimiento", 7), 7),
    diasSinRespuesta: numero(parametro(wb, "DiasSinRespuesta", 21), 21),
    crearAuto: normalizar(String(parametro(wb, "CrearPostulacionesAutomaticas", "Sí"))) === "si",
    marcarSinRespuesta: normalizar(String(parametro(wb, "MarcarSinRespuestaAutomatico", "No"))) === "si",
  };
}

function mensajeYaRegistrado(ctx: Contexto, idMensaje: string): boolean {
  return buscarFila(ctx.com, "ID del mensaje", idMensaje) >= 0 || buscarFila(ctx.pend, "ID del mensaje", idMensaje) >= 0;
}

// ------------------------------------------------- clasificador por reglas ---
// Se usa (a) como respaldo si la IA no responde y (b) para exigir señales
// textuales explícitas antes de aplicar automáticamente Avance/Oferta/Rechazo.
const PATRONES: [string, RegExp[]][] = [
  ["Rechazo", [
    /lamentamos (informar|comunicar|contarte)/, /no (has|ha) sido (seleccionad|elegid)/, /no (fue|fuiste|fuimos) (seleccionad|elegid)/,
    /no (continuaras|continuara|seguiras|avanzaras) (en|con) (el|este|nuestro) proceso/, /no (podremos|vamos a|podemos) (avanzar|continuar|seguir) con (tu|su)/,
    /(decidido|optado|decidimos) (por )?(continuar|avanzar|seguir) con otr[ao]s? (candidat|postulant|perfil)/, /otros candidatos cuyo perfil/,
    /\bunfortunately\b/, /not (to )?(move|moving|proceed|proceeding) forward/, /(have|has) not been selected/, /will not be (moving|progressing|proceeding)/,
    /no seguir adelante con (tu|su) postulacion/, /(tu|su) perfil no se ajusta/,
  ]],
  ["Oferta laboral", [
    /carta oferta/, /oferta (laboral|formal|de (trabajo|practica|memoria|pasantia))/, /nos (complace|alegra) (ofrecerte|ofrecerle|informarte que (has|fuiste) seleccionad)/,
    /(has|ha) sido seleccionad[ao] para (el cargo|la posicion|la vacante|el puesto|la practica|la memoria|la pasantia|el internship|el programa)/,
    /(fuiste|has sido) (elegid|seleccionad)[ao] para (realizar|desarrollar) (tu|la|su) (practica|memoria|pasantia)/,
    /job offer/, /offer letter/, /pleased to offer/, /delighted to offer/, /(we are|i am) happy to offer/,
  ]],
  ["Exámenes médicos", [/examen(es)? (pre-?)?ocupacional/, /examen(es)? medico/, /evaluacion (pre-?)?ocupacional/, /bateria de examenes/, /medical (exam|check|assessment)/]],
  ["Entrevista realizada", [
    /gracias por (tu|su) (participacion|tiempo) en (la|nuestra) entrevista/, /thank you for (interviewing|taking the time to interview|your time during the interview)/,
    /luego de (tu|su|la) entrevista/,
  ]],
  ["Invitación a entrevista", [
    /invit\w* a (una |la |tu |su )?(segunda |primera )?(video ?)?entrevista/, /(agendar|coordinar|programar|confirmar) (una |la |tu |su )?(video ?)?entrevista/,
    /entrevista (personal|por video|virtual|presencial|online|tecnica|con (el|la|tu) (jefatura|area|gerente|equipo|lider))/, /citad[ao] a (una )?entrevista/,
    /interview (invitation|invite)/, /(schedule|invite you to|invitation to) (an |a )?(video |phone )?interview/,
  ]],
  ["Evaluación realizada", [
    /(hemos recibido|recibimos) (tus|sus) (respuestas|resultados|evaluacion)/, /(completaste|has completado|finalizaste) (la|el|tu|su) (evaluacion|test|prueba|assessment)/,
    /thank you for completing/,
  ]],
  ["Invitación a evaluación", [
    /evaluacion (online|en linea|psicolaboral|psicometrica|tecnica|grupal)/, /\b(test|prueba)s? (online|en linea|psicometric|tecnic|de (ingles|razonamiento|personalidad|conocimiento))/,
    /\bassess?ment\b/, /dinamica grupal/, /(rendir|realizar|completar|responder) (el|la|una|un|los|las|tu|su) (test|evaluacion|prueba|cuestionario)/,
    /\b(pymetrics|hackerrank|codility|talentlens|cut-e)\b/, /online (test|assessment)/, /caso (tecnico|practico)/,
  ]],
  ["Solicitud de documentos", [
    /(enviar|adjuntar|remitir|necesitamos|solicitamos|requerimos|subir)[^.]{0,60}(certificado|documento|documentacion|concentracion de notas|cedula|carnet|antecedentes|cv actualizado|curriculum actualizado|ficha)/,
    /please (send|provide|upload)[^.]{0,40}(document|transcript|certificate)/,
  ]],
  ["Avance en el proceso", [
    /seleccionad[ao] para (continuar|avanzar|la siguiente|la proxima|seguir)/,
    /(avanzar|avanzas|avanzaste|pasas|pasaste|has pasado|has avanzado|continuas|quedaste) (a|en) (la )?(siguiente|proxima|nueva|segunda|tercera) etapa/,
    /continuas en (el|nuestro) proceso/, /sigues en (el|nuestro) proceso/, /\b(etapa final|ultima etapa)\b/, /\bpreseleccionad/,
    /(move|moving|moved) (you )?forward (to|in)/, /next (stage|step|round) of (the|our) (process|recruitment|selection)/, /have been shortlisted/,
  ]],
  ["Revisión curricular", [
    /(esta|se encuentra|estan) (siendo )?(en )?revision/, /estamos (revisando|evaluando|analizando) (tu|su|los|tus|sus) (postulacion|perfil|antecedentes|cv|curriculum)/,
    /\b(under|in) review\b/, /being reviewed/, /reviewing your application/,
  ]],
  ["Confirmación de recepción", [
    /(hemos recibido|recibimos|acusamos recibo de) (correctamente )?(tu|su|la|el|tus|sus) (postulacion|cv|curriculum|solicitud|antecedentes|candidatura)/,
    /we (have )?received your (application|cv|resume)/, /application received/, /acuso recibo/,
  ]],
  ["Confirmación de postulación", [
    /gracias por (postular|tu postulacion|su postulacion|tu interes en|aplicar)/,
    /(tu|su) (postulacion|solicitud) (ha sido|fue|quedo|se ha) (enviada|registrada|completada|ingresada|exitosa|recibida con exito|enviado)/,
    /postulacion exitosa/, /thank you for (applying|your application)/, /application (has been |was )?(submitted|completed|sent)/, /successfully (applied|submitted)/,
    /(tu|su) solicitud se envio/, /(adjunto|envio|remito) (mi |el |la )?(cv|curriculum|carta)[^.]{0,80}postul/,
  ]],
  ["Seguimiento administrativo", [
    /\brecordatorio\b/, /\breminder\b/, /(reprogramar|reprogramacion|cambio de (fecha|horario))/, /actualiza(r)? (tus|sus) (datos|perfil)/,
    /proceso (se encuentra )?(en pausa|suspendido|pospuesto)/, /confirma(r)? (tu|su) asistencia/,
  ]],
];

const CONF_BASE: { [k: string]: number } = {
  "Confirmación de postulación": 0.88,
  "Confirmación de recepción": 0.88,
  "Revisión curricular": 0.87,
  "Seguimiento administrativo": 0.8,
};

interface ResultadoReglas { tipo: string; encontrados: string[]; confianza: number; evidencia: string; }

function clasificarPorReglas(textoNorm: string): ResultadoReglas {
  const encontrados: string[] = [];
  let evidencia = "";
  for (const [tipo, regs] of PATRONES) {
    for (const r of regs) {
      const m = r.exec(textoNorm);
      if (m) {
        encontrados.push(tipo);
        if (!evidencia) {
          evidencia = m[0];
        }
        break;
      }
    }
  }
  if (encontrados.length === 0) {
    return { tipo: "Correo irrelevante", encontrados, confianza: 0, evidencia: "" };
  }
  const tipo = encontrados[0];
  const conf = encontrados.length === 1 ? CONF_BASE[tipo] ?? 0.8 : 0.7;
  return { tipo, encontrados, confianza: conf, evidencia };
}

function contienePalabraClave(ctx: Contexto, textoNorm: string): string {
  for (const p of ctx.palabras) {
    const esc = p.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    if (new RegExp(`(^|[^a-z0-9])${esc}([^a-z0-9]|$)`).test(textoNorm)) {
      return p;
    }
  }
  return "";
}

// ------------------------------------------------------- empresa/match ----
interface EmpresaResuelta { nombre: string; enCatalogo: boolean; fuente: string; }

function esAts(ctx: Contexto, dominio: string): boolean {
  return ctx.ats.some((a) => coincideDominio(dominio, a.dominio));
}

function dominioContraparte(c: Correo): string {
  if (c.direccion === "Enviado") {
    const primero = String(c.destinatarios ?? "").split(/[;,]/)[0] ?? "";
    return dominioDe(primero);
  }
  return dominioDe(c.remitenteCorreo);
}

function empresaPorNombre(ctx: Contexto, nombre: string): InfoEmpresa | null {
  const n = normalizar(nombre);
  if (n.length < 2) {
    return null;
  }
  for (const e of ctx.empresas) {
    const nombres = [e.nombre].concat(e.alias).map((x) => normalizar(x));
    if (nombres.some((x) => x === n)) {
      return e;
    }
  }
  for (const e of ctx.empresas) {
    const nombres = [e.nombre].concat(e.alias).map((x) => normalizar(x)).filter((x) => x.length >= 3);
    if (nombres.some((x) => n.includes(x) || (n.length >= 4 && x.includes(n)))) {
      return e;
    }
  }
  return null;
}

function resolverEmpresa(ctx: Contexto, c: Correo, empresaIA: string | null, textoNorm: string): EmpresaResuelta | null {
  const dom = dominioContraparte(c);
  if (dom && !esAts(ctx, dom)) {
    for (const e of ctx.empresas) {
      if (e.dominios.some((d) => coincideDominio(dom, d))) {
        return { nombre: e.nombre, enCatalogo: true, fuente: `dominio ${dom}` };
      }
    }
  }
  if (empresaIA && empresaIA.trim()) {
    const e = empresaPorNombre(ctx, empresaIA);
    if (e) {
      return { nombre: e.nombre, enCatalogo: true, fuente: "nombre detectado" };
    }
    return { nombre: empresaIA.trim(), enCatalogo: false, fuente: "nombre detectado (no está en catálogo)" };
  }
  // Último recurso: nombre o alias del catálogo mencionado en el texto
  const hallados = ctx.empresas.filter((e) =>
    [e.nombre].concat(e.alias).some((x) => {
      const nx = normalizar(x).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      return nx.length >= 3 && new RegExp(`(^|[^a-z0-9])${nx}([^a-z0-9]|$)`).test(textoNorm);
    })
  );
  if (hallados.length === 1) {
    return { nombre: hallados[0].nombre, enCatalogo: true, fuente: "mencionada en el texto" };
  }
  return null;
}

interface Coincidencia { fila: number; ambiguo: boolean; motivo: string; candidatos: string[]; }

function buscarPostulacion(ctx: Contexto, empresa: EmpresaResuelta | null, cargo: string | null, idConv: string): Coincidencia {
  // 1) Misma conversación ya vinculada
  if (idConv) {
    const cConv = col(ctx.com, "ID de conversación");
    const cId = col(ctx.com, "ID de postulación");
    const cConf = col(ctx.com, "Confirmación manual");
    for (let i = ctx.com.v.length - 1; i >= 0; i--) {
      const r = ctx.com.v[i];
      if (String(r[cConv] ?? "") === idConv && String(r[cId] ?? "").trim() && String(r[cConf] ?? "") !== "Descartada") {
        const f = buscarFila(ctx.post, "ID de postulación", String(r[cId]));
        if (f >= 0) {
          return { fila: f, ambiguo: false, motivo: "misma conversación de correo", candidatos: [String(r[cId])] };
        }
      }
    }
  }
  if (!empresa) {
    return { fila: -1, ambiguo: false, motivo: "empresa no identificada", candidatos: [] };
  }
  const ne = normalizar(empresa.nombre);
  const cands: number[] = [];
  ctx.post.v.forEach((_, i) => {
    if (normalizar(texto(ctx.post, i, "Empresa")) === ne && texto(ctx.post, i, "ID de postulación")) {
      cands.push(i);
    }
  });
  const ids = cands.map((i) => texto(ctx.post, i, "ID de postulación"));
  if (cands.length === 0) {
    return { fila: -1, ambiguo: false, motivo: "sin postulaciones registradas para la empresa", candidatos: [] };
  }
  const abiertos = cands.filter((i) => {
    const e = ctx.estados.get(texto(ctx.post, i, "Estado actual"));
    return !e || e.categoria !== "Final";
  });
  if (cargo && cargo.trim()) {
    const puntajes = cands
      .map((i) => ({ i, s: similitud(cargo, texto(ctx.post, i, "Cargo o programa")) }))
      .sort((a, b) => b.s - a.s);
    const mejor = puntajes[0];
    const segundo = puntajes.length > 1 ? puntajes[1].s : 0;
    if (mejor.s >= 0.5 && mejor.s - segundo >= 0.15) {
      return { fila: mejor.i, ambiguo: false, motivo: `cargo coincide (${Math.round(mejor.s * 100)}%)`, candidatos: ids };
    }
    if (mejor.s >= 0.5) {
      return { fila: -1, ambiguo: true, motivo: "varias postulaciones con cargo similar", candidatos: ids };
    }
    const sinCargo = cands.filter((i) => esPorVerificar(texto(ctx.post, i, "Cargo o programa")));
    if (sinCargo.length === 1 && abiertos.length <= 1) {
      return { fila: sinCargo[0], ambiguo: false, motivo: "registro de la empresa sin cargo definido", candidatos: ids };
    }
    return { fila: -1, ambiguo: false, motivo: "el cargo no coincide con postulaciones existentes", candidatos: ids };
  }
  if (abiertos.length === 1) {
    return { fila: abiertos[0], ambiguo: false, motivo: "única postulación abierta en la empresa", candidatos: ids };
  }
  return { fila: -1, ambiguo: true, motivo: abiertos.length > 1 ? "varias postulaciones abiertas y el correo no indica el cargo" :
    "solo hay procesos finalizados en la empresa", candidatos: ids };
}

// ---------------------------------------------------------- estados -------
function estadoPropuesto(ctx: Contexto, tipo: string, etapaIA: string | null, actual: string, direccion: string, enviado: boolean): string {
  if (direccion === "Enviado") {
    return enviado ? "Postulación enviada" : SIN_CAMBIO;
  }
  const t = ctx.tipos.get(tipo);
  if (!t || !t.cambia) {
    return SIN_CAMBIO;
  }
  if (etapaIA && t.permitidos.indexOf(etapaIA) >= 0) {
    return etapaIA;
  }
  if (tipo === "Avance en el proceso" && actual === "En selección – etapa 1" && t.permitidos.indexOf("En selección – etapa 2") >= 0) {
    return "En selección – etapa 2";
  }
  return t.defecto;
}

interface Evaluacion { aplicar: boolean; estadoFinal: string; revisar: boolean; nota: string; }

/** Reglas de transición. No decide por confianza (eso se hace aparte). */
function evaluarCambio(ctx: Contexto, actual: string, propuesto: string, direccion: string): Evaluacion {
  if (propuesto === SIN_CAMBIO || propuesto === actual) {
    return { aplicar: false, estadoFinal: actual, revisar: false, nota: "" };
  }
  if (propuesto === "Postulación enviada" && direccion !== "Enviado") {
    return { aplicar: false, estadoFinal: actual, revisar: true, nota: "«Postulación enviada» requiere evidencia de envío." };
  }
  const ea = ctx.estados.get(actual);
  const ep = ctx.estados.get(propuesto);
  if (!ep) {
    return { aplicar: false, estadoFinal: actual, revisar: true, nota: `Estado «${propuesto}» no existe en Catálogos.` };
  }
  if (!ea || ea.categoria === "Especial" || actual === "") {
    return { aplicar: true, estadoFinal: propuesto, revisar: false, nota: "" };
  }
  if (ea.categoria === "Final") {
    return { aplicar: false, estadoFinal: actual, revisar: true, nota: `El proceso ya estaba en «${actual}»; un cambio requiere confirmación manual.` };
  }
  if (ep.categoria === "Final" || (ep.categoria === "Selección" && ea.categoria === "Selección")) {
    return { aplicar: true, estadoFinal: propuesto, revisar: false, nota: "" };
  }
  if (ep.orden < ea.orden) {
    return { aplicar: false, estadoFinal: actual, revisar: false, nota: `No se retrocede de «${actual}» a «${propuesto}»; se registra solo en el historial.` };
  }
  return { aplicar: true, estadoFinal: propuesto, revisar: false, nota: "" };
}

// ------------------------------------------------------ postulaciones ------
function siguienteId(ctx: Contexto): string {
  let max = 0;
  ctx.post.v.forEach((_, i) => {
    const m = /^POST-(\d+)$/.exec(texto(ctx.post, i, "ID de postulación"));
    if (m) {
      max = Math.max(max, parseInt(m[1], 10));
    }
  });
  return `POST-${String(max + 1).padStart(4, "0")}`;
}

function inferirTipoOportunidad(cargo: string): string {
  const n = normalizar(cargo);
  if (/memori|tesis|thesis/.test(n)) return "Memoria de título";
  if (/practic/.test(n)) return "Práctica profesional";
  if (/pasant/.test(n)) return "Pasantía";
  if (/graduate|trainee/.test(n)) return "Programa graduate";
  if (/intern/.test(n)) return "Internship";
  return "Por verificar";
}

function asegurarEmpresa(ctx: Contexto, nombre: string, dominio: string): void {
  if (ctx.empresas.some((e) => normalizar(e.nombre) === normalizar(nombre))) {
    return;
  }
  const generico = /(gmail|outlook|hotmail|yahoo|live|icloud)\./.test(dominio);
  const datos = new Map<string, Valor>([
    ["Empresa", nombre],
    ["Industria", ""],
    ["Dominios de correo", dominio && !generico && !esAts(ctx, dominio) ? dominio : ""],
    ["Notas", "Agregada automáticamente desde un correo: completar industria y dominios."],
  ]);
  const fila = agregarFila(ctx.emp, datos);
  ctx.empresas.push({ nombre, industria: "", dominios: [], alias: [], fila });
}

function industriaDe(ctx: Contexto, empresa: string): string {
  const e = ctx.empresas.find((x) => normalizar(x.nombre) === normalizar(empresa));
  return e ? e.industria : "";
}

function crearPostulacion(ctx: Contexto, empresa: string, cargo: string, estado: string, fechaSerial: number,
  hoy: number, evidencia: string, contacto: string, obs: string): number {
  const id = siguienteId(ctx);
  const datos = new Map<string, Valor>([
    ["ID de postulación", id],
    ["Empresa", empresa],
    ["Cargo o programa", cargo || "Por verificar"],
    ["Industria", industriaDe(ctx, empresa)],
    ["Especialidad", "Por verificar"],
    ["Ubicación", "Por verificar"],
    ["Tipo de oportunidad", inferirTipoOportunidad(cargo)],
    ["Fecha de postulación", Math.floor(fechaSerial)],
    ["Estado actual", estado],
    ["Fecha de última actualización", Math.floor(hoy)],
    ["Evidencia de envío", evidencia],
    ["Correo de contacto", contacto],
    ["Observaciones", obs],
    ["Origen del registro", ORIGEN_AUTO],
  ]);
  return agregarFila(ctx.post, datos);
}

interface DatosEvento {
  fechaSerial: number;
  hoy: number;
  tipo: string;
  estadoNuevo: string | null; // null = no cambia
  cargo: string;
  accion: string;
  fechaLimite: string | null;
  contacto: string;
  evidenciaEnvio: string;
}

function actualizarPostulacion(ctx: Contexto, fila: number, d: DatosEvento): void {
  const p = ctx.post;
  if (d.estadoNuevo) {
    escribir(p, fila, "Estado actual", d.estadoNuevo);
  }
  escribir(p, fila, "Fecha de última actualización", Math.floor(d.hoy));
  escribir(p, fila, "Última comunicación", `${serialATexto(d.fechaSerial)} · ${d.tipo}`);
  if (d.cargo && esPorVerificar(texto(p, fila, "Cargo o programa"))) {
    escribir(p, fila, "Cargo o programa", d.cargo);
  }
  if (d.cargo && esPorVerificar(texto(p, fila, "Tipo de oportunidad"))) {
    const t = inferirTipoOportunidad(d.cargo);
    if (t !== "Por verificar") {
      escribir(p, fila, "Tipo de oportunidad", t);
    }
  }
  if (d.estadoNuevo && ["Postulación enviada", "Postulación recibida"].indexOf(d.estadoNuevo) >= 0 &&
    typeof leer(p, fila, "Fecha de postulación") !== "number") {
    escribir(p, fila, "Fecha de postulación", Math.floor(d.fechaSerial));
  }
  if (d.evidenciaEnvio && !texto(p, fila, "Evidencia de envío")) {
    escribir(p, fila, "Evidencia de envío", d.evidenciaEnvio);
  }
  if (d.contacto && !texto(p, fila, "Correo de contacto")) {
    escribir(p, fila, "Correo de contacto", d.contacto);
  }
  if (d.accion) {
    escribir(p, fila, "Próxima acción", d.accion);
    const lim = isoASerial(d.fechaLimite);
    escribir(p, fila, "Fecha de seguimiento", Math.floor(lim !== null ? lim : d.fechaSerial + ctx.diasSeguimiento));
  } else if (d.estadoNuevo) {
    const e = ctx.estados.get(d.estadoNuevo);
    if (e && e.categoria === "Final" && d.estadoNuevo !== "Oferta recibida") {
      escribir(p, fila, "Próxima acción", "");
      escribir(p, fila, "Fecha de seguimiento", "");
    } else if (e && e.categoria === "Postulación") {
      escribir(p, fila, "Próxima acción", "Hacer seguimiento si no hay respuesta");
      escribir(p, fila, "Fecha de seguimiento", Math.floor(d.fechaSerial + ctx.diasSinRespuesta));
    }
  }
}
// ============================ FIN BLOQUE COMÚN ==============================
