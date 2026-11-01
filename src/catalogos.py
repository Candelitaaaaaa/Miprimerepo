"""Catálogos y datos semilla del sistema de seguimiento de postulaciones.

Este archivo es la única fuente de verdad para:
  * estados del proceso y sus reglas,
  * tipos de comunicación y el estado que proponen,
  * parámetros de confianza,
  * empresas iniciales (registros "Por verificar"),
  * datos ficticios de prueba (solo para el archivo DEMO).

Lo usa src/build_workbook.py para generar el Excel.
"""
import datetime as dt

# ---------------------------------------------------------------------------
# Estados (orden = avance; categoría = cómo la trata el motor de reglas)
#   Categoría de regla:
#     Inicial      -> antes de postular (se puede avanzar libremente)
#     Postulación  -> enviada / recibida / en revisión
#     Selección    -> etapas activas; se permite moverse entre ellas
#     Final        -> oferta, seleccionada, rechazada, cerrado: cualquier
#                     cambio posterior requiere aprobación manual
#     Especial     -> Por verificar / Sin respuesta: cualquier evidencia lo reemplaza
# ---------------------------------------------------------------------------
ESTADOS = [
    # Estado, Orden, Grupo (dashboard), Categoría de regla, Activa, Cuenta como postulada, Descripción
    ("Oportunidad identificada", 1, "Pre-postulación", "Inicial", "No", "No",
     "Oferta detectada; aún no se prepara la postulación."),
    ("En preparación", 2, "Pre-postulación", "Inicial", "No", "No",
     "CV, carta o formulario en preparación."),
    ("Postulación enviada", 3, "Enviada", "Postulación", "Sí", "Sí",
     "Existe evidencia de envío (correo enviado, comprobante o confirmación manual)."),
    ("Postulación recibida", 4, "En revisión", "Postulación", "Sí", "Sí",
     "La empresa o su portal confirmó la recepción. No implica selección."),
    ("En revisión", 5, "En revisión", "Postulación", "Sí", "Sí",
     "La empresa indicó que revisa los antecedentes."),
    ("En selección – etapa 1", 6, "Selección", "Selección", "Sí", "Sí",
     "Se informó explícitamente que avanzó a una etapa de selección."),
    ("En selección – etapa 2", 7, "Selección", "Selección", "Sí", "Sí",
     "Segunda etapa de selección informada explícitamente."),
    ("Evaluación online", 8, "Evaluación", "Selección", "Sí", "Sí",
     "Invitación a test, assessment o evaluación en línea."),
    ("Entrevista con reclutamiento", 9, "Entrevista", "Selección", "Sí", "Sí",
     "Entrevista con RR.HH. o reclutamiento."),
    ("Entrevista con jefatura", 10, "Entrevista", "Selección", "Sí", "Sí",
     "Entrevista con el área o jefatura."),
    ("Evaluación técnica o grupal", 11, "Evaluación", "Selección", "Sí", "Sí",
     "Caso técnico, assessment center o dinámica grupal."),
    ("Exámenes médicos", 12, "Evaluación", "Selección", "Sí", "Sí",
     "Exámenes pre-ocupacionales solicitados."),
    ("Etapa final", 13, "Selección", "Selección", "Sí", "Sí",
     "La empresa indicó que es la última etapa."),
    ("Oferta recibida", 14, "Oferta", "Final", "Sí", "Sí",
     "Oferta o carta oferta recibida; pendiente de respuesta."),
    ("Seleccionada", 15, "Oferta", "Final", "No", "Sí",
     "Selección confirmada para el cargo/programa."),
    ("Rechazada", 90, "Rechazo", "Final", "No", "Sí",
     "La empresa informó que no continúa el proceso."),
    ("Sin respuesta", 50, "Sin respuesta", "Especial", "Sí", "Sí",
     "Sin comunicación durante el plazo definido en Parámetros."),
    ("Proceso cerrado", 95, "Cerrado", "Final", "No", "Sí",
     "Cerrado por decisión propia, desistimiento o fin del programa."),
    ("Por verificar", 0, "Por verificar", "Especial", "No", "No",
     "Registro sin evidencia suficiente; revisar manualmente."),
]

SIN_CAMBIO = "(sin cambio)"

# Tipo, Estado propuesto por defecto, Estados permitidos (;), Cambia estado, Sensible, Descripción
TIPOS_COMUNICACION = [
    ("Confirmación de postulación", "Postulación recibida", "Postulación recibida", "Sí", "No",
     "Portal o empresa confirma que la postulación quedó registrada: «Gracias por postular», «Tu postulación fue enviada con éxito»."),
    ("Confirmación de recepción", "Postulación recibida", "Postulación recibida", "Sí", "No",
     "«Hemos recibido tu postulación/CV». Solo acusa recibo; NO implica selección."),
    ("Revisión curricular", "En revisión", "En revisión", "Sí", "No",
     "«Tu postulación está siendo revisada». No implica avance."),
    ("Avance en el proceso", "En selección – etapa 1",
     "En selección – etapa 1;En selección – etapa 2;Etapa final", "Sí", "Sí",
     "«Has sido seleccionada para continuar / pasas a la siguiente etapa». Requiere alta certeza."),
    ("Invitación a evaluación", "Evaluación online",
     "Evaluación online;Evaluación técnica o grupal", "Sí", "No",
     "Invitación a test, assessment, caso o dinámica. No implica haber aprobado etapas previas."),
    ("Evaluación realizada", SIN_CAMBIO, "", "No", "No",
     "Confirmación de que se completó una evaluación. No cambia el estado."),
    ("Invitación a entrevista", "Entrevista con reclutamiento",
     "Entrevista con reclutamiento;Entrevista con jefatura", "Sí", "No",
     "«Te invitamos a una entrevista», coordinación de horario de entrevista."),
    ("Entrevista realizada", SIN_CAMBIO, "", "No", "No",
     "Agradecimiento o confirmación posterior a una entrevista. No cambia el estado."),
    ("Solicitud de documentos", SIN_CAMBIO, "", "No", "No",
     "Piden certificados, notas, cédula, etc. Genera una próxima acción."),
    ("Exámenes médicos", "Exámenes médicos", "Exámenes médicos", "Sí", "No",
     "Citación a exámenes pre-ocupacionales."),
    ("Oferta laboral", "Oferta recibida", "Oferta recibida;Seleccionada", "Sí", "Sí",
     "«Has sido seleccionada para el cargo», carta oferta. Requiere alta certeza."),
    ("Rechazo", "Rechazada", "Rechazada", "Sí", "Sí",
     "«Lamentamos informarte que no continuarás». Requiere alta certeza."),
    ("Seguimiento administrativo", SIN_CAMBIO, "", "No", "No",
     "Recordatorios, reprogramaciones, actualización de datos, procesos en pausa."),
    ("Correo irrelevante", SIN_CAMBIO, "", "No", "No",
     "Newsletters, ofertas masivas, alertas de empleo sin postulación asociada."),
]

PARAMETROS = [
    ("UmbralConfianzaAutomatica", 0.85,
     "Confianza mínima (0-1) para aplicar automáticamente un tipo NO sensible."),
    ("UmbralConfianzaSensible", 0.95,
     "Confianza mínima para aplicar automáticamente Avance, Oferta o Rechazo (además deben coincidir las señales textuales)."),
    ("DiasSinRespuesta", 21,
     "Días sin comunicación (estado Enviada / En revisión) para alertar «sin respuesta»."),
    ("DiasSeguimiento", 7,
     "Días después de un correo para proponer la fecha de seguimiento si el correo no indica otra."),
    ("CrearPostulacionesAutomaticas", "Sí",
     "Sí = una confirmación clara de una empresa/cargo nuevo crea la postulación. No = va a Pendientes."),
    ("MarcarSinRespuestaAutomatico", "No",
     "Sí = el script ActualizarDashboard cambia el estado a «Sin respuesta» al cumplirse el plazo."),
    ("MesInicioDashboard", dt.date(2026, 7, 1),
     "Primer mes del gráfico de evolución mensual (muestra 12 meses)."),
    ("ModeloIA", "claude-opus-5-5",
     "Informativo: modelo usado en el flujo de Power Automate."),
    ("VersionEsquema", "1.0", "Versión de la estructura del archivo (no modificar)."),
]

INDUSTRIAS = [
    "Minería", "Minería no metálica y litio", "Proveedores de equipos mineros",
    "Servicios mineros", "Forestal y celulosa", "Tecnología industrial",
    "Energía", "Manufactura", "Consultoría", "Otra",
]
ESPECIALIDADES = [
    "Ingeniería Civil Industrial", "Ingeniería Civil en Minas",
    "Doble titulación (Industrial + Minas)", "Por verificar",
]
TIPOS_OPORTUNIDAD = [
    "Memoria de título", "Práctica profesional", "Pasantía", "Internship",
    "Programa graduate", "Por verificar",
]

# Empresa, Industria, Dominios de correo (;), Alias (;), Prioridad, Notas
# Los dominios son referenciales: confirmar con el primer correo real recibido.
EMPRESAS = [
    ("BHP", "Minería", "bhp.com", "BHP Billiton;Minera Escondida;Escondida;Spence", "Alta", ""),
    ("Collahuasi", "Minería", "collahuasi.cl", "Compañía Minera Doña Inés de Collahuasi;CMDIC", "Alta", ""),
    ("Gold Fields", "Minería", "goldfields.com", "Salares Norte;Gold Fields Salares Norte", "Alta", ""),
    ("CMPC", "Forestal y celulosa", "cmpc.com;cmpc.cl", "Empresas CMPC;Softys", "Media", ""),
    ("Teck", "Minería", "teck.com", "Teck Resources;Quebrada Blanca;Carmen de Andacollo", "Alta", ""),
    ("Epiroc", "Proveedores de equipos mineros", "epiroc.com", "", "Media", ""),
    ("Caserones", "Minería", "lundinmining.com", "Minera Lumina Copper;Lundin Mining;SCM Minera Lumina Copper Chile", "Media", "Verificar dominio de correo."),
    ("Hexagon", "Tecnología industrial", "hexagon.com", "Hexagon Mining", "Media", ""),
    ("Codelco", "Minería", "codelco.cl", "Corporación Nacional del Cobre", "Alta", ""),
    ("Sandvik", "Proveedores de equipos mineros", "sandvik.com", "Sandvik Mining and Rock Solutions", "Media", ""),
    ("Komatsu", "Proveedores de equipos mineros", "komatsu.com", "Komatsu Chile;Komatsu Cummins;KCH", "Media", "Verificar dominio de correo."),
    ("ENAEX", "Servicios mineros", "enaex.com", "Enaex", "Media", ""),
    ("AMSA", "Minería", "aminerals.cl", "Antofagasta Minerals;Antofagasta plc;Los Pelambres;Centinela;Zaldívar", "Alta", ""),
    ("Rio Tinto", "Minería", "riotinto.com", "", "Alta", ""),
    ("SQM", "Minería no metálica y litio", "sqm.com", "Sociedad Química y Minera", "Alta", ""),
]

# Plataformas de reclutamiento: su dominio NO identifica a la empresa.
DOMINIOS_ATS = [
    # Dominio, Plataforma, Relevante por dominio (Sí = todo correo de este dominio pasa a la IA)
    ("myworkday.com", "Workday", "Sí"), ("myworkdayjobs.com", "Workday", "Sí"),
    ("successfactors.com", "SAP SuccessFactors", "Sí"), ("successfactors.eu", "SAP SuccessFactors", "Sí"),
    ("sapsf.com", "SAP SuccessFactors", "Sí"), ("sapsf.eu", "SAP SuccessFactors", "Sí"),
    ("taleo.net", "Oracle Taleo", "Sí"), ("oraclecloud.com", "Oracle Recruiting", "Sí"),
    ("greenhouse.io", "Greenhouse", "Sí"), ("lever.co", "Lever", "Sí"),
    ("smartrecruiters.com", "SmartRecruiters", "Sí"), ("icims.com", "iCIMS", "Sí"),
    ("teamtailor.com", "Teamtailor", "Sí"), ("hiringroom.com", "HiringRoom", "Sí"),
    ("trabajando.com", "Trabajando.com", "Sí"), ("laborum.cl", "Laborum", "No"),
    ("computrabajo.com", "Computrabajo", "No"), ("bne.cl", "Bolsa Nacional de Empleo", "No"),
    ("buk.cl", "Buk", "Sí"), ("hirevue.com", "HireVue", "Sí"), ("pymetrics.ai", "pymetrics", "Sí"),
    ("linkedin.com", "LinkedIn", "No"), ("indeed.com", "Indeed", "No"), ("gupy.io", "Gupy", "Sí"),
]

# Palabras clave del prefiltro (sin tildes, minúsculas). Si un correo no viene de
# un dominio conocido y no contiene ninguna, no se envía a la IA.
PALABRAS_CLAVE = [
    "postulacion", "postulaste", "postular", "proceso de seleccion", "seleccion",
    "practica profesional", "memoria", "pasantia", "internship", "intern",
    "graduate", "trainee", "entrevista", "evaluacion", "assessment", "candidato",
    "candidata", "curriculum", "cv", "oferta laboral", "carta oferta",
    "examen preocupacional", "examenes", "application", "interview", "recruit",
    "reclutamiento", "talento", "vacante", "solicitud", "your application",
]

# ---------------------------------------------------------------------------
# Datos ficticios (solo archivo DEMO). Todos con prefijo [DEMO] / DEMO-.
# ---------------------------------------------------------------------------
D = dt.date
DT = dt.datetime
ORIGEN_DEMO = "Datos ficticios (prueba)"
ORIGEN_INICIAL = "Registro inicial"

DEMO_EMPRESAS = [
    ("[DEMO] Minera Los Andes", "Minería", "demo-losandes.example", "", "Prueba", "Empresa ficticia."),
    ("[DEMO] Cobre Austral", "Minería", "demo-cobreaustral.example", "", "Prueba", "Empresa ficticia."),
    ("[DEMO] Celulosa del Sur", "Forestal y celulosa", "demo-celulosasur.example", "", "Prueba", "Empresa ficticia."),
    ("[DEMO] Equipos Pacífico", "Proveedores de equipos mineros", "demo-eqpacifico.example", "", "Prueba", "Empresa ficticia."),
    ("[DEMO] Litio Atacama", "Minería no metálica y litio", "demo-litio.example", "", "Prueba", "Empresa ficticia."),
]

# ID, Empresa, Cargo, Industria, Especialidad, Ubicación, Tipo, Fecha post., Estado,
# Fecha últ. act., Última comunicación, Próxima acción, Fecha seguimiento, Correo, Link, Evidencia, Obs
DEMO_POSTULACIONES = [
    ("DEMO-0001", "[DEMO] Minera Los Andes", "Memoria Planificación Mina 2027", "Minería",
     "Ingeniería Civil en Minas", "Santiago", "Memoria de título", D(2026, 8, 4), "Entrevista con jefatura",
     D(2026, 10, 22), "22-10-2026 · Invitación a entrevista", "Preparar entrevista con jefatura", D(2026, 11, 5),
     "talento@demo-losandes.example", "https://demo-losandes.example/empleos/123", "DEMO: correo enviado 04-08-2026", ""),
    ("DEMO-0002", "[DEMO] Minera Los Andes", "Práctica Mejora Continua", "Minería",
     "Ingeniería Civil Industrial", "Faena Antofagasta", "Práctica profesional", D(2026, 9, 10), "Rechazada",
     D(2026, 10, 2), "02-10-2026 · Rechazo", "", None,
     "talento@demo-losandes.example", "", "", ""),
    ("DEMO-0003", "[DEMO] Cobre Austral", "Programa Memoristas 2027", "Minería",
     "Doble titulación (Industrial + Minas)", "Santiago", "Memoria de título", D(2026, 9, 1), "Evaluación online",
     D(2026, 10, 25), "25-10-2026 · Invitación a evaluación", "Completar test online", D(2026, 11, 3),
     "seleccion@demo-cobreaustral.example", "", "", ""),
    ("DEMO-0004", "[DEMO] Celulosa del Sur", "Memoria Supply Chain", "Forestal y celulosa",
     "Ingeniería Civil Industrial", "Santiago", "Memoria de título", D(2026, 9, 15), "Postulación recibida",
     D(2026, 9, 15), "15-09-2026 · Confirmación de recepción", "Hacer seguimiento", D(2026, 10, 15),
     "", "", "", "Debería activar alerta de sin respuesta."),
    ("DEMO-0005", "[DEMO] Equipos Pacífico", "Internship Operaciones", "Proveedores de equipos mineros",
     "Ingeniería Civil en Minas", "Santiago", "Internship", D(2026, 10, 1), "En revisión",
     D(2026, 10, 20), "20-10-2026 · Revisión curricular", "Esperar respuesta", D(2026, 11, 10),
     "", "", "", ""),
    ("DEMO-0006", "[DEMO] Litio Atacama", "Memoria Procesos Litio", "Minería no metálica y litio",
     "Doble titulación (Industrial + Minas)", "Faena Salar de Atacama", "Memoria de título", D(2026, 7, 20),
     "Oferta recibida", D(2026, 10, 28), "28-10-2026 · Oferta laboral", "Responder oferta", D(2026, 11, 4),
     "rrhh@demo-litio.example", "", "DEMO: comprobante portal", ""),
    ("DEMO-0007", "[DEMO] Cobre Austral", "Pasantía Verano Geomecánica", "Minería",
     "Ingeniería Civil en Minas", "Faena Calama", "Pasantía", None, "En preparación",
     D(2026, 10, 30), "", "Terminar carta de presentación", D(2026, 11, 6), "", "", "", ""),
    ("DEMO-0008", "[DEMO] Equipos Pacífico", "Graduate Program 2027", "Proveedores de equipos mineros",
     "Ingeniería Civil Industrial", "Santiago", "Programa graduate", None, "Oportunidad identificada",
     D(2026, 10, 30), "", "Revisar requisitos", D(2026, 11, 12), "", "https://demo-eqpacifico.example/graduate", "", ""),
]

# ID msg, ID post., Empresa, Fecha, Remitente, Correo, Asunto, Tipo, Est. anterior, Est. propuesto,
# Est. confirmado, Resumen, Enlace, Confianza, Confirmación manual, Dirección, ID conversación,
# Cargo detectado, Justificación, Acción pendiente, Fecha evento, Método
DEMO_COMUNICACIONES = [
    ("DEMO-MSG-001", "DEMO-0001", "[DEMO] Minera Los Andes", DT(2026, 8, 4, 10, 15), "Yo (correo enviado)", "yo@demo.example",
     "Postulación Memoria Planificación Mina 2027", "Confirmación de postulación", "En preparación",
     "Postulación enviada", "Postulación enviada", "Correo enviado con CV y carta.", "", 0.97,
     "No requerida (automática)", "Enviado", "DEMO-CONV-1", "Memoria Planificación Mina 2027",
     "Correo enviado por la usuaria con CV adjunto.", "", "", "Claude API"),
    ("DEMO-MSG-002", "DEMO-0001", "[DEMO] Minera Los Andes", DT(2026, 8, 5, 9, 0), "Talento Los Andes", "talento@demo-losandes.example",
     "Hemos recibido tu postulación", "Confirmación de recepción", "Postulación enviada",
     "Postulación recibida", "Postulación recibida", "Acusan recibo de la postulación.", "", 0.98,
     "No requerida (automática)", "Recibido", "DEMO-CONV-1", "Memoria Planificación Mina 2027",
     "«Hemos recibido tu postulación».", "", "", "Claude API"),
    ("DEMO-MSG-003", "DEMO-0001", "[DEMO] Minera Los Andes", DT(2026, 9, 2, 12, 30), "Talento Los Andes", "talento@demo-losandes.example",
     "Avanzas a la siguiente etapa", "Avance en el proceso", "Postulación recibida",
     "En selección – etapa 1", "En selección – etapa 1", "Informan que continúa a la etapa 1.", "", 0.96,
     "No requerida (automática)", "Recibido", "DEMO-CONV-1", "Memoria Planificación Mina 2027",
     "«Has sido seleccionada para continuar».", "", "", "Claude API"),
    ("DEMO-MSG-004", "DEMO-0001", "[DEMO] Minera Los Andes", DT(2026, 9, 20, 16, 0), "Talento Los Andes", "talento@demo-losandes.example",
     "Invitación entrevista RR.HH.", "Invitación a entrevista", "En selección – etapa 1",
     "Entrevista con reclutamiento", "Entrevista con reclutamiento", "Invitan a entrevista con reclutamiento.", "", 0.95,
     "No requerida (automática)", "Recibido", "DEMO-CONV-1", "Memoria Planificación Mina 2027",
     "«Te invitamos a una entrevista».", "Confirmar horario", "2026-09-25", "Claude API"),
    ("DEMO-MSG-005", "DEMO-0001", "[DEMO] Minera Los Andes", DT(2026, 10, 22, 11, 45), "Jefa de Planificación", "planificacion@demo-losandes.example",
     "Entrevista con el área", "Invitación a entrevista", "Entrevista con reclutamiento",
     "Entrevista con jefatura", "Entrevista con jefatura", "Invitan a entrevista con jefatura de planificación.", "", 0.93,
     "No requerida (automática)", "Recibido", "DEMO-CONV-2", "Memoria Planificación Mina 2027",
     "«Te invitamos a una entrevista con la jefatura».", "Preparar entrevista", "2026-11-05", "Claude API"),
    ("DEMO-MSG-006", "DEMO-0002", "[DEMO] Minera Los Andes", DT(2026, 9, 11, 8, 0), "Portal Empleos", "no-reply@myworkday.com",
     "Gracias por postular", "Confirmación de postulación", "En preparación",
     "Postulación recibida", "Postulación recibida", "Portal confirma la postulación.", "", 0.97,
     "No requerida (automática)", "Recibido", "DEMO-CONV-3", "Práctica Mejora Continua",
     "«Gracias por postular».", "", "", "Claude API"),
    ("DEMO-MSG-007", "DEMO-0002", "[DEMO] Minera Los Andes", DT(2026, 10, 2, 17, 10), "Portal Empleos", "no-reply@myworkday.com",
     "Actualización de tu postulación", "Rechazo", "Postulación recibida",
     "Rechazada", "Rechazada", "Informan que no continúa en el proceso.", "", 0.92,
     "Confirmada", "Recibido", "DEMO-CONV-3", "Práctica Mejora Continua",
     "«Lamentamos informarte que no continuarás».", "", "", "Claude API"),
    ("DEMO-MSG-008", "DEMO-0003", "[DEMO] Cobre Austral", DT(2026, 9, 2, 10, 0), "Selección Cobre Austral", "seleccion@demo-cobreaustral.example",
     "Recepción de antecedentes", "Confirmación de recepción", "Por verificar",
     "Postulación recibida", "Postulación recibida", "Acusan recibo del CV.", "", 0.96,
     "No requerida (automática)", "Recibido", "DEMO-CONV-4", "Programa Memoristas 2027",
     "«Hemos recibido tus antecedentes».", "", "", "Claude API"),
    ("DEMO-MSG-009", "DEMO-0003", "[DEMO] Cobre Austral", DT(2026, 10, 25, 9, 30), "Selección Cobre Austral", "seleccion@demo-cobreaustral.example",
     "Invitación a evaluación online", "Invitación a evaluación", "Postulación recibida",
     "Evaluación online", "Evaluación online", "Invitan a completar test online.", "", 0.94,
     "No requerida (automática)", "Recibido", "DEMO-CONV-4", "Programa Memoristas 2027",
     "«Te invitamos a rendir una evaluación online».", "Completar test", "2026-11-03", "Claude API"),
    ("DEMO-MSG-010", "DEMO-0004", "[DEMO] Celulosa del Sur", DT(2026, 9, 15, 14, 0), "Celulosa del Sur", "no-reply@successfactors.com",
     "Hemos recibido tu postulación", "Confirmación de recepción", "Por verificar",
     "Postulación recibida", "Postulación recibida", "Confirmación automática de recepción.", "", 0.97,
     "No requerida (automática)", "Recibido", "DEMO-CONV-5", "Memoria Supply Chain",
     "«Hemos recibido tu postulación».", "", "", "Claude API"),
    ("DEMO-MSG-011", "DEMO-0005", "[DEMO] Equipos Pacífico", DT(2026, 10, 2, 9, 0), "Equipos Pacífico", "careers@demo-eqpacifico.example",
     "Application received", "Confirmación de recepción", "Por verificar",
     "Postulación recibida", "Postulación recibida", "Application received.", "", 0.96,
     "No requerida (automática)", "Recibido", "DEMO-CONV-6", "Internship Operaciones",
     "«We have received your application».", "", "", "Claude API"),
    ("DEMO-MSG-012", "DEMO-0005", "[DEMO] Equipos Pacífico", DT(2026, 10, 20, 15, 0), "Equipos Pacífico", "careers@demo-eqpacifico.example",
     "Your application is under review", "Revisión curricular", "Postulación recibida",
     "En revisión", "En revisión", "Indican que la postulación está en revisión.", "", 0.93,
     "No requerida (automática)", "Recibido", "DEMO-CONV-6", "Internship Operaciones",
     "«Your application is under review».", "", "", "Claude API"),
    ("DEMO-MSG-013", "DEMO-0006", "[DEMO] Litio Atacama", DT(2026, 7, 21, 10, 0), "RR.HH. Litio Atacama", "rrhh@demo-litio.example",
     "Recepción postulación memoria", "Confirmación de recepción", "Por verificar",
     "Postulación recibida", "Postulación recibida", "Acusan recibo.", "", 0.97,
     "No requerida (automática)", "Recibido", "DEMO-CONV-7", "Memoria Procesos Litio",
     "«Recibimos tu postulación».", "", "", "Claude API"),
    ("DEMO-MSG-014", "DEMO-0006", "[DEMO] Litio Atacama", DT(2026, 8, 18, 10, 0), "RR.HH. Litio Atacama", "rrhh@demo-litio.example",
     "Entrevista", "Invitación a entrevista", "Postulación recibida",
     "Entrevista con reclutamiento", "Entrevista con reclutamiento", "Invitan a entrevista.", "", 0.95,
     "No requerida (automática)", "Recibido", "DEMO-CONV-7", "Memoria Procesos Litio",
     "«Te invitamos a una entrevista».", "", "", "Claude API"),
    ("DEMO-MSG-015", "DEMO-0006", "[DEMO] Litio Atacama", DT(2026, 9, 30, 10, 0), "RR.HH. Litio Atacama", "rrhh@demo-litio.example",
     "Exámenes preocupacionales", "Exámenes médicos", "Entrevista con reclutamiento",
     "Exámenes médicos", "Exámenes médicos", "Solicitan exámenes preocupacionales.", "", 0.94,
     "No requerida (automática)", "Recibido", "DEMO-CONV-7", "Memoria Procesos Litio",
     "«Debes realizar exámenes preocupacionales».", "Agendar exámenes", "2026-10-07", "Claude API"),
    ("DEMO-MSG-016", "DEMO-0006", "[DEMO] Litio Atacama", DT(2026, 10, 28, 18, 0), "RR.HH. Litio Atacama", "rrhh@demo-litio.example",
     "Carta oferta memoria", "Oferta laboral", "Exámenes médicos",
     "Oferta recibida", "Oferta recibida", "Envían carta oferta para la memoria.", "", 0.90,
     "Confirmada", "Recibido", "DEMO-CONV-7", "Memoria Procesos Litio",
     "«Nos complace ofrecerte la memoria».", "Responder oferta", "2026-11-04", "Claude API"),
    ("DEMO-MSG-017", "", "", DT(2026, 10, 29, 12, 0), "Talento Andes", "talento@demo-losandes.example",
     "Novedades de tu proceso", "Avance en el proceso", "",
     "En selección – etapa 2", "", "Correo ambiguo: hay dos postulaciones abiertas en la empresa.", "", 0.70,
     "Pendiente", "Recibido", "DEMO-CONV-8", "",
     "No indica el cargo; dos postulaciones posibles.", "", "", "Claude API"),
    ("DEMO-MSG-018", "", "", DT(2026, 10, 30, 9, 0), "Reclutamiento", "jobs@demo-desconocida.example",
     "Invitación a proceso", "Invitación a evaluación", "",
     "Evaluación online", "", "Empresa no registrada en catálogos.", "", 0.62,
     "Pendiente", "Recibido", "DEMO-CONV-9", "Analista Junior",
     "Empresa no reconocida.", "", "", "Reglas"),
]

DEMO_PENDIENTES = [
    # ID msg, Fecha, Remitente, Correo, Asunto, Resumen, Enlace, Emp sug, Cargo sug, Tipo sug, Estado sug,
    # ID post sug, Confianza, Motivo
    ("DEMO-MSG-017", DT(2026, 10, 29, 12, 0), "Talento Andes", "talento@demo-losandes.example",
     "Novedades de tu proceso", "Correo ambiguo: hay dos postulaciones abiertas en la empresa.", "",
     "[DEMO] Minera Los Andes", "", "Avance en el proceso", "En selección – etapa 2", "DEMO-0001", 0.70,
     "Tipo sensible con confianza bajo el umbral; varias postulaciones posibles."),
    ("DEMO-MSG-018", DT(2026, 10, 30, 9, 0), "Reclutamiento", "jobs@demo-desconocida.example",
     "Invitación a proceso", "Empresa no registrada en catálogos.", "",
     "[DEMO] Empresa Desconocida", "Analista Junior", "Invitación a evaluación", "Evaluación online", "", 0.62,
     "No existe postulación registrada para esta empresa."),
]
