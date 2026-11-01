# Guía de configuración: conectar el sistema a tu cuenta real

Sigue los pasos en orden. Tiempo estimado: 1,5–2 horas la primera vez.
**Nada de esto está configurado todavía en tu cuenta.** Cada paso indica cómo verificar que quedó funcionando.

## 0. Requisitos

| Requisito | Cómo verificarlo |
|---|---|
| Cuenta Microsoft 365 **del trabajo o la escuela** (p. ej., la de la UAI) con Outlook, OneDrive y Excel para la Web | Entra en <https://www.office.com> con esa cuenta |
| Office Scripts habilitado | Abre cualquier Excel en el navegador: debe aparecer la pestaña **Automatizar** |
| Power Automate | <https://make.powerautomate.com> abre con tu cuenta |
| (Opcional, para usar IA) Licencia **Power Automate Premium** o prueba gratuita | En Power Automate, al agregar la acción **HTTP** no debe aparecer «Premium – requiere actualización» |
| (Opcional) Cuenta en la consola de Anthropic con crédito | <https://console.anthropic.com> |

Notas importantes:
- El conector *Run script* **solo funciona con OneDrive para el trabajo o la escuela / SharePoint**. Si tu correo es personal (@outlook.com / @hotmail.com), necesitas una cuenta Microsoft 365 de trabajo o escuela para alojar el Excel y ejecutar los flujos.
- Si la universidad tiene deshabilitado Office Scripts o Power Automate, el administrador de TI tiene que habilitarlo; ningún paso de esta guía puede saltarse esa restricción.
- Todo corre en la nube de Microsoft: **no depende de que tu computador esté encendido**.

## 1. Subir el Excel a OneDrive

1. Descarga `excel/Seguimiento_Postulaciones_2027.xlsx` (el archivo **sin** `_DEMO`).
2. En OneDrive crea la carpeta `Postulaciones` y sube el archivo ahí.
3. Ábrelo en Excel para la Web y verifica: hojas Dashboard, Postulaciones (15 empresas «Por verificar»), Comunicaciones, Catálogos, Pendientes, Guía y Cálculos.
4. Antes de conectarlo a nada, prueba con el archivo `_DEMO` (sección 7).

## 2. Crear los Office Scripts

Para cada archivo de la carpeta `office-scripts/` (no la subcarpeta `src/`):
1. En Excel para la Web: **Automatizar → Nuevo script**.
2. Borra el contenido de ejemplo, pega el archivo completo y presiona **Guardar script**.
3. Cambia el nombre del script exactamente a: `RegistrarCorreo`, `EvaluarRelevancia`, `ProcesarPendientes`, `ActualizarDashboard`, `CrearAnalisisDinamico` y `LimpiarDatosFicticios`.

Los scripts quedan en tu OneDrive (`Documentos/Office Scripts`) y sirven para cualquier libro con esta estructura.

Verificación: ejecuta **ActualizarDashboard** desde Automatizar. En la esquina superior derecha del Dashboard debe aparecer «Actualizado por script: …».

Luego ejecuta **CrearAnalisisDinamico** una vez: crea la hoja «Análisis» con 5 tablas dinámicas y 4 segmentadores.

Opcional: en la hoja Pendientes, **Automatizar → (script) → … → Agregar en libro** para tener un botón de ProcesarPendientes.

## 3. Elegir el modo de clasificación

| Modo | Requiere | Resultado |
|---|---|---|
| **IA (recomendado)** | Power Automate Premium + clave de la API de Claude | Clasifica los 14 tipos, extrae empresa, cargo, fechas y acciones. Aplica sola lo que es claro y manda lo dudoso a Pendientes |
| **Reglas** | Solo licencias estándar | Reconoce frases típicas en español e inglés. Registra confirmaciones claras; todo avance, oferta o rechazo queda en Pendientes |

Puedes empezar en modo reglas y pasar a IA después: solo se agrega el bloque HTTP.

## 4. Configurar la API de Claude (modo IA)

1. Entra en <https://console.anthropic.com> → **Settings → Billing** y carga crédito (por ejemplo, US$10; el costo estimado está en `02_flujos_power_automate.md`).
2. **Settings → Limits**: define un límite mensual de gasto (por ejemplo, US$10) para evitar sorpresas.
3. **API Keys → Create Key**, con el nombre `power-automate-postulaciones`. Copia la clave: no se vuelve a mostrar.
4. Guárdala solo en el encabezado `x-api-key` de la acción HTTP, con **Secure inputs** activado. No la pegues en el Excel, en correos ni en el repositorio.
5. (Recomendado) Antes de crear el flujo, valida el clasificador desde un computador con Python:
   ```bash
   pip install anthropic
   export ANTHROPIC_API_KEY="tu-clave"
   python3 tests/probar_clasificador.py
   ```
   Clasifica 12 correos ficticios y muestra los aciertos. Si alguno falla, revisa o ajusta `tools/generar_artefactos_ia.py` y regenera con `python3 tools/generar_artefactos_ia.py`.

Datos que se envían a la API: remitente, destinatarios, asunto, fecha, hasta 12.000 caracteres del texto del correo y la lista de empresas y postulaciones abiertas. Solo se envían los correos que pasan el prefiltro (dominios del catálogo, plataformas de reclutamiento o palabras clave).

## 5. Conectar Outlook y crear el Flujo 1 (correos recibidos)

1. En <https://make.powerautomate.com>: **Crear → Flujo de nube automatizado**, con el nombre «Postulaciones – correos recibidos».
2. Disparador: **When a new email arrives (V3)** (Office 365 Outlook). Al agregarlo se te pedirá iniciar sesión: usa la misma cuenta del correo. Esta es la conexión con Outlook; el flujo solo **lee** correos.
3. Construye las acciones exactamente como en `02_flujos_power_automate.md` (Flujo 1). Al agregar *Run script* se crea la conexión con Excel Online (Business).
4. En los ajustes del disparador: **Concurrency control → On → 1**.
5. Guarda el flujo y **déjalo desactivado** (*Turn off*) hasta la prueba controlada del paso 8.

## 6. Crear los Flujos 2, 3 y 4

- **Flujo 2 (enviados):** *Save As* del Flujo 1, carpeta `Sent Items`, `direccion` = `Enviado`.
- **Flujo 3 (histórico):** flujo de nube instantáneo; ver diseño.
- **Flujo 4 (diario):** flujo programado que ejecuta `ActualizarDashboard`.

## 7. Prueba en seco con el archivo DEMO

1. Sube `excel/Seguimiento_Postulaciones_2027_DEMO.xlsx` a otra carpeta (por ejemplo, `Postulaciones/Pruebas`).
2. Recorre el Dashboard: debe mostrar 6 postulaciones enviadas, 20 empresas, 5 activas, 2 en revisión, 2 en selección, 3 entrevistas, 1 evaluación pendiente, 1 oferta, 1 rechazo y 1 sin respuesta.
3. Cambia el filtro Industria a «Minería»: debe mostrar 3 enviadas, 10 empresas, 2 activas y 2 en selección.
4. En Pendientes completa: fila DEMO-MSG-017 → ID de postulación confirmada `DEMO-0001`, Decisión «Registrar sin cambiar estado»; fila DEMO-MSG-018 → Empresa `[DEMO] Empresa Desconocida`, Cargo `Analista Junior`, Estado `Evaluación online`, Decisión «Crear nueva postulación». Ejecuta **ProcesarPendientes**: ambas filas deben quedar en «Procesado = Sí» y aparecer una postulación nueva `POST-0016`.
5. Ejecuta **LimpiarDatosFicticios**: deben quedar solo los 15 registros iniciales.

## 8. Prueba controlada con tu correo real

1. Apunta temporalmente el Flujo 1 al archivo DEMO (en ambas acciones *Run script*) y actívalo.
2. Desde otra cuenta tuya, envíate un correo con asunto «Hemos recibido tu postulación» y texto «Hemos recibido tu postulación a Práctica Profesional en [DEMO] Litio Atacama».
3. En el historial de ejecuciones del flujo revisa que:
   - `EvaluarRelevancia` devolvió `procesar: true`;
   - `Llamar_Claude` respondió `200` (modo IA);
   - `RegistrarCorreo` devolvió `accion` = `registrado`, `creado` o `pendiente`.
4. En el Excel DEMO verifica la nueva fila en Comunicaciones y que el enlace abre el correo en Outlook.
5. Envía el **mismo** correo otra vez con *Resubmit*: el resultado debe ser `duplicado` y no debe aparecer una fila nueva.
6. Envía un correo irrelevante (por ejemplo, «Almuerzo el viernes»): `procesar` debe ser `false`.
7. Cuando todo lo anterior esté OK, apunta los flujos al archivo real, actívalos y ejecuta el Flujo 3 (histórico), mes por mes.

Solo después de completar este paso puedes considerar que **Outlook está conectado y la automatización funciona**.

## 9. Revisión inicial de las 15 empresas

Los registros POST-0001 a POST-0015 están en «Por verificar» porque no hay antecedentes de su estado. Después de la importación histórica:
- Si llegó una confirmación, el sistema ya los habrá actualizado a «Postulación recibida» con su cargo.
- Para los que sigan en «Por verificar», decide a mano: «Oportunidad identificada», «En preparación» o «Postulación enviada» (solo si tienes evidencia: escribe cuál en «Evidencia de envío»).
- Si postulaste a dos programas de la misma empresa, agrega una fila nueva (`POST-0016`, …) en lugar de reutilizar la existente.
