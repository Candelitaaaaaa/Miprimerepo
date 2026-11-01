# Mantenimiento del sistema

## Rutina recomendada

| Frecuencia | Tarea |
|---|---|
| 2–3 veces por semana | Hoja **Pendientes**: completar las celdas amarillas, elegir una Decisión y ejecutar **ProcesarPendientes** |
| Semanal | Revisar «Próximos seguimientos» y «Seguimientos vencidos» en el Dashboard |
| Semanal | Power Automate → *Mis flujos* → revisar ejecuciones con error |
| Mensual | Revisar el gasto en console.anthropic.com y la columna «Nivel de confianza» de Comunicaciones |
| Al agregar empresas | Completar dominios y alias en Catálogos > Empresas |

## Agregar o cambiar catálogos

Todas las listas están en la hoja **Catálogos**, como tablas estructuradas. Para agregar un elemento, escribe en la primera fila vacía **debajo** de la tabla: Excel la extiende sola.

- **Empresas:** completa `Dominios de correo` (por ejemplo, `bhp.com`; varios separados por `;`) y `Alias` (nombres alternativos, separados por `;`). Los dominios permiten vincular correos aunque no mencionen el nombre de la empresa. Los dominios iniciales son referenciales: confírmalos con el primer correo real de cada empresa.
- **Estados:** puedes agregar o renombrar estados. Cada estado necesita:
  - `Orden` (avance; los números mayores son etapas posteriores),
  - `Grupo` (agrupación del Dashboard: Pre-postulación, Enviada, En revisión, Selección, Evaluación, Entrevista, Oferta, Rechazo, Sin respuesta, Cerrado, Por verificar),
  - `Categoría de regla` (Inicial, Postulación, Selección, Final, Especial),
  - `Activa` y `Cuenta como postulada` (Sí/No).
  Si renombras un estado, actualiza también `Estado propuesto por defecto` y `Estados permitidos` en Tipos de comunicación. Si quieres que la IA lo pueda proponer, agrégalo a `ESTADOS` en `src/catalogos.py` y regenera los artefactos (ver más abajo).
- **Tipos de comunicación:** las columnas `Cambia estado` y `Sensible` controlan las reglas. Un tipo «Sensible = Sí» exige confianza ≥ UmbralConfianzaSensible y una frase explícita.
- **Palabras clave y dominios de plataformas:** controlan el prefiltro. Agrega una palabra si notas correos relevantes que no se procesan; marca «Relevante por dominio = No» si una plataforma envía mucha publicidad.

## Parámetros (Catálogos > Parámetros)

| Parámetro | Valor inicial | Efecto |
|---|---|---|
| UmbralConfianzaAutomatica | 0,85 | Bájalo para que más correos se apliquen solos; súbelo para revisar más |
| UmbralConfianzaSensible | 0,95 | Umbral para Avance, Oferta y Rechazo |
| DiasSinRespuesta | 21 | Días para la alerta «sin respuesta» |
| DiasSeguimiento | 7 | Plazo por defecto de la próxima acción |
| CrearPostulacionesAutomaticas | Sí | Con «No», toda postulación nueva pasa por Pendientes |
| MarcarSinRespuestaAutomatico | No | Con «Sí», el script diario cambia el estado a «Sin respuesta» |
| MesInicioDashboard | 01-07-2026 | Primer mes del gráfico mensual |

No cambies el nombre de la tabla ni de los parámetros.

## Correcciones manuales

- **Cambiar un estado a mano:** edítalo directamente en Postulaciones y anota el motivo en Observaciones. El historial de correos no se altera.
- **Un correo quedó vinculado a la postulación equivocada:** en Comunicaciones cambia «ID de postulación», «Estado confirmado» y marca «Confirmación manual = Corregida». Luego corrige el estado en Postulaciones si corresponde.
- **Anular un correo mal registrado:** «Confirmación manual = Descartada». No borres la fila.
- **Agregar una postulación sin correo** (por ejemplo, desde un portal): agrega la fila en Postulaciones con un ID nuevo (`POST-00NN`, el siguiente número), Origen = Manual.

## Lo que no se debe hacer

- Renombrar hojas, tablas o columnas. Los scripts las buscan por su nombre exacto y fallarán con un mensaje que dice cuál falta.
- Escribir en las columnas grises (son fórmulas) o en la hoja Cálculos.
- Guardar el archivo como «Hoja de cálculo Open XML estricta» (el conector no la admite).
- Ordenar la tabla de Parámetros o mover sus columnas.

## Regenerar archivos desde el repositorio

```bash
python3 src/build_workbook.py              # vuelve a crear los dos Excel (¡archivos vacíos de datos reales!)
python3 tools/bundle_office_scripts.py     # vuelve a armar office-scripts/*.ts desde office-scripts/src
python3 tools/generar_artefactos_ia.py     # vuelve a generar prompt, esquema y plantilla de la IA
tests/run_tests.sh                         # pruebas (requiere LibreOffice, node y tsc)
```

`build_workbook.py` crea un archivo **nuevo**: no lo uses para reemplazar el Excel que ya tiene tus datos. Para cambiar la estructura de un archivo en uso, edítalo en Excel.

Si cambias el prompt o los tipos, vuelve a pegar `prompt_sistema.txt` y `plantilla_cuerpo_claude.json` en las acciones Compose del flujo.

## Cambiar el modelo de IA

Edita `model` en la acción Compose `Plantilla_cuerpo` (y en `tools/generar_artefactos_ia.py`). Antes de usarlo en el flujo, ejecuta `tests/probar_clasificador.py` con el modelo nuevo. Actualiza también el parámetro informativo `ModeloIA`.

## Límites a tener presentes

- *Run script*: 1.600 ejecuciones por usuario al día; 120 segundos por ejecución.
- Excel para la Web funciona bien con miles de filas, pero si Comunicaciones supera ~5.000 filas, archiva los años anteriores en otro libro (copia y luego deja una fila resumen).
- El enlace a Outlook usa el identificador de Graph, que cambia si mueves el correo de carpeta. El asunto y la fecha quedan registrados para buscarlo.
