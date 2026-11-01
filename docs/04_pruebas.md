# Pruebas realizadas y pendientes

Fecha: 01-11-2026. Todas las pruebas usaron **datos ficticios** (prefijo `[DEMO]` / `DEMO-` y dominios `*.example`), salvo los 15 registros iniciales «Por verificar».

## Probado ✅

### Excel (fórmulas y dashboard)
Recalculado con LibreOffice (motor distinto de Excel; ver limitaciones).

| Verificación | Resultado |
|---|---|
| Errores de fórmula, archivo de producción | 0 de 1.141 fórmulas |
| Errores de fórmula, archivo DEMO | 0 de 1.214 fórmulas |
| KPIs DEMO sin filtros (enviadas, empresas, activas, revisión, selección, entrevistas, evaluaciones, ofertas, rechazos, sin respuesta) | 6, 20, 5, 2, 2, 3, 1, 1, 1, 1 — coinciden con el cálculo manual de los datos ficticios |
| Filtro Industria = Minería | 3, 10, 2, 0, 2, 2, 1, 0, 1, 0 — coincide |
| Filtros Tipo = Memoria de título + Especialidad = Doble titulación | 2, 2, 2, 0, 1, 1, 1, 1, 0, 0 — coincide |
| Días desde última comunicación, alerta «sin respuesta» (DEMO-0004 con 47 días) | Correcto |
| Días de espera por etapa (p. ej., recepción 1 día después del envío) | Correcto |
| Evolución mensual (jul 1, ago 1, sep 3, oct 1) | Correcto |
| Próximos seguimientos ordenados por fecha | Correcto |
| Tablas con nombres de columna exactos y columnas calculadas declaradas | Verificado en el XML |
| Formatos condicionales y validaciones sin referencias estructuradas (no admitidas por Excel) | Verificado en el XML |
| Render visual del Dashboard | Revisado en PDF exportado |

### Office Scripts (lógica)
Compilados con TypeScript en modo `strict` contra las firmas de la API de ExcelScript que se usan, y ejecutados con un **libro simulado** (`tests/mock.ts`) cargado desde las tablas reales del Excel generado. `tests/run_tests.sh` → **75 verificaciones OK, 0 fallas**.

Escenarios cubiertos en `RegistrarCorreo` (42 verificaciones):
1. Confirmación de recepción vincula el registro inicial de la empresa, completa el cargo e infiere el tipo de oportunidad; **no** llena «Evidencia de envío».
2. Mismo mensaje dos veces → `duplicado`, sin filas nuevas.
3. Rechazo con 97 % de confianza pero sin una frase explícita → Pendientes.
4. Rechazo explícito → «Rechazada» y limpia la próxima acción.
5. Correo nuevo sobre un proceso finalizado → Pendientes.
6. Si la IA sugiere «Postulación enviada» para un correo recibido → se registra «Postulación recibida».
7. Correo **enviado** con CV → crea la postulación «Postulación enviada» con evidencia (enlace al correo).
8. «Has sido seleccionada para continuar» → «En selección – etapa 1».
9. Correo de «en revisión» posterior → no retrocede; queda solo en el historial.
10. Vínculo por ID de conversación + próxima acción con fecha límite.
11. Sin IA (modo reglas): confirmación clara → aplicada.
12. Modo reglas con avance → siempre Pendientes.
13. Empresa desconocida con baja confianza → Pendientes.
14. Correo irrelevante con alta confianza → no se registra.
15. Respuesta de IA inválida → respaldo por reglas.
16. Asunto que empieza con «=» se guarda como texto (no como fórmula).
17. Oferta explícita → «Oferta recibida».
18. Correo sin cargo → se vincula a la única postulación abierta de la empresa.
19. Integridad: sin IDs de mensaje ni de postulación duplicados; nunca se escribió sobre columnas calculadas.

También: `EvaluarRelevancia` (prefiltro, duplicados, contexto para la IA), `ProcesarPendientes` (vincular, crear, error de validación visible, idempotencia), `LimpiarDatosFicticios` y `ActualizarDashboard` (marcar «Sin respuesta»). El clasificador por reglas acierta **12/12** correos ficticios de `tests/correos_prueba.json`.

## No probado ⚠️ (requiere tu cuenta)

| Componente | Por qué no se pudo probar aquí | Cómo probarlo |
|---|---|---|
| Apertura del .xlsx en Excel para la Web / escritorio | Este entorno no tiene Excel | Abrir el archivo; no debe aparecer «Reparar». Ver `01_configuracion.md`, paso 1 |
| Ejecución real de los Office Scripts en Excel | Se probó la lógica con un libro simulado; el motor real de Excel no estaba disponible | `01_configuracion.md`, pasos 2 y 7 |
| Tablas dinámicas y segmentadores (`CrearAnalisisDinamico`) | Solo verificación de tipos | Ejecutar el script una vez |
| Flujos de Power Automate | Requieren tu cuenta y tus licencias | `01_configuracion.md`, paso 8 |
| Conexión con Outlook | Ídem | Ídem |
| Clasificación con Claude API | No hay clave de API en este entorno | `python3 tests/probar_clasificador.py` |

**Outlook no está conectado y la automatización no está funcionando** hasta que completes el paso 8 de la guía de configuración.

## Cómo repetir las pruebas

```bash
# Requisitos: python3 + openpyxl, LibreOffice, node, TypeScript (tsc)
tests/run_tests.sh
```
