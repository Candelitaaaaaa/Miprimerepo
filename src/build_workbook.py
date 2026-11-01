"""Genera Seguimiento_Postulaciones_2027.xlsx (producción) y la versión DEMO.

Uso:
    python3 src/build_workbook.py            # genera ambos archivos en excel/

El archivo de producción contiene solo los 15 registros iniciales «Por verificar».
El archivo DEMO agrega datos ficticios (prefijo [DEMO] / DEMO-) para probar
fórmulas, dashboard y scripts.
"""
from __future__ import annotations

import datetime as dt
import os
import sys

from openpyxl import Workbook
from openpyxl.chart import BarChart, DoughnutChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.worksheet.table import Table, TableColumn, TableFormula, TableStyleInfo

sys.path.insert(0, os.path.dirname(__file__))
import catalogos as C  # noqa: E402

FECHA_CREACION = dt.date(2026, 11, 1)

# ---------------------------------------------------------------------------
# Estilo
# ---------------------------------------------------------------------------
FONT = "Arial"
NAVY = "1F2A44"
INK = "0B0B0B"
INK2 = "52514E"
MUTED = "8A8984"
SURFACE = "FCFCFB"
LINE = "E4E3DF"
BLUE = "2A78D6"
GOOD = "0CA30C"
WARN = "FAB219"
SERIOUS = "EC835A"
CRIT = "D03B3B"
GREY = "A8A7A2"
INPUT_FILL = "FFF7D6"
CAT = ["2A78D6", "EB6834", "1BAF7A", "EDA100", "E87BA4", "008300", "6250D6", "E34948"]


def font(size=10, bold=False, color=INK, italic=False):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic)


def fill(hex_):
    return PatternFill("solid", start_color=hex_, end_color=hex_)


THIN = Side(style="thin", color=LINE)

# ---------------------------------------------------------------------------
# Definición de tablas
# ---------------------------------------------------------------------------
POST_COLS = [
    "ID de postulación", "Empresa", "Cargo o programa", "Industria", "Especialidad",
    "Ubicación", "Tipo de oportunidad", "Fecha de postulación", "Estado actual",
    "Fecha de última actualización", "Última comunicación", "Próxima acción",
    "Fecha de seguimiento", "Días desde última comunicación", "Correo de contacto",
    "Link a la oferta", "Evidencia de envío", "Observaciones",
    # Columnas auxiliares (calculadas) para dashboard y reglas
    "Grupo de estado", "Activa", "Cuenta como postulada", "Mes de postulación",
    "Alerta sin respuesta", "N° entrevistas", "Origen del registro",
]
COM_COLS = [
    "ID del mensaje", "ID de postulación", "Empresa", "Fecha de recepción", "Remitente",
    "Correo del remitente", "Asunto", "Tipo de comunicación", "Estado anterior",
    "Estado propuesto", "Estado confirmado", "Resumen del mensaje",
    "Enlace al correo original en Outlook", "Nivel de confianza de clasificación",
    "Confirmación manual",
    # Columnas adicionales de trazabilidad
    "Dirección", "ID de conversación", "Cargo detectado", "Justificación de la clasificación",
    "Acción pendiente detectada", "Fecha del evento", "Método de clasificación",
    "Fecha de registro", "Días de espera",
]
PEND_COLS = [
    "ID del mensaje", "Fecha de recepción", "Remitente", "Correo del remitente", "Asunto",
    "Resumen del mensaje", "Enlace al correo", "Empresa sugerida", "Cargo sugerido",
    "Tipo sugerido", "Estado sugerido", "ID de postulación sugerida", "Confianza",
    "Motivo de revisión",
    # Confirmación manual (celdas amarillas)
    "Empresa confirmada", "Cargo confirmado", "ID de postulación confirmada",
    "Tipo de comunicación confirmado", "Estado confirmado", "Decisión",
    # Resultado (lo escribe el script ProcesarPendientes)
    "Procesado", "Fecha de procesamiento", "Resultado del procesamiento",
]
PEND_INPUT = {"Empresa confirmada", "Cargo confirmado", "ID de postulación confirmada",
              "Tipo de comunicación confirmado", "Estado confirmado", "Decisión"}
DECISIONES = ["Pendiente", "Aplicar", "Crear nueva postulación",
              "Registrar sin cambiar estado", "Descartar (irrelevante)"]
CONFIRMACIONES = ["No requerida (automática)", "Pendiente", "Confirmada", "Corregida", "Descartada"]
ORIGENES = [C.ORIGEN_INICIAL, "Manual", "Automático (correo)", C.ORIGEN_DEMO]


def P(col):
    return f"tblPostulaciones[[#This Row],[{col}]]"


def Cm(col):
    return f"tblComunicaciones[[#This Row],[{col}]]"


PARAM = 'INDEX(tblParametros[Valor],MATCH("{0}",tblParametros[Parámetro],0))'

_ULT = (f'MAX(_xlfn.MAXIFS(tblComunicaciones[Fecha de recepción],tblComunicaciones[ID de postulación],'
        f'{P("ID de postulación")}),N({P("Fecha de postulación")}))')

POST_FORMULAS = {
    "Días desde última comunicación":
        f'=IF({P("ID de postulación")}="","",IF({_ULT}=0,"",TODAY()-INT({_ULT})))',
    "Grupo de estado":
        f'=IFERROR(INDEX(tblEstados[Grupo],MATCH({P("Estado actual")},tblEstados[Estado],0)),"Por verificar")',
    "Activa":
        f'=IFERROR(INDEX(tblEstados[Activa],MATCH({P("Estado actual")},tblEstados[Estado],0)),"No")',
    "Cuenta como postulada":
        f'=IFERROR(INDEX(tblEstados[Cuenta como postulada],MATCH({P("Estado actual")},tblEstados[Estado],0)),"No")',
    "Mes de postulación":
        f'=IF(ISNUMBER({P("Fecha de postulación")}),DATE(YEAR({P("Fecha de postulación")}),MONTH({P("Fecha de postulación")}),1),"")',
    "Alerta sin respuesta":
        (f'=IF({P("ID de postulación")}="","",IF(OR({P("Estado actual")}="Sin respuesta",'
         f'AND({P("Activa")}="Sí",OR({P("Grupo de estado")}="Enviada",{P("Grupo de estado")}="En revisión"),'
         f'N({P("Días desde última comunicación")})>={PARAM.format("DiasSinRespuesta")})),"Sí","No"))'),
    "N° entrevistas":
        (f'=IF({P("ID de postulación")}="",0,COUNTIFS(tblComunicaciones[ID de postulación],{P("ID de postulación")},'
         f'tblComunicaciones[Tipo de comunicación],"Invitación a entrevista",'
         f'tblComunicaciones[Confirmación manual],"<>Descartada"))'),
}

_PREV = (f'MAX(_xlfn.MAXIFS(tblComunicaciones[Fecha de recepción],tblComunicaciones[ID de postulación],'
         f'{Cm("ID de postulación")},tblComunicaciones[Fecha de recepción],"<"&{Cm("Fecha de recepción")}),'
         f'IFERROR(N(INDEX(tblPostulaciones[Fecha de postulación],MATCH({Cm("ID de postulación")},'
         f'tblPostulaciones[ID de postulación],0))),0))')
COM_FORMULAS = {
    "Días de espera":
        (f'=IFERROR(IF(OR({Cm("ID de postulación")}="",{Cm("Estado confirmado")}="",'
         f'{Cm("Estado confirmado")}={Cm("Estado anterior")}),"",IF({_PREV}=0,"",'
         f'MAX(0,INT({Cm("Fecha de recepción")})-INT({_PREV})))),"")'),
}

DATE_FMT = "dd-mm-yyyy"
DATETIME_FMT = "dd-mm-yyyy hh:mm"

POST_WIDTHS = {
    "ID de postulación": 13, "Empresa": 24, "Cargo o programa": 32, "Industria": 22,
    "Especialidad": 24, "Ubicación": 18, "Tipo de oportunidad": 19, "Fecha de postulación": 13,
    "Estado actual": 26, "Fecha de última actualización": 13, "Última comunicación": 32,
    "Próxima acción": 28, "Fecha de seguimiento": 13, "Días desde última comunicación": 12,
    "Correo de contacto": 28, "Link a la oferta": 28, "Evidencia de envío": 26, "Observaciones": 40,
    "Grupo de estado": 15, "Activa": 8, "Cuenta como postulada": 11, "Mes de postulación": 11,
    "Alerta sin respuesta": 11, "N° entrevistas": 10, "Origen del registro": 22,
}
COM_WIDTHS = {
    "ID del mensaje": 22, "ID de postulación": 13, "Empresa": 22, "Fecha de recepción": 16,
    "Remitente": 22, "Correo del remitente": 28, "Asunto": 34, "Tipo de comunicación": 25,
    "Estado anterior": 24, "Estado propuesto": 24, "Estado confirmado": 24,
    "Resumen del mensaje": 44, "Enlace al correo original en Outlook": 22,
    "Nivel de confianza de clasificación": 12, "Confirmación manual": 22, "Dirección": 10,
    "ID de conversación": 18, "Cargo detectado": 26, "Justificación de la clasificación": 40,
    "Acción pendiente detectada": 26, "Fecha del evento": 13, "Método de clasificación": 14,
    "Fecha de registro": 16, "Días de espera": 10,
}


# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------
def sheet_title(ws, title, subtitle, width_cols):
    ws.sheet_view.showGridLines = False
    ws["A1"] = title
    ws["A1"].font = font(16, True, NAVY)
    ws["A2"] = subtitle
    ws["A2"].font = font(9, False, INK2, italic=True)
    ws.row_dimensions[1].height = 26


def add_table(ws, name, cols, rows, header_row, first_col=1, style="TableStyleLight9",
              formulas=None, widths=None, date_cols=(), datetime_cols=(), pct_cols=(), int_cols=()):
    """Escribe encabezados + filas y crea una tabla estructurada."""
    formulas = formulas or {}
    for j, c in enumerate(cols):
        cell = ws.cell(row=header_row, column=first_col + j, value=c)
        cell.font = font(10, True, "FFFFFF")
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    for i, r in enumerate(rows):
        for j, c in enumerate(cols):
            cell = ws.cell(row=header_row + 1 + i, column=first_col + j)
            if c in formulas:
                cell.value = formulas[c]
            else:
                v = r[j] if j < len(r) else None
                cell.value = v if v != "" else None
            cell.font = font(10)
            cell.alignment = Alignment(vertical="top", wrap_text=False)
            if c in date_cols:
                cell.number_format = DATE_FMT
            if c in datetime_cols:
                cell.number_format = DATETIME_FMT
            if c in pct_cols:
                cell.number_format = "0%"
            if c in int_cols:
                cell.number_format = "0"
    n = max(len(rows), 1)
    if not rows:  # tabla vacía: una fila en blanco (Excel exige al menos una)
        for j, c in enumerate(cols):
            cell = ws.cell(row=header_row + 1, column=first_col + j)
            if c in formulas:
                cell.value = formulas[c]
            cell.font = font(10)
    last_col = get_column_letter(first_col + len(cols) - 1)
    ref = f"{get_column_letter(first_col)}{header_row}:{last_col}{header_row + n}"
    t = Table(displayName=name, ref=ref)
    t.tableStyleInfo = TableStyleInfo(name=style, showRowStripes=True, showColumnStripes=False)
    # Las columnas se declaran explícitamente con el texto exacto del encabezado
    # (si no, openpyxl usa «Column1…» y Excel «repara» el archivo al abrirlo).
    for j, c in enumerate(cols, start=1):
        tc = TableColumn(id=j, name=c)
        if c in formulas:
            tc.calculatedColumnFormula = TableFormula(attr_text=formulas[c][1:])
        t.tableColumns.append(tc)
    ws.add_table(t)
    ws.row_dimensions[header_row].height = 32
    if widths:
        for j, c in enumerate(cols):
            ws.column_dimensions[get_column_letter(first_col + j)].width = widths.get(c, 14)
    return t, header_row + n


def col_letter(cols, name, first_col=1):
    return get_column_letter(first_col + cols.index(name))


# ---------------------------------------------------------------------------
# Hojas
# ---------------------------------------------------------------------------
def build_catalogos(wb, demo):
    ws = wb.create_sheet("Catálogos")
    sheet_title(ws, "Catálogos y parámetros",
                "Agregue filas al final de cada tabla para crear nuevas empresas, industrias, estados o tipos. "
                "No cambie los nombres de las columnas ni el orden de la tabla de Parámetros.", 30)
    hr = 4
    empresas = list(C.EMPRESAS) + (list(C.DEMO_EMPRESAS) if demo else [])
    specs = [
        ("tblEmpresas", ["Empresa", "Industria", "Dominios de correo", "Alias o nombres alternativos", "Prioridad", "Notas"],
         empresas, {"Empresa": 24, "Industria": 24, "Dominios de correo": 24,
                    "Alias o nombres alternativos": 34, "Prioridad": 10, "Notas": 26}),
        ("tblIndustrias", ["Industria"], [(x,) for x in C.INDUSTRIAS], {"Industria": 28}),
        ("tblEspecialidades", ["Especialidad"], [(x,) for x in C.ESPECIALIDADES], {"Especialidad": 34}),
        ("tblTiposOportunidad", ["Tipo de oportunidad"], [(x,) for x in C.TIPOS_OPORTUNIDAD], {"Tipo de oportunidad": 22}),
        ("tblEstados", ["Estado", "Orden", "Grupo", "Categoría de regla", "Activa", "Cuenta como postulada", "Descripción"],
         C.ESTADOS, {"Estado": 27, "Orden": 8, "Grupo": 15, "Categoría de regla": 13, "Activa": 8,
                     "Cuenta como postulada": 11, "Descripción": 48}),
        ("tblTiposComunicacion", ["Tipo de comunicación", "Estado propuesto por defecto", "Estados permitidos",
                                  "Cambia estado", "Sensible", "Descripción y ejemplos"],
         C.TIPOS_COMUNICACION, {"Tipo de comunicación": 27, "Estado propuesto por defecto": 27,
                                "Estados permitidos": 40, "Cambia estado": 9, "Sensible": 9,
                                "Descripción y ejemplos": 60}),
        ("tblParametros", ["Parámetro", "Valor", "Descripción"], C.PARAMETROS,
         {"Parámetro": 29, "Valor": 16, "Descripción": 60}),
        ("tblDominiosATS", ["Dominio", "Plataforma", "Relevante por dominio"], C.DOMINIOS_ATS,
         {"Dominio": 22, "Plataforma": 26, "Relevante por dominio": 12}),
        ("tblPalabrasClave", ["Palabra clave"], [(x,) for x in C.PALABRAS_CLAVE], {"Palabra clave": 24}),
        ("tblOpciones", ["Decisión", "Confirmación manual", "Origen del registro", "Dirección"],
         _zip_lists(DECISIONES, CONFIRMACIONES, ORIGENES, ["Recibido", "Enviado"]),
         {"Decisión": 28, "Confirmación manual": 26, "Origen del registro": 24, "Dirección": 11}),
    ]
    col = 1
    positions = {}
    for name, cols, rows, widths in specs:
        add_table(ws, name, cols, rows, hr, first_col=col, widths=widths)
        positions[name] = col
        ws.cell(row=hr - 1, column=col, value=name.replace("tbl", "")).font = font(11, True, NAVY)
        col += len(cols) + 1
    # Formato de parámetros
    pc = positions["tblParametros"]
    for i, (k, v, _) in enumerate(C.PARAMETROS):
        cell = ws.cell(row=hr + 1 + i, column=pc + 1)
        cell.fill = fill(INPUT_FILL)
        if isinstance(v, dt.date):
            cell.number_format = DATE_FMT
        elif isinstance(v, float):
            cell.number_format = "0.00"
    for row in ws.iter_rows(min_row=hr + 1, max_row=ws.max_row):
        for c in row:
            c.alignment = Alignment(vertical="top", wrap_text=True)
    ws.freeze_panes = "A5"
    return ws


def _zip_lists(*lists):
    n = max(len(x) for x in lists)
    return [tuple(x[i] if i < len(x) else None for x in lists) for i in range(n)]


def build_postulaciones(wb, demo):
    ws = wb.create_sheet("Postulaciones")
    sheet_title(ws, "Base de postulaciones",
                "Una fila por proceso (una empresa puede tener varias). Columnas grises = calculadas: no escribir en ellas. "
                "Power Automate actualiza Estado, Última comunicación, Próxima acción y fechas.", 25)
    rows = []
    for i, e in enumerate(C.EMPRESAS, start=1):
        rows.append((f"POST-{i:04d}", e[0], "Por verificar", e[1], "Por verificar", "Por verificar",
                     "Por verificar", None, "Por verificar", FECHA_CREACION, "", "Verificar si existe oportunidad, "
                     "postulación preparada o enviada", None, None, "", "", "",
                     "Registro inicial sin antecedentes: no se ha confirmado oportunidad, preparación ni envío.",
                     None, None, None, None, None, None, C.ORIGEN_INICIAL))
    if demo:
        for d in C.DEMO_POSTULACIONES:
            r = list(d[:13]) + [None] + list(d[13:17]) + [None] * 6 + [C.ORIGEN_DEMO]
            rows.append(tuple(r))
    hr = 4
    _, last = add_table(ws, "tblPostulaciones", POST_COLS, rows, hr, formulas=POST_FORMULAS,
                        widths=POST_WIDTHS,
                        date_cols={"Fecha de postulación", "Fecha de última actualización", "Fecha de seguimiento",
                                   "Mes de postulación"},
                        int_cols={"Días desde última comunicación", "N° entrevistas"})
    for c in POST_FORMULAS:
        L = col_letter(POST_COLS, c)
        for r in range(hr + 1, last + 1):
            ws[f"{L}{r}"].fill = fill("F0EFEC")
            ws[f"{L}{r}"].font = font(10, color=INK2)
    ws[f"{col_letter(POST_COLS, 'Mes de postulación')}{hr}"].comment = Comment(
        "Calculada. Primer día del mes de postulación; alimenta el gráfico mensual.", "Sistema")
    ws.freeze_panes = "C5"
    lastrow = 2000
    rng = lambda name: f"{col_letter(POST_COLS, name)}5:{col_letter(POST_COLS, name)}{lastrow}"  # noqa: E731
    dv_specs = [
        ("Empresa", "=nmEmpresas", "warning", "Empresa nueva",
         "La empresa no está en Catálogos. Puede continuar, pero agréguela en tblEmpresas para que aparezca en gráficos."),
        ("Industria", "=nmIndustrias", "warning", "Industria nueva", "Agregue la industria en Catálogos."),
        ("Especialidad", "=nmEspecialidades", "warning", "Especialidad nueva", "Agregue la especialidad en Catálogos."),
        ("Tipo de oportunidad", "=nmTiposOportunidad", "warning", "Tipo nuevo", "Agregue el tipo en Catálogos."),
        ("Estado actual", "=nmEstados", "stop", "Estado no válido",
         "Use un estado de Catálogos > Estados. Para crear uno nuevo, agréguelo primero allí."),
        ("Origen del registro", "=nmOrigenes", "warning", "Origen", "Seleccione un origen."),
    ]
    for col, formula, style, title, msg in dv_specs:
        dv = DataValidation(type="list", formula1=formula, allow_blank=True, errorStyle=style,
                            error=msg, errorTitle=title, showErrorMessage=True)
        dv.add(rng(col))
        ws.add_data_validation(dv)
    dvd = DataValidation(type="date", operator="greaterThan", formula1="36526", allow_blank=True,
                         error="Ingrese una fecha (dd-mm-aaaa).", showErrorMessage=True)
    for c in ("Fecha de postulación", "Fecha de seguimiento"):
        dvd.add(rng(c))
    ws.add_data_validation(dvd)

    # Formato condicional
    E = col_letter(POST_COLS, "Estado actual")
    G = col_letter(POST_COLS, "Grupo de estado")
    A = col_letter(POST_COLS, "Activa")
    M = col_letter(POST_COLS, "Fecha de seguimiento")
    W = col_letter(POST_COLS, "Alerta sin respuesta")
    O = col_letter(POST_COLS, "Origen del registro")
    est_rng = f"{E}5:{E}{lastrow}"
    for grp, color, fcolor in [("Oferta", "E3F4E3", "0A6B0A"), ("Rechazo", "FBE4E4", "9E2B2B"),
                               ("Selección", "E6EFFA", "1C4F8F"), ("Entrevista", "ECE8FB", "45369E"),
                               ("Evaluación", "E6EFFA", "1C4F8F"), ("Sin respuesta", "FFF3D6", "8A5A00"),
                               ("Por verificar", "F0EFEC", INK2)]:
        ws.conditional_formatting.add(est_rng, FormulaRule(formula=[f'${G}5="{grp}"'], fill=fill(color),
                                                           font=Font(name=FONT, color=fcolor, bold=True)))
    ws.conditional_formatting.add(f"{M}5:{M}{lastrow}", FormulaRule(
        formula=[f'AND(ISNUMBER(${M}5),${M}5<TODAY(),${G}5<>"Rechazo",${G}5<>"Cerrado")'],
        fill=fill("FBE4E4"), font=Font(name=FONT, color=CRIT, bold=True)))
    ws.conditional_formatting.add(f"{W}5:{W}{lastrow}", FormulaRule(
        formula=[f'${W}5="Sí"'], fill=fill("FFF3D6"), font=Font(name=FONT, color="8A5A00", bold=True)))
    ws.conditional_formatting.add(f"A5:{col_letter(POST_COLS, 'Observaciones')}{lastrow}", FormulaRule(
        formula=[f'${O}5="{C.ORIGEN_DEMO}"'], font=Font(name=FONT, color="6250D6", italic=True)))
    _ = A
    return ws


def build_comunicaciones(wb, demo):
    ws = wb.create_sheet("Comunicaciones")
    sheet_title(ws, "Historial de comunicaciones",
                "Un evento por correo. Nunca se sobrescribe ni se borra: las correcciones se registran en "
                "«Estado confirmado» y «Confirmación manual». La clave anti-duplicados es «ID del mensaje».", 24)
    rows = []
    if demo:
        for d in C.DEMO_COMUNICACIONES:
            r = list(d) + [d[3]]  # Fecha de registro = fecha del correo (demo)
            r = r[:22] + [r[22], None]
            rows.append(tuple(r))
    _, last = add_table(ws, "tblComunicaciones", COM_COLS, rows, 4, formulas=COM_FORMULAS, widths=COM_WIDTHS,
                        datetime_cols={"Fecha de recepción", "Fecha de registro"},
                        pct_cols={"Nivel de confianza de clasificación"}, int_cols={"Días de espera"})
    L = col_letter(COM_COLS, "Días de espera")
    for r in range(5, last + 1):
        ws[f"{L}{r}"].fill = fill("F0EFEC")
    ws.freeze_panes = "B5"
    lastrow = 5000
    conf = col_letter(COM_COLS, "Nivel de confianza de clasificación")
    cm = col_letter(COM_COLS, "Confirmación manual")
    ws.conditional_formatting.add(f"{conf}5:{conf}{lastrow}", FormulaRule(
        formula=[f'AND(ISNUMBER(${conf}5),${conf}5<pUmbralAuto)'],
        fill=fill("FFF3D6")))
    ws.conditional_formatting.add(f"{cm}5:{cm}{lastrow}", FormulaRule(
        formula=[f'${cm}5="Pendiente"'], fill=fill("FFF3D6"), font=Font(name=FONT, bold=True, color="8A5A00")))
    ws.conditional_formatting.add(f"A5:{get_column_letter(len(COM_COLS))}{lastrow}", FormulaRule(
        formula=[f'LEFT($A5,5)="DEMO-"'], font=Font(name=FONT, color="6250D6", italic=True)))
    for col, formula in [("Confirmación manual", "=nmConfirmaciones"), ("Tipo de comunicación", "=nmTiposComunicacion"),
                         ("Estado confirmado", "=nmEstados")]:
        dv = DataValidation(type="list", formula1=formula, allow_blank=True, errorStyle="stop", showErrorMessage=True)
        Lc = col_letter(COM_COLS, col)
        dv.add(f"{Lc}5:{Lc}{lastrow}")
        ws.add_data_validation(dv)
    return ws


def build_pendientes(wb, demo):
    ws = wb.create_sheet("Pendientes")
    sheet_title(ws, "Pendientes de revisión",
                "Correos que la clasificación no pudo resolver con certeza. Complete las columnas amarillas, elija una "
                "«Decisión» y ejecute el script ProcesarPendientes (Automatizar > ProcesarPendientes).", 23)
    rows = []
    if demo:
        for d in C.DEMO_PENDIENTES:
            rows.append(tuple(list(d) + [None, None, None, None, None, "Pendiente", "No", None, None]))
    widths = {c: 18 for c in PEND_COLS}
    widths.update({"ID del mensaje": 22, "Asunto": 30, "Resumen del mensaje": 40, "Motivo de revisión": 40,
                   "Resultado del procesamiento": 40, "Empresa confirmada": 24, "Cargo confirmado": 26,
                   "Tipo de comunicación confirmado": 26, "Estado confirmado": 26, "Decisión": 26,
                   "Estado sugerido": 24, "Tipo sugerido": 24, "Empresa sugerida": 24})
    _, last = add_table(ws, "tblPendientes", PEND_COLS, rows, 4, widths=widths,
                        datetime_cols={"Fecha de recepción", "Fecha de procesamiento"}, pct_cols={"Confianza"})
    lastrow = 3000
    for c in PEND_INPUT:
        L = col_letter(PEND_COLS, c)
        for r in range(5, max(last, 60) + 1):
            ws[f"{L}{r}"].fill = fill(INPUT_FILL)
    for col, formula, style in [("Empresa confirmada", "=nmEmpresas", "warning"),
                                ("ID de postulación confirmada", "=nmIDsPostulacion", "warning"),
                                ("Tipo de comunicación confirmado", "=nmTiposComunicacion", "stop"),
                                ("Estado confirmado", "=nmEstados", "stop"),
                                ("Decisión", "=nmDecisiones", "stop")]:
        dv = DataValidation(type="list", formula1=formula, allow_blank=True, errorStyle=style, showErrorMessage=True,
                            error="Seleccione un valor de la lista (Catálogos).")
        L = col_letter(PEND_COLS, col)
        dv.add(f"{L}5:{L}{lastrow}")
        ws.add_data_validation(dv)
    dec = col_letter(PEND_COLS, "Decisión")
    proc = col_letter(PEND_COLS, "Procesado")
    ws.conditional_formatting.add(f"A5:{get_column_letter(len(PEND_COLS))}{lastrow}", FormulaRule(
        formula=[f'${proc}5="Sí"'], font=Font(name=FONT, color=MUTED)))
    ws.conditional_formatting.add(f"{dec}5:{dec}{lastrow}", FormulaRule(
        formula=[f'AND(${dec}5="Pendiente",${proc}5<>"Sí")'], font=Font(name=FONT, bold=True, color="8A5A00")))
    ws.freeze_panes = "B5"
    ws["A3"] = ("Decisiones: Aplicar = vincula a la postulación confirmada y aplica el estado confirmado · "
                "Crear nueva postulación = crea el registro con Empresa/Cargo confirmados · "
                "Registrar sin cambiar estado = solo historial · Descartar = no relevante.")
    ws["A3"].font = font(9, color=INK2)
    return ws


def build_calculos(wb):
    ws = wb.create_sheet("Cálculos")
    sheet_title(ws, "Cálculos del dashboard (no editar)",
                "Hoja auxiliar: resúmenes con fórmulas que alimentan los gráficos y los filtros del Dashboard.", 10)
    hdr = lambda cell, text: (ws.__setitem__(cell, text), setattr(ws[cell], "font", font(10, True, NAVY)))  # noqa: E731
    N = 60
    # --- Listas para filtros --------------------------------------------------------------
    hdr("B4", "Filtro Industria")
    hdr("C4", "Filtro Tipo")
    hdr("D4", "Filtro Especialidad")
    for col, tbl in (("B", "tblIndustrias[Industria]"), ("C", "tblTiposOportunidad[Tipo de oportunidad]"),
                     ("D", "tblEspecialidades[Especialidad]")):
        ws[f"{col}5"] = "(Todas)"
        for k in range(1, N):
            ws[f"{col}{5 + k}"] = f'=IFERROR(INDEX({tbl},{k}),"")'
    # --- Criterios -----------------------------------------------------------------------
    hdr("F4", "Filtro")
    hdr("G4", "Selección")
    hdr("H4", "Criterio")
    for i, (lab, cell) in enumerate((("Industria", "Dashboard!$D$5"), ("Tipo de oportunidad", "Dashboard!$J$5"),
                                     ("Especialidad", "Dashboard!$O$5"))):
        r = 5 + i
        ws[f"F{r}"] = lab
        ws[f"G{r}"] = f'=IF({cell}="","(Todas)",{cell})'
        ws[f"H{r}"] = f'=IF(G{r}="(Todas)","<>§§",G{r})'
    ws["F9"] = "Última ejecución ActualizarDashboard"
    ws["G9"] = None
    ws["G9"].number_format = DATETIME_FMT

    F = (",tblPostulaciones[ID de postulación],\"<>\",tblPostulaciones[Industria],'Cálculos'!$H$5,"
         "tblPostulaciones[Tipo de oportunidad],'Cálculos'!$H$6,tblPostulaciones[Especialidad],'Cálculos'!$H$7")

    # --- Por etapa (J:K) -----------------------------------------------------------------
    hdr("J4", "Estado")
    hdr("K4", "Registros")
    for k in range(1, 41):
        r = 4 + k
        ws[f"J{r}"] = f'=IFERROR(INDEX(tblEstados[Estado],{k}),"")'
        ws[f"K{r}"] = f'=IF(J{r}="","",COUNTIFS(tblPostulaciones[Estado actual],J{r}{F}))'
    # --- Por empresa (N:P lista, R:S top 12) ---------------------------------------------
    hdr("N4", "Empresa (catálogo)")
    hdr("O4", "Registros")
    hdr("P4", "Puntaje orden")
    for k in range(1, 151):
        r = 4 + k
        ws[f"N{r}"] = f'=IFERROR(INDEX(tblEmpresas[Empresa],{k}),"")'
        ws[f"O{r}"] = f'=IF(N{r}="","",COUNTIFS(tblPostulaciones[Empresa],N{r}{F}))'
        ws[f"P{r}"] = f'=IF(N{r}="",-1,O{r}+(1000-{k})/100000)'
    hdr("R4", "Top empresas")
    hdr("S4", "Registros")
    for k in range(1, 13):
        r = 4 + k
        ws[f"R{r}"] = (f'=IF({k}>COUNTIF($N$5:$N$154,"?*"),"",INDEX($N$5:$N$154,'
                       f'MATCH(LARGE($P$5:$P$154,{k}),$P$5:$P$154,0)))')
        ws[f"S{r}"] = f'=IF(R{r}="","",INDEX($O$5:$O$154,MATCH(LARGE($P$5:$P$154,{k}),$P$5:$P$154,0)))'
    # --- Por industria (U:V) -------------------------------------------------------------
    hdr("U4", "Industria")
    hdr("V4", "Registros")
    for k in range(1, 41):
        r = 4 + k
        ws[f"U{r}"] = f'=IFERROR(INDEX(tblIndustrias[Industria],{k}),"")'
        ws[f"V{r}"] = f'=IF(U{r}="","",COUNTIFS(tblPostulaciones[Industria],U{r}{F}))'
    # --- Mensual (X:Z) -------------------------------------------------------------------
    hdr("X4", "Mes")
    hdr("Y4", "Etiqueta")
    hdr("Z4", "Postulaciones")
    meses = '"Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"'
    ws["X5"] = (f'=DATE(YEAR({PARAM.format("MesInicioDashboard")}),MONTH({PARAM.format("MesInicioDashboard")}),1)')
    for k in range(12):
        r = 5 + k
        if k:
            ws[f"X{r}"] = f"=DATE(YEAR(X{r - 1}),MONTH(X{r - 1})+1,1)"
        ws[f"X{r}"].number_format = DATE_FMT
        ws[f"Y{r}"] = f'=CHOOSE(MONTH(X{r}),{meses})&" "&RIGHT(YEAR(X{r}),2)'
        ws[f"Z{r}"] = f'=COUNTIFS(tblPostulaciones[Mes de postulación],X{r}{F})'
    # --- Especialidad (AB:AC) ------------------------------------------------------------
    hdr("AB4", "Especialidad")
    hdr("AC4", "Registros")
    for k in range(1, 21):
        r = 4 + k
        ws[f"AB{r}"] = f'=IFERROR(INDEX(tblEspecialidades[Especialidad],{k}),"")'
        ws[f"AC{r}"] = f'=IF(AB{r}="","",COUNTIFS(tblPostulaciones[Especialidad],AB{r}{F}))'
    # --- Resultados (AE:AF) --------------------------------------------------------------
    hdr("AE4", "Resultado")
    hdr("AF4", "Registros")
    res = [
        ("En curso", f'=COUNTIFS(tblPostulaciones[Activa],"Sí",tblPostulaciones[Alerta sin respuesta],"No",'
                     f'tblPostulaciones[Grupo de estado],"<>Oferta"{F})'),
        ("Oferta", f'=COUNTIFS(tblPostulaciones[Grupo de estado],"Oferta"{F})'),
        ("Rechazada", f'=COUNTIFS(tblPostulaciones[Grupo de estado],"Rechazo"{F})'),
        ("Sin respuesta", f'=COUNTIFS(tblPostulaciones[Alerta sin respuesta],"Sí",'
                          f'tblPostulaciones[Grupo de estado],"<>Oferta"{F})'),
        ("Cerrado", f'=COUNTIFS(tblPostulaciones[Grupo de estado],"Cerrado"{F})'),
        ("Pre-postulación", f'=COUNTIFS(tblPostulaciones[Grupo de estado],"Pre-postulación"{F})'),
        ("Por verificar", f'=COUNTIFS(tblPostulaciones[Grupo de estado],"Por verificar"{F})'),
    ]
    for i, (lab, f) in enumerate(res):
        ws[f"AE{5 + i}"] = lab
        ws[f"AF{5 + i}"] = f
    # --- Espera promedio (AH:AI) ---------------------------------------------------------
    hdr("AH4", "Estado alcanzado")
    hdr("AI4", "Días promedio")
    for k in range(1, 41):
        r = 4 + k
        ws[f"AH{r}"] = f"=J{r}"
        ws[f"AI{r}"] = (f'=IF(AH{r}="","",IFERROR(AVERAGEIFS(tblComunicaciones[Días de espera],'
                        f'tblComunicaciones[Estado confirmado],AH{r}),0))')
        ws[f"AI{r}"].number_format = "0.0"
    # --- Próximos seguimientos (AK:AM) ----------------------------------------------------
    hdr("AK4", "k")
    hdr("AL4", "Clave (fecha+fila)")
    hdr("AM4", "Fila")
    for k in range(1, 9):
        r = 4 + k
        ws[f"AK{r}"] = k
        f = (f'=IFERROR(SMALL(IF((tblPostulaciones[Grupo de estado]<>"Rechazo")*(tblPostulaciones[Grupo de estado]<>"Cerrado")'
             f'*ISNUMBER(tblPostulaciones[Fecha de seguimiento]),tblPostulaciones[Fecha de seguimiento]'
             f'+ROW(tblPostulaciones[Fecha de seguimiento])/1000000),AK{r}),"")')
        ws[f"AL{r}"] = ArrayFormula(f"AL{r}", f)
        ws[f"AM{r}"] = f'=IF(AL{r}="","",ROUND(MOD(AL{r},1)*1000000,0))'
    for row in ws.iter_rows(min_row=5, max_row=200):
        for c in row:
            c.font = font(9, color=INK2)
    ws.sheet_properties.tabColor = GREY
    return ws, F


def kpi_tile(ws, top, left_col, label, formula, caption, accent):
    L = get_column_letter(left_col)
    R = get_column_letter(left_col + 2)
    ws.merge_cells(f"{L}{top}:{R}{top}")
    ws.merge_cells(f"{L}{top + 1}:{R}{top + 2}")
    ws.merge_cells(f"{L}{top + 3}:{R}{top + 3}")
    ws[f"{L}{top}"] = label.upper()
    ws[f"{L}{top}"].font = font(8, True, INK2)
    ws[f"{L}{top}"].alignment = Alignment(horizontal="left", vertical="bottom", indent=1)
    ws[f"{L}{top + 1}"] = formula
    ws[f"{L}{top + 1}"].font = font(26, True, INK)
    ws[f"{L}{top + 1}"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws[f"{L}{top + 1}"].number_format = "0"
    ws[f"{L}{top + 3}"] = caption
    ws[f"{L}{top + 3}"].font = font(7, False, MUTED)
    ws[f"{L}{top + 3}"].alignment = Alignment(horizontal="left", vertical="top", indent=1, wrap_text=True)
    for r in range(top, top + 4):
        for c in range(left_col, left_col + 3):
            cell = ws.cell(row=r, column=c)
            cell.fill = fill("FFFFFF")
            cell.border = Border(
                top=Side(style="thick", color=accent) if r == top else None,
                bottom=THIN if r == top + 3 else None,
                left=THIN if c == left_col else None,
                right=THIN if c == left_col + 2 else None,
            )


def _axis_fix(chart):
    for ax in (chart.x_axis, chart.y_axis):
        ax.delete = False
        ax.majorGridlines = None
        ax.txPr = None
    chart.legend = None


def _labels():
    dl = DataLabelList()
    dl.showVal = True
    dl.showCatName = False
    dl.showSerName = False
    dl.showLegendKey = False
    dl.showPercent = False
    dl.showLeaderLines = False
    return dl


def bar_chart(title, cats, vals, horizontal=True, color=BLUE, number_format="0"):
    ch = BarChart()
    ch.type = "bar" if horizontal else "col"
    ch.title = title
    ch.style = 2
    ch.add_data(vals, titles_from_data=False)
    ch.set_categories(cats)
    s = ch.series[0]
    s.graphicalProperties = GraphicalProperties(solidFill=color)
    s.graphicalProperties.line.noFill = True
    ch.gapWidth = 60
    _axis_fix(ch)
    if horizontal:
        ch.x_axis.scaling.orientation = "maxMin"
    ch.dataLabels = _labels()
    ch.y_axis.number_format = number_format
    ch.y_axis.delete = True
    ch.width = 15.8
    ch.height = 8.2
    return ch


def doughnut(title, cats, vals, colors):
    ch = DoughnutChart(holeSize=58)
    ch.title = title
    ch.add_data(vals, titles_from_data=False)
    ch.set_categories(cats)
    s = ch.series[0]
    s.dPt = []
    for i, colr in enumerate(colors):
        pt = DataPoint(idx=i)
        pt.graphicalProperties = GraphicalProperties(solidFill=colr)
        pt.graphicalProperties.line.solidFill = "FFFFFF"
        s.dPt.append(pt)
    ch.dataLabels = _labels()
    ch.legend.position = "r"
    ch.width = 15.8
    ch.height = 8.2
    return ch


def build_dashboard(wb, calc, F):
    ws = wb["Dashboard"]
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 90
    widths = [2] + [9.5, 9.5, 9.5, 2.2] * 5
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for r in range(1, 140):
        for c in range(1, 22):
            ws.cell(row=r, column=c).fill = fill(SURFACE)
    # Encabezado
    for c in range(1, 22):
        for r in (1, 2, 3):
            ws.cell(row=r, column=c).fill = fill(NAVY)
    ws["B2"] = "Seguimiento de Postulaciones 2027"
    ws["B2"].font = font(18, True, "FFFFFF")
    ws["B3"] = ("Prácticas, memorias, pasantías e internships · Ing. Civil Industrial + Minas (UAI) · "
                "Disponibilidad para memoria desde enero 2027")
    ws["B3"].font = font(9, False, "C9D3E6")
    ws.row_dimensions[2].height = 30
    ws["Q2"] = "=TODAY()"
    ws["Q2"].number_format = '"Hoy: "dd-mm-yyyy'
    ws["Q2"].font = font(9, True, "FFFFFF")
    ws["Q2"].alignment = Alignment(horizontal="right")
    ws.merge_cells("Q2:T2")
    ws["Q3"] = "='Cálculos'!G9"
    ws["Q3"].number_format = '"Actualizado por script: "dd-mm-yyyy hh:mm;;"Script de actualización aún no ejecutado"'
    ws["Q3"].font = font(8, False, "C9D3E6")
    ws["Q3"].alignment = Alignment(horizontal="right")
    ws.merge_cells("Q3:T3")
    # Filtros (etiqueta, rango etiqueta, celda valor, rango valor, nombre de lista)
    for lab, lrng, cell, vrng, name in (("Industria", "B5:C5", "D5", "D5:F5", "nmFiltroIndustria"),
                                        ("Tipo de oportunidad", "G5:I5", "J5", "J5:L5", "nmFiltroTipo"),
                                        ("Especialidad", "M5:N5", "O5", "O5:Q5", "nmFiltroEspecialidad")):
        first = lrng.split(":")[0]
        ws[first] = lab
        ws[first].font = font(9, True, INK2)
        ws[first].alignment = Alignment(horizontal="right", vertical="center", indent=1)
        ws.merge_cells(lrng)
        ws[cell] = "(Todas)"
        ws[cell].font = font(10, True, BLUE)
        ws[cell].alignment = Alignment(horizontal="left", vertical="center", indent=1)
        for row in ws[vrng]:
            for c in row:
                c.fill = fill(INPUT_FILL)
                c.border = Border(bottom=Side(style="thin", color=BLUE))
        ws.merge_cells(vrng)
        dv = DataValidation(type="list", formula1=f"={name}", allow_blank=False, showErrorMessage=True,
                            error="Seleccione un valor de la lista.")
        dv.add(cell)
        ws.add_data_validation(dv)
    ws.merge_cells("R5:T5")
    ws["R5"] = "Filtros aplican a indicadores y gráficos (salvo espera promedio)"
    ws["R5"].font = font(7, False, MUTED, italic=True)
    ws["R5"].alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[5].height = 22

    def cnt(crit):
        return f"=COUNTIFS({crit}{F})"

    sel = ("(('Cálculos'!$G$5=\"(Todas)\")+(tblPostulaciones[Industria]='Cálculos'!$G$5)>0)"
           "*(('Cálculos'!$G$6=\"(Todas)\")+(tblPostulaciones[Tipo de oportunidad]='Cálculos'!$G$6)>0)"
           "*(('Cálculos'!$G$7=\"(Todas)\")+(tblPostulaciones[Especialidad]='Cálculos'!$G$7)>0)")
    m = f'(tblPostulaciones[ID de postulación]<>"")*(tblPostulaciones[Empresa]<>"")*{sel}'
    empresas = (f"=SUMPRODUCT({m}/(COUNTIFS(tblPostulaciones[Empresa],tblPostulaciones[Empresa]&\"\"{F})+({m}=0)))")

    k1 = [
        ("Postulaciones enviadas", cnt('tblPostulaciones[Cuenta como postulada],"Sí"'),
         "Estados desde «Postulación enviada» en adelante", BLUE),
        ("Empresas", empresas, "Empresas distintas en la base (todos los estados)", BLUE),
        ("Postulaciones activas", cnt('tblPostulaciones[Activa],"Sí"'), "Procesos abiertos según Catálogos > Estados", BLUE),
        ("En revisión", cnt('tblPostulaciones[Grupo de estado],"En revisión"'), "Recibidas o en revisión curricular", BLUE),
        ("En etapa de selección",
         f'=SUMPRODUCT(COUNTIFS(tblPostulaciones[Grupo de estado],{{"Selección","Evaluación","Entrevista"}}{F}))',
         "Etapas, evaluaciones y entrevistas en curso", BLUE),
    ]
    k2 = [
        ("Entrevistas", f"=SUMIFS(tblPostulaciones[N° entrevistas]{F})", "Invitaciones a entrevista registradas", "6250D6"),
        ("Evaluaciones pendientes", cnt('tblPostulaciones[Grupo de estado],"Evaluación"'),
         "Evaluación online, técnica/grupal o exámenes", "6250D6"),
        ("Ofertas recibidas", cnt('tblPostulaciones[Grupo de estado],"Oferta"'), "Oferta recibida o seleccionada", GOOD),
        ("Rechazos", cnt('tblPostulaciones[Grupo de estado],"Rechazo"'), "Procesos con rechazo informado", CRIT),
        ("Sin respuesta", cnt('tblPostulaciones[Alerta sin respuesta],"Sí"'),
         "Estado «Sin respuesta» o sin correos en el plazo de Parámetros", WARN),
    ]
    for i, (lab, f, cap, acc) in enumerate(k1):
        kpi_tile(ws, 7, 2 + i * 4, lab, f, cap, acc)
    for i, (lab, f, cap, acc) in enumerate(k2):
        kpi_tile(ws, 12, 2 + i * 4, lab, f, cap, acc)
    for r in (8, 9, 13, 14):
        ws.row_dimensions[r].height = 19
    for r in (10, 15):
        ws.row_dimensions[r].height = 22
    # Línea de alertas
    ws["B17"] = "Pendientes de revisión:"
    ws["E17"] = '=COUNTIFS(tblPendientes[ID del mensaje],"<>",tblPendientes[Procesado],"<>Sí")'
    ws["G17"] = "Seguimientos vencidos:"
    ws["J17"] = ('=COUNTIFS(tblPostulaciones[Fecha de seguimiento],"<"&TODAY(),tblPostulaciones[Grupo de estado],'
                 '"<>Rechazo",tblPostulaciones[Grupo de estado],"<>Cerrado")')
    ws["L17"] = "Registros por verificar:"
    ws["O17"] = cnt('tblPostulaciones[Grupo de estado],"Por verificar"')
    for c in ("B17", "G17", "L17"):
        ws[c].font = font(9, True, INK2)
    for c in ("E17", "J17", "O17"):
        ws[c].font = font(11, True, CRIT)
        ws[c].alignment = Alignment(horizontal="left")

    # Gráficos
    def ref(c1, r1, c2, r2):
        return Reference(calc, min_col=c1, min_row=r1, max_col=c2, max_row=r2)

    n_est = len(C.ESTADOS)
    n_ind = len(C.INDUSTRIAS)
    n_esp = len(C.ESPECIALIDADES)
    charts = [
        bar_chart("1. Registros por etapa (estado actual)", ref(10, 5, 10, 4 + n_est), ref(11, 5, 11, 4 + n_est)),
        bar_chart("2. Registros por empresa (top 12)", ref(18, 5, 18, 16), ref(19, 5, 19, 16)),
        bar_chart("3. Registros por industria", ref(21, 5, 21, 4 + n_ind), ref(22, 5, 22, 4 + n_ind)),
        bar_chart("4. Evolución mensual de postulaciones", ref(25, 5, 25, 16), ref(26, 5, 26, 16), horizontal=False),
        doughnut("5. Distribución por especialidad", ref(28, 5, 28, 4 + n_esp), ref(29, 5, 29, 4 + n_esp),
                 CAT[:n_esp - 1] + [GREY]),
        doughnut("6. Resultados de los procesos", ref(31, 5, 31, 11), ref(32, 5, 32, 11),
                 [BLUE, GOOD, CRIT, WARN, "5F5E5A", "CFCEC9", GREY]),
        bar_chart("7. Días promedio de espera hasta cada etapa", ref(34, 5, 34, 4 + n_est), ref(35, 5, 35, 4 + n_est),
                  color="6250D6", number_format="0.0"),
    ]
    anchors = ["B19", "L19", "B37", "L37", "B55", "L55", "B73"]
    for ch, a in zip(charts, anchors):
        ws.add_chart(ch, a)

    # Próximos seguimientos
    ws["L73"] = "Próximos seguimientos"
    ws["L73"].font = font(12, True, NAVY)
    heads = [("L", "Fecha"), ("N", "Empresa"), ("Q", "Próxima acción")]
    for col, h in heads:
        ws[f"{col}75"] = h
        ws[f"{col}75"].font = font(8, True, INK2)
        ws[f"{col}75"].border = Border(bottom=THIN)
    ws.merge_cells("L75:M75")
    ws.merge_cells("N75:P75")
    ws.merge_cells("Q75:T75")
    for k in range(8):
        r = 76 + k
        cr = 5 + k
        rowref = f"'Cálculos'!$AM${cr}"
        ws[f"L{r}"] = f'=IF({rowref}="","",INDEX(Postulaciones!$M:$M,{rowref}))'
        ws[f"L{r}"].number_format = DATE_FMT
        ws[f"N{r}"] = (f'=IF({rowref}="","",INDEX(Postulaciones!$B:$B,{rowref})&" · "&'
                       f'INDEX(Postulaciones!$C:$C,{rowref}))')
        ws[f"Q{r}"] = f'=IF({rowref}="","",INDEX(Postulaciones!$L:$L,{rowref}))'
        ws.merge_cells(f"L{r}:M{r}")
        ws.merge_cells(f"N{r}:P{r}")
        ws.merge_cells(f"Q{r}:T{r}")
        for col in ("L", "N", "Q"):
            ws[f"{col}{r}"].font = font(9, False, INK)
            ws[f"{col}{r}"].border = Border(bottom=THIN)
            ws[f"{col}{r}"].alignment = Alignment(vertical="center", shrink_to_fit=True)
    ws.conditional_formatting.add("L76:M83", FormulaRule(formula=['AND(ISNUMBER($L76),$L76<TODAY())'],
                                                         font=Font(name=FONT, color=CRIT, bold=True)))
    ws["L85"] = ("Gráficos 1-3 y 5 cuentan todos los registros de la base; los indicadores «enviadas» y el gráfico "
                 "mensual solo cuentan postulaciones con evidencia. Para tablas dinámicas y segmentadores ejecute "
                 "el script CrearAnalisisDinamico (hoja «Análisis»).")
    ws["L85"].font = font(8, False, MUTED, italic=True)
    ws["L85"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells("L85:T88")
    ws.sheet_properties.tabColor = NAVY
    return ws


def build_guia(wb, demo):
    ws = wb.create_sheet("Guía")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 120
    lines = [
        ("Guía rápida", "h"),
        ("Este archivo es la base de datos y el dashboard del sistema. La automatización (Power Automate + Office Scripts + "
         "Claude API) está documentada en la carpeta docs/ del repositorio.", ""),
        ("Qué puede editar", "h2"),
        ("• Celdas amarillas: filtros del Dashboard, parámetros (Catálogos > Parámetros) y confirmaciones en Pendientes.", ""),
        ("• Postulaciones: puede agregar filas a mano (use un ID nuevo, p. ej. POST-0016) y editar columnas blancas.", ""),
        ("• Catálogos: agregue empresas (con dominios de correo y alias), industrias, especialidades, tipos y estados "
         "al final de cada tabla.", ""),
        ("Qué NO debe editar", "h2"),
        ("• Columnas grises de Postulaciones y Comunicaciones (son fórmulas), la hoja Cálculos y los nombres de columnas.", ""),
        ("• No borre filas de Comunicaciones: el historial es la trazabilidad. Para anular un evento marque "
         "«Confirmación manual» = Descartada.", ""),
        ("Ejemplo de fila válida (Postulaciones)", "h2"),
        ("POST-0016 | Codelco | Memoria Planificación Minera 2027 | Minería | Ingeniería Civil en Minas | Santiago | "
         "Memoria de título | 05-11-2026 | Postulación enviada | 05-11-2026 | … | Hacer seguimiento | 19-11-2026", ""),
        ("Scripts (pestaña Automatizar en Excel para la Web)", "h2"),
        ("• ActualizarDashboard: recalcula, ajusta rangos de gráficos, refresca tablas dinámicas y (opcional) marca "
         "«Sin respuesta».", ""),
        ("• ProcesarPendientes: aplica las decisiones tomadas en la hoja Pendientes.", ""),
        ("• CrearAnalisisDinamico: crea la hoja Análisis con tablas dinámicas y segmentadores (ejecutar una vez).", ""),
        ("• RegistrarCorreo / EvaluarRelevancia: los usa Power Automate; no requieren ejecución manual.", ""),
        ("• LimpiarDatosFicticios: elimina filas DEMO (solo archivo DEMO).", ""),
        ("Reglas de confiabilidad", "h2"),
        ("• «Postulación enviada» solo se asigna con evidencia de envío (correo enviado por usted o confirmación manual). "
         "Una confirmación de la empresa registra «Postulación recibida».", ""),
        ("• Avance, Oferta y Rechazo requieren confianza ≥ UmbralConfianzaSensible y señales textuales explícitas; si no, "
         "van a Pendientes.", ""),
        ("• Un correo de recepción o revisión nunca hace retroceder un proceso que ya está en selección.", ""),
        ("• Cambios sobre procesos finalizados (oferta, seleccionada, rechazada, cerrado) siempre requieren revisión manual.", ""),
    ]
    if demo:
        lines.insert(1, ("ARCHIVO DE PRUEBA: contiene datos ficticios (prefijo [DEMO] / DEMO-, texto morado en cursiva). "
                         "No usar como archivo real.", "warn"))
    r = 2
    for text, kind in lines:
        ws[f"B{r}"] = text
        if kind == "h":
            ws[f"B{r}"].font = font(16, True, NAVY)
        elif kind == "h2":
            ws[f"B{r}"].font = font(11, True, NAVY)
            r += 0
        elif kind == "warn":
            ws[f"B{r}"].font = font(10, True, "6250D6")
            ws[f"B{r}"].fill = fill("ECE8FB")
        else:
            ws[f"B{r}"].font = font(10, False, INK)
        ws[f"B{r}"].alignment = Alignment(wrap_text=True, vertical="top")
        r += 1 if kind != "h2" else 1
    return ws


def defined_names(wb):
    names = {
        "nmEmpresas": "tblEmpresas[Empresa]",
        "nmIndustrias": "tblIndustrias[Industria]",
        "nmEspecialidades": "tblEspecialidades[Especialidad]",
        "nmTiposOportunidad": "tblTiposOportunidad[Tipo de oportunidad]",
        "nmEstados": "tblEstados[Estado]",
        "nmTiposComunicacion": "tblTiposComunicacion[Tipo de comunicación]",
        "nmDecisiones": "tblOpciones[Decisión]",
        "nmConfirmaciones": "tblOpciones[Confirmación manual]",
        "nmOrigenes": "tblOpciones[Origen del registro]",
        "nmIDsPostulacion": "tblPostulaciones[ID de postulación]",
        # Los formatos condicionales no admiten referencias estructuradas: se usa un nombre.
        "pUmbralAuto": PARAM.format("UmbralConfianzaAutomatica"),
        "nmFiltroIndustria": "OFFSET('Cálculos'!$B$5,0,0,COUNTIF('Cálculos'!$B$5:$B$64,\"?*\"),1)",
        "nmFiltroTipo": "OFFSET('Cálculos'!$C$5,0,0,COUNTIF('Cálculos'!$C$5:$C$64,\"?*\"),1)",
        "nmFiltroEspecialidad": "OFFSET('Cálculos'!$D$5,0,0,COUNTIF('Cálculos'!$D$5:$D$64,\"?*\"),1)",
    }
    for k, v in names.items():
        wb.defined_names[k] = DefinedName(k, attr_text=v)


def build(demo: bool, path: str):
    wb = Workbook()
    wb.active.title = "Dashboard"
    build_postulaciones(wb, demo)
    build_comunicaciones(wb, demo)
    build_catalogos(wb, demo)
    build_pendientes(wb, demo)
    calc, F = build_calculos(wb)
    build_guia(wb, demo)
    build_dashboard(wb, calc, F)
    defined_names(wb)
    # Orden de hojas: Dashboard, Postulaciones, Comunicaciones, Catálogos, Pendientes, Guía, Cálculos
    order = ["Dashboard", "Postulaciones", "Comunicaciones", "Catálogos", "Pendientes", "Guía", "Cálculos"]
    wb._sheets = [wb[n] for n in order]
    wb.active = 0
    wb.calculation.fullCalcOnLoad = True
    wb.properties.title = "Seguimiento de Postulaciones 2027"
    wb.properties.creator = "Sistema de seguimiento de postulaciones"
    wb.save(path)
    return path


if __name__ == "__main__":
    root = os.path.join(os.path.dirname(__file__), "..", "excel")
    os.makedirs(root, exist_ok=True)
    out = sys.argv[1] if len(sys.argv) > 1 else root
    p1 = build(False, os.path.join(out, "Seguimiento_Postulaciones_2027.xlsx"))
    p2 = build(True, os.path.join(out, "Seguimiento_Postulaciones_2027_DEMO.xlsx"))
    print(p1)
    print(p2)
