"""Builds modelo/modelo_calculo.xlsx.

Step 0: only the "Entradas" and "Registro" sheets.
Every input gets a defined name so later sheets can use readable formulas
(e.g. =DAU*REQ_POR_USUARIO).
"""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.workbook.defined_name import DefinedName

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "modelo" / "modelo_calculo.xlsx"

HEADER_FILL = PatternFill("solid", fgColor="1F3864")
HEADER_FONT = Font(bold=True, color="FFFFFF")
SECTION_FILL = PatternFill("solid", fgColor="D9E1F2")
SECTION_FONT = Font(bold=True)
INPUT_FONT = Font(color="0000FF")  # blue = value typed by hand
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")

SRC = "Enunciado §{}"

# (name, label, value, unit, number_format, source, classification, note)
INPUTS = [
    ("General", None),
    ("RESPONSABLE", "Persona responsable de verificar", "Juliana Briceño", "-", "@",
     "Decisión del equipo", "Documentado", "Trabajo individual"),
    ("Demanda", None),
    ("DAU", "Usuarios activos diarios", 200_000, "usuarios/día", "#,##0",
     SRC.format("4.2"), "Documentado", "Escenario hipotético común"),
    ("REQ_POR_USUARIO", "Peticiones promedio por usuario", 1.5, "peticiones/usuario/día", "0.0",
     SRC.format("4.2"), "Documentado", ""),
    ("FACTOR_PICO", "Factor de pico sobre el promedio diario", 6, "x", "0",
     SRC.format("4.2"), "Documentado", ""),
    ("TOK_ENTRADA", "Tokens de entrada promedio por petición", 4_000, "tokens", "#,##0",
     SRC.format("4.2"), "Documentado", "Incluye instrucciones de sistema, historial y documentos"),
    ("TOK_SALIDA", "Tokens de salida promedio por petición", 500, "tokens", "#,##0",
     SRC.format("4.2"), "Documentado", ""),
    ("CONTEXTO_MAX", "Ventana de contexto máxima (entrada + salida)", 8_192, "tokens", "#,##0",
     SRC.format("4.2"), "Documentado", ""),
    ("SALIDA_MAX", "Salida máxima por petición", 1_000, "tokens", "#,##0",
     SRC.format("4.2"), "Documentado", ""),
    ("FACTOR_CRECIMIENTO", "Escenario de estrés: multiplicador de demanda", 2, "x", "0",
     SRC.format("4.2"), "Documentado", "Se presenta aparte del TCO base"),
    ("Objetivos de servicio (perfil controlado)", None),
    ("SLO_TTFT", "TTFT p95 máximo", 2, "s", "0.0",
     SRC.format("4.3"), "Documentado", "Medido desde el cliente"),
    ("SLO_TPOT", "Tiempo por token de salida p95 máximo", 40, "ms/token", "0",
     SRC.format("4.3"), "Documentado", "Promedio por petición, evaluado en p95"),
    ("SLO_TOTAL", "Tiempo total p95 para 500 tokens de salida", 30, "s", "0",
     SRC.format("4.3"), "Documentado", ""),
    ("SLO_DISPONIBILIDAD", "Disponibilidad mensual objetivo", 0.999, "%", "0.0%",
     SRC.format("4.3"), "Documentado", ""),
    ("SLO_ERRORES", "Tasa máxima de errores", 0.001, "% de peticiones válidas", "0.0%",
     SRC.format("4.3"), "Documentado", ""),
    ("NODOS_PERDIDOS", "Nodos de cómputo que se pueden perder sin bajar del pico", 1, "nodos", "0",
     SRC.format("4.3"), "Documentado", ""),
    ("Calidad", None),
    ("CALIDAD_MIN", "Calidad mínima frente al modelo de referencia", 0.90, "%", "0%",
     SRC.format("4.4"), "Documentado", "Referencia: Llama 3.1 70B Instruct BF16"),
    ("TAREAS_EVAL", "Tareas de evaluación en español", 100, "tareas", "0",
     SRC.format("4.4"), "Documentado", "No se entregó el conjunto: queda como plan de evaluación"),
    ("Presupuesto (USD, valores hipotéticos)", None),
    ("CAPEX_MAX", "CAPEX máximo", 8_000_000, "USD", "#,##0",
     SRC.format("4.5"), "Documentado", ""),
    ("OPEX_MAX", "OPEX anual máximo", 1_500_000, "USD/año", "#,##0",
     SRC.format("4.5"), "Documentado", ""),
    ("TCO_MAX", "TCO máximo a 3 años", 12_500_000, "USD", "#,##0",
     SRC.format("4.5"), "Documentado", ""),
    ("ANIOS_TCO", "Horizonte del TCO", 3, "años", "0",
     SRC.format("4.5"), "Documentado", "Demanda base constante"),
    ("Constantes y convenciones", None),
    ("SEG_DIA", "Segundos por día", 86_400, "s/día", "#,##0",
     "Definición", "Documentado", ""),
    ("HORAS_ANIO", "Horas por año", 8_760, "h/año", "#,##0",
     "Definición (365 días)", "Documentado", ""),
    ("BYTES_GB", "Bytes por GB (decimal)", 1_000_000_000, "bytes/GB", "#,##0",
     "Convención adoptada", "Supuesto", "Se usa GB decimal en todo el modelo; 1 GiB = 1,0737 GB"),
]

# (decision, value_unit, source, uncertainty, classification)
LOG = [
    ("Escenario de demanda base", "200.000 usuarios/día × 1,5 peticiones",
     "Enunciado §4.2", "Ninguna: es el escenario común de comparación", "Documentado"),
    ("Perfil de tokens controlado", "4.000 entrada / 500 salida",
     "Enunciado §4.2", "La mezcla real de longitudes variará; se evalúa aparte", "Documentado"),
    ("Modelo de referencia de calidad", "Llama 3.1 70B Instruct, BF16",
     "Enunciado §4.4", "Revisión exacta por congelar", "Documentado"),
    ("Evaluación de calidad", "Plan de evaluación pendiente, sin puntajes",
     "Enunciado §4.4 (no se entregó el conjunto de 100 tareas)",
     "Alta: la calidad no se mide en este trabajo", "Supuesto"),
    ("Evidencia de rendimiento", "Benchmarks públicos + supuestos con sensibilidad",
     "Enunciado §8.3 (no hay acceso a GPUs)",
     "Alta hasta validar con una medición propia", "Supuesto"),
    ("Unidad de memoria", "GB decimal (10^9 bytes)",
     "Convención adoptada", "Ninguna, solo hay que ser consistente", "Supuesto"),
]


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = WRAP
        cell.border = BORDER


def build_inputs(wb):
    ws = wb.active
    ws.title = "Entradas"
    ws["A1"] = "Entradas del modelo"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = "Texto azul = valor escrito a mano. Todo lo demás del libro se calcula a partir de esta hoja."
    ws["A2"].font = Font(italic=True, color="595959")

    headers = ["Nombre", "Parámetro", "Valor", "Unidad", "Fuente", "Clasificación", "Nota"]
    ws.append([])
    ws.append(headers)
    style_header(ws, ws.max_row, len(headers))

    for item in INPUTS:
        if item[1] is None:
            ws.append([item[0]])
            r = ws.max_row
            for c in range(1, len(headers) + 1):
                ws.cell(row=r, column=c).fill = SECTION_FILL
                ws.cell(row=r, column=c).font = SECTION_FONT
            continue
        name, label, value, unit, fmt, source, klass, note = item
        ws.append([name, label, value, unit, source, klass, note])
        r = ws.max_row
        vcell = ws.cell(row=r, column=3)
        vcell.font = INPUT_FONT
        vcell.number_format = fmt
        for c in range(1, len(headers) + 1):
            ws.cell(row=r, column=c).border = BORDER
            ws.cell(row=r, column=c).alignment = WRAP
        wb.defined_names[name] = DefinedName(name, attr_text=f"Entradas!$C${r}")

    for col, width in zip("ABCDEFG", [22, 46, 16, 22, 26, 15, 48]):
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A5"


def build_log(wb):
    ws = wb.create_sheet("Registro")
    ws["A1"] = "Registro de decisiones y datos"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = ("Clasificación: Documentado (fuente identificable), Medido (experimento reproducible), "
                "Calculado (fórmulas visibles), Supuesto (requiere validación). "
                "Una cifra dada por una IA sigue siendo Supuesto hasta verificarla.")
    ws["A2"].font = Font(italic=True, color="595959")
    ws.merge_cells("A2:G2")
    ws.row_dimensions[2].height = 32
    ws["A2"].alignment = WRAP

    headers = ["#", "Decisión o dato", "Valor y unidad", "Fuente o cálculo",
               "Incertidumbre", "Responsable de verificar", "Clasificación"]
    ws.append([])
    ws.append(headers)
    style_header(ws, ws.max_row, len(headers))

    for i, (decision, value, source, unc, klass) in enumerate(LOG, start=1):
        ws.append([i, decision, value, source, unc, "=RESPONSABLE", klass])
        r = ws.max_row
        for c in range(1, len(headers) + 1):
            ws.cell(row=r, column=c).border = BORDER
            ws.cell(row=r, column=c).alignment = WRAP

    for col, width in zip("ABCDEFG", [5, 34, 30, 36, 36, 22, 15]):
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A5"


def main():
    wb = Workbook()
    build_inputs(wb)
    build_log(wb)
    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT)
    print(f"Saved {OUT.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
