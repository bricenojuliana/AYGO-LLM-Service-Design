"""Builds modelo/modelo_calculo.xlsx.

Sheets so far: Entradas, Demanda, Registro.
Every input and key result gets a defined name so formulas stay readable
(e.g. =DAU*REQ_POR_USUARIO). Values typed by hand are blue.
"""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.workbook.properties import CalcProperties

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "modelo" / "modelo_calculo.xlsx"

HEADER_FILL = PatternFill("solid", fgColor="1F3864")
HEADER_FONT = Font(bold=True, color="FFFFFF")
SECTION_FILL = PatternFill("solid", fgColor="D9E1F2")
SECTION_FONT = Font(bold=True)
INPUT_FONT = Font(color="0000FF")  # blue = value typed by hand
NOTE_FONT = Font(italic=True, color="595959")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")

SRC = "Enunciado §{}"
CFG = "config.json de Meta-Llama-3.1-70B-Instruct (espejo NousResearch, HF)"

# ---------------------------------------------------------------- Entradas
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
    ("TIEMPO_SISTEMA", "Tiempo promedio de una petición en el sistema", 20, "s", "0.0",
     SRC.format("6.3") + " (valor de ejemplo)", "Supuesto",
     "Incluye espera, prefill y generación. Se reemplaza cuando haya evidencia de rendimiento (paso 3)"),
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
    ("Modelo: Llama 3.1 70B Instruct (config.json)", None),
    ("CAPAS", "Capas (num_hidden_layers)", 80, "capas", "0",
     CFG, "Documentado", "Cotejar con el artefacto exacto que se use"),
    ("ATT_HEADS", "Cabezas de atención (num_attention_heads)", 64, "cabezas", "0",
     CFG, "Documentado", ""),
    ("KV_HEADS", "Cabezas KV (num_key_value_heads)", 8, "cabezas", "0",
     CFG, "Documentado", "No es lo mismo que las cabezas de atención (GQA)"),
    ("HIDDEN", "Dimensión oculta (hidden_size)", 8_192, "-", "#,##0",
     CFG, "Documentado", ""),
    ("INTERMEDIA", "Dimensión intermedia del MLP (intermediate_size)", 28_672, "-", "#,##0",
     CFG, "Documentado", ""),
    ("VOCAB", "Tamaño del vocabulario (vocab_size)", 128_256, "tokens", "#,##0",
     CFG, "Documentado", ""),
    ("Hardware y memoria", None),
    ("GPU_MEM", "Memoria por GPU (H100 SXM)", 80, "GB", "0",
     "NVIDIA, página de producto H100", "Documentado", "La guía de la DGX H100 da 640 GB para 8 GPUs"),
    ("H100_BW", "Ancho de banda de memoria por GPU (H100 SXM)", 3.35, "TB/s", "0.00",
     "NVIDIA, página de producto H100", "Documentado", ""),
    ("H100_BF16_SPARSE", "Cómputo BF16 por GPU, con dispersión (sparsity)", 1_979, "TFLOPS", "#,##0",
     "NVIDIA, página de producto H100", "Documentado", "La página marca este valor \"With sparsity\""),
    ("FACTOR_DISPERSION", "Factor de la dispersión 2:4 sobre el valor denso", 2, "x", "0",
     "Convención de NVIDIA para sparsity 2:4", "Supuesto",
     "La página no publica el valor denso; un modelo denso como Llama no aprovecha la dispersión"),
    ("Rendimiento (dato sintético del enunciado)", None),
    ("GPUS_REPLICA", "GPUs por réplica del diseño", 4, "GPUs", "0",
     SRC.format("8.4"), "Documentado", "Modelo BF16; la versión FP8 queda como variante a futuro"),
    ("CAP_REPLICA", "Peticiones por segundo que atiende una réplica cumpliendo los SLOs", 2, "peticiones/s", "0.00",
     SRC.format("8.4") + " (sintético)", "Supuesto",
     "No es un benchmark real; perfil 4.000 entrada / 500 salida"),
    ("CAIDA_SENS", "Caída de rendimiento para la sensibilidad", 0.30, "%", "0%",
     SRC.format("14.4"), "Documentado", ""),
    ("Réplicas, nodos y holgura", None),
    ("UTIL_DISENO", "Factor de utilización de diseño", 0.70, "%", "0%",
     SRC.format("9") + " (ejemplo)", "Supuesto",
     "Carga admitida frente a la capacidad del punto de operación; no es el % de GPU del monitoreo"),
    ("GPUS_NODO", "GPUs por nodo (servidor DGX H100)", 8, "GPUs", "0",
     "Guía de usuario DGX H100/H200 de NVIDIA", "Documentado", ""),
    ("NODOS_POR_RACK", "Nodos de cómputo por rack", 3, "nodos", "0",
     SRC.format("12.2") + " (ejemplo)", "Supuesto", "Se revisa con energía y refrigeración en el paso 7"),
    ("USO_UTIL", "Fracción de la memoria que se deja usar", 0.80, "%", "0%",
     SRC.format("7.4") + " (ejemplo)", "Supuesto", "El resto queda como holgura"),
    ("OVERHEAD_REPLICA", "Activaciones, motor y espacios de trabajo por réplica", 40, "GB", "0",
     SRC.format("7.4") + " (ejemplo)", "Supuesto",
     "Hay que verificarlo con el motor real; con 2 GPUs podría ser menor"),
    ("PET_EJEMPLO", "Peticiones simultáneas del ejemplo del enunciado", 40, "peticiones", "0",
     SRC.format("7.4"), "Documentado", "Para comparar con el ejemplo"),
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

# ---------------------------------------------------------------- Demanda
# Main table: one column for the base scenario, one for the stress scenario.
# Templates: {c} = this column, {m} = demand multiplier ("" or "*FACTOR_CRECIMIENTO"),
# {KEY} = row number of another result in this table.
# (key, label, template, unit, number_format, note)
DEMAND = [
    ("Peticiones", None),
    ("PET_DIA", "Peticiones por día", "=DAU*REQ_POR_USUARIO{m}", "peticiones/día", "#,##0",
     "Usuarios × peticiones por usuario"),
    ("RPS_PROM", "Tasa promedio", "={c}{PET_DIA}/SEG_DIA", "peticiones/s", "0.00",
     "Peticiones del día repartidas en 86.400 s"),
    ("RPS_PICO", "Tasa en el pico", "={c}{RPS_PROM}*FACTOR_PICO", "peticiones/s", "0.00",
     "Es la cifra que manda en el dimensionamiento"),
    ("Tokens por segundo en el pico", None),
    ("TOKS_ENT_PICO", "Tokens de entrada (prefill)", "={c}{RPS_PICO}*TOK_ENTRADA", "tokens/s", "#,##0",
     "Trabajo antes de que empiece la respuesta"),
    ("TOKS_SAL_PICO", "Tokens de salida (decode)", "={c}{RPS_PICO}*TOK_SALIDA", "tokens/s", "#,##0",
     "Se generan uno a uno; aquí pesan la memoria y la concurrencia"),
    ("RATIO_ENT_SAL", "Relación entrada / salida", "=TOK_ENTRADA/TOK_SALIDA", "x", "0.0",
     "La misma suma de tokens con otra proporción es otra carga"),
    ("TOK_POR_PET", "Tokens acumulados por petición al terminar", "=TOK_ENTRADA+TOK_SALIDA", "tokens",
     "#,##0", "Lo que llega a ocupar en KV cache una petición promedio (paso 2)"),
    ("Tokens por día", None),
    ("TOK_ENT_DIA", "Tokens de entrada por día", "={c}{PET_DIA}*TOK_ENTRADA", "tokens/día", "#,##0",
     "Sirve luego para costo por millón de tokens"),
    ("TOK_SAL_DIA", "Tokens de salida por día", "={c}{PET_DIA}*TOK_SALIDA", "tokens/día", "#,##0", ""),
    ("Concurrencia (ley de Little)", None),
    ("CONC_PROM", "Peticiones en el sistema, hora promedio", "={c}{RPS_PROM}*TIEMPO_SISTEMA",
     "peticiones", "0", "Tasa × tiempo promedio en el sistema"),
    ("CONC_PICO", "Peticiones en el sistema, hora pico", "={c}{RPS_PICO}*TIEMPO_SISTEMA",
     "peticiones", "0", "Suma las que esperan en cola y las que se están procesando"),
    ("TOK_VIVOS_PICO", "Tokens vivos en el pico (cota superior)", "={c}{CONC_PICO}*{c}{TOK_POR_PET}",
     "tokens", "#,##0",
     "Supone que todas las peticiones están completas; las que van a mitad ocupan menos"),
]

INPUT_LENGTHS = [1_000, 2_000, 4_000, 6_000, 7_692, 8_000]
OUTPUT_LENGTHS = [250, 500, 750, 1_000]
SYSTEM_TIMES = [10, 15, 20, 22, 30]

# ---------------------------------------------------------------- Registro
# (decision, value_unit, source, uncertainty, classification)
# value_unit may be a formula (starts with "=").
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
    ("Tasa en el pico", '=ROUND(RPS_PICO,2)&" peticiones/s"',
     "Demanda: peticiones/día ÷ 86.400 × factor de pico", "La del factor de pico (6, dado)", "Calculado"),
    ("Tokens por segundo en el pico",
     '=ROUND(TOKS_ENT_PICO,0)&" entrada / "&ROUND(TOKS_SAL_PICO,0)&" salida (tokens/s)"',
     "Demanda: tasa pico × tokens promedio", "Depende de que los promedios se mantengan", "Calculado"),
    ("Tiempo promedio en el sistema", '=TIEMPO_SISTEMA&" s"',
     "Ejemplo del enunciado §6.3", "Alta: falta evidencia de rendimiento (paso 3)", "Supuesto"),
    ("Concurrencia en el pico", '=ROUND(CONC_PICO,0)&" peticiones"',
     "Ley de Little: tasa pico × tiempo promedio", "Hereda la del tiempo promedio", "Calculado"),
    ("Duplicar la entrada promedio", "8.000 + 500 = 8.500 tokens > 8.192",
     "Demanda: tabla de sensibilidad de entrada",
     "Ninguna: el promedio ya no cabe en la ventana; el escenario exige cambiar la ventana o recortar",
     "Calculado"),
    ("Alternativas de modelo comparadas", "Llama 3.1 70B Instruct en BF16 y en FP8",
     "Decisión de Juliana (13-nov-2026); §7.1 pide al menos dos", "Baja", "Documentado"),
    ("GPU para el cálculo", "H100 SXM 80 GB",
     "Decisión de Juliana: la del ejemplo del enunciado; dato de la página de NVIDIA", "Baja", "Documentado"),
    ("Arquitectura del modelo", "80 capas, 64 cabezas, 8 cabezas KV, hidden 8.192",
     "config.json (espejo NousResearch)", "Cotejar con el artefacto exacto de Meta", "Documentado"),
    ("Parámetros totales", '=ROUND(PARAM_TOTAL/1E9,2)&" miles de millones"',
     "Calculado desde config.json (capas + embeddings)", "Baja: coincide con los 71B que muestra HF",
     "Calculado"),
    ("KV cache por token (16 bits)", '=ROUND(KV_TOKEN_BF16/1E6,3)&" MB/token"',
     "2 × capas × cabezas KV × dimensión × 2 bytes", "Baja para esta arquitectura", "Calculado"),
    ("Memoria de pesos", '=ROUND(MEM_PESOS_BF16,1)&" GB (BF16) / "&ROUND(MEM_PESOS_FP8_4,1)&" GB (FP8)"',
     "Parámetros × bytes; en FP8 embeddings y lm_head siguen en 16 bits", "Faltan escalas y metadatos",
     "Calculado"),
    ("Calidad de FP8 en español", "Sin medir",
     "La ficha de RedHat pone fuera de alcance los idiomas distintos al inglés",
     "Alta: puede no llegar al 90 % de la referencia", "Supuesto"),
    ("Holgura de memoria", "80 % utilizable", "Ejemplo del enunciado §7.4", "Media", "Supuesto"),
    ("Activaciones y motor", "40 GB por réplica", "Ejemplo del enunciado §7.4",
     "Alta: depende del motor y del tamaño de lote", "Supuesto"),
    ("Fuente del rendimiento", "Dato sintético del enunciado + pruebas de orden de magnitud",
     "Decisión de Juliana (opción A); MLPerf no tiene un caso comparable", "Alta", "Documentado"),
    ("Modelo del diseño", "Llama 3.1 70B BF16; FP8 queda como variante a futuro",
     "Decisión de Juliana: no hay datos de velocidad ni de calidad en español para FP8", "Baja", "Documentado"),
    ("Capacidad por réplica", '=CAP_REPLICA&" peticiones/s con 4×H100"',
     "Enunciado §8.4, explícitamente sintético", "Alta: falta medirla", "Supuesto"),
    ("Cómputo usado en el punto de operación", '=ROUND(USO_COMPUTO*100,0)&" % del máximo denso"',
     "Rendimiento: tokens/s × 2 × parámetros ÷ cómputo de 4 H100", "Media: depende del valor denso supuesto",
     "Calculado"),
    ("Tiempo mínimo por paso de generación", '=ROUND(T_PASO_MIN,1)&" ms (límite: 40 ms)"',
     "Rendimiento: (pesos + KV) ÷ ancho de banda", "Media: es un mínimo teórico", "Calculado"),
    ("Cómputo BF16 de la H100", "1.979 TFLOPS con dispersión; ~990 densos",
     "NVIDIA, página H100 (valor con sparsity) ÷ 2", "Media: NVIDIA no publica el denso en esa página",
     "Supuesto"),
    ("Factor de utilización de diseño", "70 %", "Ejemplo del enunciado §9", "Media", "Supuesto"),
    ("Inventario de cómputo",
     '=NODOS&" nodos, "&GPUS_TOTAL&" GPUs, "&REPL_TOTAL&" réplicas, "&RACKS_COMPUTO&" racks"',
     "Inventario: pico ÷ (capacidad × 70 % × 2 réplicas por nodo) + 1 nodo",
     "Hereda la del dato sintético de capacidad", "Calculado"),
    ("Dominio de falla cubierto", "Un nodo; no un rack completo",
     "Inventario: tabla de fallas", "Baja", "Calculado"),
    ("Nodos por rack", "3", "Ejemplo del enunciado §12.2", "Se revisa con energía en el paso 7", "Supuesto"),
]


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = WRAP
        cell.border = BORDER


def style_section(ws, row, ncols):
    for c in range(1, ncols + 1):
        ws.cell(row=row, column=c).fill = SECTION_FILL
        ws.cell(row=row, column=c).font = SECTION_FONT


def style_body(ws, row, ncols):
    for c in range(1, ncols + 1):
        ws.cell(row=row, column=c).border = BORDER
        ws.cell(row=row, column=c).alignment = WRAP


def add_name(wb, name, sheet, col, row):
    wb.defined_names[name] = DefinedName(name, attr_text=f"{sheet}!${col}${row}")


def fit_page(ws):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def title(ws, text, note):
    ws["A1"] = text
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = note
    ws["A2"].font = NOTE_FONT


def build_inputs(wb):
    ws = wb.active
    ws.title = "Entradas"
    title(ws, "Entradas del modelo",
          "Texto azul = valor escrito a mano. Todo lo demás del libro se calcula a partir de esta hoja.")

    headers = ["Nombre", "Parámetro", "Valor", "Unidad", "Fuente", "Clasificación", "Nota"]
    ws.append([])
    ws.append(headers)
    style_header(ws, ws.max_row, len(headers))

    for item in INPUTS:
        if item[1] is None:
            ws.append([item[0]])
            style_section(ws, ws.max_row, len(headers))
            continue
        name, label, value, unit, fmt, source, klass, note = item
        ws.append([name, label, value, unit, source, klass, note])
        r = ws.max_row
        ws.cell(row=r, column=3).font = INPUT_FONT
        ws.cell(row=r, column=3).number_format = fmt
        style_body(ws, r, len(headers))
        add_name(wb, name, "Entradas", "C", r)

    for col, width in zip("ABCDEFG", [22, 46, 16, 22, 26, 15, 48]):
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A5"


def build_demand(wb):
    ws = wb.create_sheet("Demanda")
    title(ws, "Demanda: de usuarios a carga",
          "Todo sale de la hoja Entradas. La columna de estrés multiplica la demanda por FACTOR_CRECIMIENTO.")

    headers = ["Nombre", "Concepto", "Base", "Estrés (demanda ×2)", "Unidad", "Cómo se obtiene"]
    ncols = len(headers)
    ws.append([])
    ws.append(headers)
    style_header(ws, ws.max_row, ncols)

    # First pass: assign row numbers so templates can refer to each other.
    rows, r = {}, ws.max_row
    for item in DEMAND:
        r += 1
        if item[1] is not None:
            rows[item[0]] = r

    for item in DEMAND:
        if item[1] is None:
            ws.append([item[0]])
            style_section(ws, ws.max_row, ncols)
            continue
        key, label, tpl, unit, fmt, note = item
        base = tpl.format(c="C", m="", **rows)
        stress = tpl.format(c="D", m="*FACTOR_CRECIMIENTO", **rows)
        ws.append([key, label, base, stress, unit, note])
        r = ws.max_row
        for col in (3, 4):
            ws.cell(row=r, column=col).number_format = fmt
        style_body(ws, r, ncols)
        add_name(wb, key, "Demanda", "C", r)
        add_name(wb, key + "_ESTRES", "Demanda", "D", r)

    # ---- Sensitivity: input length
    ws.append([])
    ws.append(["Sensibilidad 1: longitud promedio de entrada (salida fija en TOK_SALIDA)"])
    style_section(ws, ws.max_row, ncols)
    ws.append(["", "Tokens de entrada", "Entrada en el pico (tokens/s)", "Tokens por petición",
               "¿Cabe en la ventana?", "Frente a la base"])
    style_header(ws, ws.max_row, ncols)
    for n in INPUT_LENGTHS:
        r = ws.max_row + 1
        ws.cell(row=r, column=2, value=n).font = INPUT_FONT
        ws.cell(row=r, column=3, value=f"=RPS_PICO*B{r}")
        ws.cell(row=r, column=4, value=f"=B{r}+TOK_SALIDA")
        ws.cell(row=r, column=5, value=f'=IF(D{r}<=CONTEXTO_MAX,"Sí","No")')
        ws.cell(row=r, column=6, value=f"=B{r}/TOK_ENTRADA")
        for col, fmt in zip((2, 3, 4, 6), ("#,##0", "#,##0", "#,##0", "0.00x")):
            ws.cell(row=r, column=col).number_format = fmt
        style_body(ws, r, ncols)
    ws.append(["", "7.692 es la entrada más larga que deja espacio a 500 tokens de salida. "
               "8.000 es el escenario de entrada duplicada del §14.4."])
    ws.cell(row=ws.max_row, column=2).font = NOTE_FONT

    # ---- Sensitivity: output length
    ws.append([])
    ws.append(["Sensibilidad 2: longitud promedio de salida (entrada fija en TOK_ENTRADA)"])
    style_section(ws, ws.max_row, ncols)
    ws.append(["", "Tokens de salida", "Salida en el pico (tokens/s)", "Tokens por petición",
               "¿Cabe en la ventana?", "¿Respeta la salida máxima?"])
    style_header(ws, ws.max_row, ncols)
    for n in OUTPUT_LENGTHS:
        r = ws.max_row + 1
        ws.cell(row=r, column=2, value=n).font = INPUT_FONT
        ws.cell(row=r, column=3, value=f"=RPS_PICO*B{r}")
        ws.cell(row=r, column=4, value=f"=TOK_ENTRADA+B{r}")
        ws.cell(row=r, column=5, value=f'=IF(D{r}<=CONTEXTO_MAX,"Sí","No")')
        ws.cell(row=r, column=6, value=f'=IF(B{r}<=SALIDA_MAX,"Sí","No")')
        for col in (2, 3, 4):
            ws.cell(row=r, column=col).number_format = "#,##0"
        style_body(ws, r, ncols)
    ws.append(["", "Una salida más larga también alarga el tiempo en el sistema y con eso la "
               "concurrencia; ese efecto se ve en la tabla 3."])
    ws.cell(row=ws.max_row, column=2).font = NOTE_FONT

    # ---- Sensitivity: time in system
    ws.append([])
    ws.append(["Sensibilidad 3: tiempo promedio en el sistema (ley de Little)"])
    style_section(ws, ws.max_row, ncols)
    ws.append(["", "Tiempo promedio (s)", "Concurrencia pico, base", "Concurrencia pico, estrés",
               "Tokens vivos, base (cota)", ""])
    style_header(ws, ws.max_row, ncols)
    for t in SYSTEM_TIMES:
        r = ws.max_row + 1
        ws.cell(row=r, column=2, value=t).font = INPUT_FONT
        ws.cell(row=r, column=3, value=f"=RPS_PICO*B{r}")
        ws.cell(row=r, column=4, value=f"=RPS_PICO_ESTRES*B{r}")
        ws.cell(row=r, column=5, value=f"=C{r}*TOK_POR_PET")
        for col, fmt in zip((2, 3, 4, 5), ("0", "#,##0", "#,##0", "#,##0")):
            ws.cell(row=r, column=col).number_format = fmt
        style_body(ws, r, ncols)
    ws.append(["", "22 s = 2 s de TTFT + 500 tokens × 40 ms. Son límites p95, no promedios, así que "
               "solo sirve como referencia de un caso lento, no como aplicación exacta de la ley."])
    ws.cell(row=ws.max_row, column=2).font = NOTE_FONT

    # ---- What Little's law does not tell
    ws.append([])
    ws.append(["Qué no dice esta hoja"])
    style_section(ws, ws.max_row, ncols)
    for text in [
        "La concurrencia junta las peticiones en cola y las que están en proceso. "
        "Cuántas hay de cada una depende de la capacidad (pasos 3 y 4).",
        "Usuarios conectados y sesiones abiertas son otra cosa: alguien puede tener el chat abierto "
        "sin una petición en curso.",
        "Los promedios esconden la variabilidad. La mezcla de peticiones cortas y largas se evalúa "
        "en el paso 3.",
    ]:
        ws.append(["", text])
        ws.merge_cells(start_row=ws.max_row, start_column=2, end_row=ws.max_row, end_column=6)
        ws.cell(row=ws.max_row, column=2).alignment = WRAP
        ws.row_dimensions[ws.max_row].height = 45

    for col, width in zip("ABCDEF", [18, 44, 22, 22, 22, 58]):
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A5"


ARCH = [
    ("HEAD_DIM", "Dimensión por cabeza", "=HIDDEN/ATT_HEADS", "-", "0",
     "hidden_size / cabezas de atención (el config no trae head_dim)"),
    ("PARAM_CAPA", "Parámetros por capa",
     "=2*HIDDEN*HIDDEN+2*HIDDEN*KV_HEADS*HEAD_DIM+3*HIDDEN*INTERMEDIA+2*HIDDEN", "parámetros", "#,##0",
     "Atención (Q, O, K, V) + MLP (3 matrices) + 2 normalizaciones; Llama no usa sesgos"),
    ("PARAM_CAPAS", "Parámetros de todas las capas", "=CAPAS*PARAM_CAPA", "parámetros", "#,##0",
     "Son las capas lineales que FP8 cuantiza"),
    ("PARAM_BORDE", "Embeddings, lm_head y norma final", "=2*VOCAB*HIDDEN+HIDDEN", "parámetros", "#,##0",
     "La versión FP8 los deja en 16 bits"),
    ("PARAM_TOTAL", "Parámetros totales", "=PARAM_CAPAS+PARAM_BORDE", "parámetros", "#,##0",
     "Debe dar cerca de 70.600 millones (Hugging Face muestra \"71B\")"),
]

# Alternatives: (column header, gpus, bytes/param in layers, bytes/param in borders, KV bytes/element)
ALTS = [
    ("BF16 · 4×H100\n(diseño)", "=GPUS_REPLICA", 2, 2, 2),
    ("FP8 · 4×H100", 4, 1, 2, 2),
    ("FP8 · 2×H100", 2, 1, 2, 2),
]

# (key, label, template, unit, number_format, note). {c} = column, {KEY} = row of another item.
# Keys starting with "in_" are per-alternative inputs.
MEMORY = [
    ("Configuración de cada alternativa (valores azules = elección)", None),
    ("in_gpus", "GPUs por réplica", None, "GPUs", "0", "La réplica queda dentro de un solo nodo de 8 GPUs"),
    ("in_bpc", "Bytes por parámetro en las capas", None, "bytes", "0", "BF16 = 2, FP8 = 1"),
    ("in_bpb", "Bytes por parámetro en embeddings y lm_head", None, "bytes", "0",
     "La ficha FP8 dice que lm_head no se cuantiza"),
    ("in_bkv", "Bytes por elemento del KV cache", None, "bytes", "0",
     "La ficha FP8 no cuantiza el KV cache: se deja en 16 bits"),
    ("Pesos", None),
    ("MEM_PESOS", "Memoria de pesos", "=(PARAM_CAPAS*{c}{in_bpc}+PARAM_BORDE*{c}{in_bpb})/BYTES_GB", "GB", "0.0",
     "Sin contar escalas ni metadatos, que son pequeños"),
    ("KV cache", None),
    ("KV_TOKEN", "KV cache por token", "=2*CAPAS*KV_HEADS*HEAD_DIM*{c}{in_bkv}", "bytes", "#,##0",
     "2 (K y V) × capas × cabezas KV × dimensión × bytes"),
    ("KV_PET_PROM", "KV por petición promedio", "={c}{KV_TOKEN}*TOK_POR_PET/BYTES_GB", "GB", "0.00",
     "Con los tokens acumulados de la hoja Demanda (4.500)"),
    ("KV_PET_MAX", "KV por petición en el peor caso permitido", "={c}{KV_TOKEN}*CONTEXTO_MAX/BYTES_GB", "GB",
     "0.00", "Ventana completa de 8.192 tokens"),
    ("Presupuesto por réplica", None),
    ("MEM_NOMINAL", "Memoria nominal", "={c}{in_gpus}*GPU_MEM", "GB", "0", "GPUs × memoria por GPU"),
    ("MEM_UTIL", "Memoria utilizable", "={c}{MEM_NOMINAL}*USO_UTIL", "GB", "0.0", "Deja la holgura definida en Entradas"),
    ("MEM_OVERHEAD", "Activaciones, motor y espacios de trabajo", "=OVERHEAD_REPLICA", "GB", "0.0", "Supuesto"),
    ("KV_DISP", "Memoria que queda para KV cache", "={c}{MEM_UTIL}-{c}{MEM_PESOS}-{c}{MEM_OVERHEAD}", "GB",
     "0.0", "Si es negativa, el modelo no cabe con esta holgura"),
    ("TOK_KV_MAX", "Tokens que caben en el KV cache", "=MAX(0,{c}{KV_DISP})*BYTES_GB/{c}{KV_TOKEN}", "tokens",
     "#,##0", "Es el límite que debe usar la política de admisión"),
    ("PET_MAX_PROM", "Peticiones promedio simultáneas que caben",
     "=INT(MAX(0,{c}{KV_DISP})/{c}{KV_PET_PROM})", "peticiones", "0", "Todas de 4.500 tokens"),
    ("PET_MAX_PEOR", "Peticiones simultáneas que caben en el peor caso",
     "=INT(MAX(0,{c}{KV_DISP})/{c}{KV_PET_MAX})", "peticiones", "0", "Todas usando la ventana completa"),
    ("Por GPU", None),
    ("PESOS_GPU", "Pesos por GPU", "={c}{MEM_PESOS}/{c}{in_gpus}", "GB", "0.0",
     "Reparto ideal; en la práctica hay desbalances y partes replicadas"),
    ("OCUP_GPU", "Ocupación por GPU con el KV lleno", "=({c}{MEM_PESOS}+{c}{MEM_OVERHEAD}+MAX(0,{c}{KV_DISP}))/{c}{in_gpus}",
     "GB", "0.0", "No debe pasar de memoria por GPU × fracción utilizable"),
    ("Ejemplo del enunciado §7.4 (PET_EJEMPLO peticiones)", None),
    ("EJ_TOTAL_PROM", "Memoria total con peticiones de 4.500 tokens",
     "={c}{MEM_PESOS}+{c}{MEM_OVERHEAD}+PET_EJEMPLO*{c}{KV_PET_PROM}", "GB", "0.0", "El enunciado da ~239 GB en BF16"),
    ("EJ_CABE_PROM", "¿Cabe en la memoria utilizable?", '=IF({c}{EJ_TOTAL_PROM}<={c}{MEM_UTIL},"Sí","No")', "-", "@", ""),
    ("EJ_TOTAL_PEOR", "Memoria total con peticiones de 8.192 tokens",
     "={c}{MEM_PESOS}+{c}{MEM_OVERHEAD}+PET_EJEMPLO*{c}{KV_PET_MAX}", "GB", "0.0", "El enunciado da ~287 GB en BF16"),
    ("EJ_CABE_PEOR", "¿Cabe en la memoria utilizable?", '=IF({c}{EJ_TOTAL_PEOR}<={c}{MEM_UTIL},"Sí","No")', "-", "@", ""),
    ("Relación con la demanda (solo memoria)", None),
    ("KV_PICO_TOTAL", "KV cache de todas las peticiones del pico (cota)",
     "=TOK_VIVOS_PICO*{c}{KV_TOKEN}/BYTES_GB", "GB", "#,##0", "Tokens vivos de la hoja Demanda × KV por token"),
    ("REPL_MIN_MEM", "Réplicas mínimas solo para que quepa el KV del pico",
     "=IF({c}{KV_DISP}>0,ROUNDUP({c}{KV_PICO_TOTAL}/{c}{KV_DISP},0),\"No cabe\")", "réplicas", "0",
     "Caber no es lo mismo que atender a tiempo: el número final sale de los pasos 3 y 4"),
    ("GPUS_MIN_MEM", "GPUs correspondientes", "=IF(ISNUMBER({c}{REPL_MIN_MEM}),{c}{REPL_MIN_MEM}*{c}{in_gpus},\"-\")",
     "GPUs", "0", ""),
]

COMPARISON = [
    ("Identificador", "meta-llama/Llama-3.1-70B-Instruct", "RedHatAI/Meta-Llama-3.1-70B-Instruct-FP8"),
    ("Versión / revisión", "Lanzado el 23-jul-2024; revisión exacta por congelar",
     "Revisión exacta por congelar"),
    ("Parámetros", "~70.600 millones, densos (no es MoE)", "Los mismos; cambia cómo se guardan"),
    ("Licencia", "Llama 3.1 Community License", "Llama 3.1 (hereda la del original)"),
    ("Español", "Soportado oficialmente. MMLU multilingüe en español: 80,05",
     "Sin evaluar: la ficha pone otros idiomas distintos al inglés fuera de alcance"),
    ("Calidad publicada", "MGSM 86,9; IFEval 87,5 (ficha de Meta)",
     "Recupera 99,88 % del promedio OpenLLM v1, solo en inglés"),
    ("Precisión de pesos / KV", "BF16 / BF16", "FP8 por tensor en capas lineales; KV en 16 bits"),
    ("Contexto soportado", "128k tokens (el servicio usa 8.192)", "Igual que el original"),
    ("Motor de inferencia", "vLLM (versión por fijar)", "vLLM con formato compressed-tensors; requiere FP8 en hardware (H100 lo tiene)"),
    ("Hardware mínimo de referencia", "4×H100 80 GB por réplica", "2×H100 80 GB por réplica"),
    ("Riesgos", "Más GPUs por réplica, más costo y energía",
     "Pérdida de calidad en español no medida; dependencia de un tercero (RedHat) para la cuantización"),
]


def build_memory(wb):
    ws = wb.create_sheet("Modelo y memoria")
    title(ws, "Modelo y memoria por réplica",
          "Datos del modelo y de la GPU en Entradas. Una réplica es una copia completa del modelo que atiende peticiones.")
    ncols = 2 + len(ALTS) + 2  # name, concept, alts..., unit, note
    ucol, ncol = 3 + len(ALTS), 4 + len(ALTS)
    alt_cols = [chr(ord("C") + i) for i in range(len(ALTS))]

    # ---- Architecture block (single value column)
    ws.append([])
    ws.append(["Nombre", "Arquitectura (común a las alternativas)", "Valor", "", "", "Unidad", "Cómo se obtiene"])
    style_header(ws, ws.max_row, ncols)
    for key, label, formula, unit, fmt, note in ARCH:
        r = ws.max_row + 1
        ws.cell(row=r, column=1, value=key)
        ws.cell(row=r, column=2, value=label)
        ws.cell(row=r, column=3, value=formula).number_format = fmt
        ws.cell(row=r, column=ucol, value=unit)
        ws.cell(row=r, column=ncol, value=note)
        style_body(ws, r, ncols)
        add_name(wb, key, "'Modelo y memoria'", "C", r)

    # ---- Per-alternative block
    ws.append([])
    ws.append(["Nombre", "Concepto"] + [a[0] for a in ALTS] + ["Unidad", "Cómo se obtiene"])
    style_header(ws, ws.max_row, ncols)
    ws.row_dimensions[ws.max_row].height = 45

    rows, r = {}, ws.max_row
    for item in MEMORY:
        r += 1
        if item[1] is not None:
            rows[item[0]] = r
    inputs_idx = {"in_gpus": 1, "in_bpc": 2, "in_bpb": 3, "in_bkv": 4}

    for item in MEMORY:
        if item[1] is None:
            ws.append([item[0]])
            style_section(ws, ws.max_row, ncols)
            continue
        key, label, tpl, unit, fmt, note = item
        r = ws.max_row + 1
        ws.cell(row=r, column=1, value="" if key.startswith("in_") else key)
        ws.cell(row=r, column=2, value=label)
        for i, (alt, col) in enumerate(zip(ALTS, alt_cols)):
            cell = ws.cell(row=r, column=3 + i)
            if key in inputs_idx:
                cell.value = alt[inputs_idx[key]]
                if not str(cell.value).startswith("="):
                    cell.font = INPUT_FONT
            else:
                cell.value = tpl.format(c=col, **rows)
            cell.number_format = fmt
        ws.cell(row=r, column=ucol, value=unit)
        ws.cell(row=r, column=ncol, value=note)
        style_body(ws, r, ncols)
        if not key.startswith("in_"):
            # Name the reference (first) alternative; others are referenced by cell.
            add_name(wb, key + "_BF16", "'Modelo y memoria'", "C", r)
            add_name(wb, key + "_FP8_4", "'Modelo y memoria'", "D", r)
            add_name(wb, key + "_FP8_2", "'Modelo y memoria'", "E", r)

    # ---- Qualitative comparison (§7.1)
    ws.append([])
    ws.append(["Comparación de alternativas (§7.1)"])
    style_section(ws, ws.max_row, ncols)
    ws.append(["", "Criterio", "Llama 3.1 70B Instruct BF16", "", "Llama 3.1 70B Instruct FP8", "", ""])
    r = ws.max_row
    ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
    ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=7)
    style_header(ws, r, ncols)
    for crit, a, b in COMPARISON:
        ws.append(["", crit, a, "", b, "", ""])
        r = ws.max_row
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
        ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=7)
        style_body(ws, r, ncols)
        ws.row_dimensions[r].height = 30

    ws.append([])
    ws.append(["", "Fuentes: config.json (espejo NousResearch en Hugging Face), ficha del modelo Llama 3.1 de Meta, "
               "ficha de RedHatAI/Meta-Llama-3.1-70B-Instruct-FP8 y página de producto H100 de NVIDIA. "
               "Consultadas el 13-nov-2026."])
    ws.cell(row=ws.max_row, column=2).font = NOTE_FONT

    for col, width in zip("ABCDEFG", [16, 46, 20, 20, 20, 12, 58]):
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "C5"


# (key, label, formula, unit, number_format, note); single value column.
PERFORMANCE = [
    ("Punto de operación por réplica (dato sintético §8.4)", None),
    ("CAP_REPLICA_R", "Peticiones por segundo cumpliendo los SLOs", "=CAP_REPLICA", "peticiones/s", "0.00",
     "Supuesto: el enunciado lo da solo para practicar"),
    ("TOKS_ENT_REPLICA", "Tokens de entrada por segundo", "=CAP_REPLICA*TOK_ENTRADA", "tokens/s", "#,##0",
     "Es el mismo punto de operación, no un máximo aparte de prefill"),
    ("TOKS_SAL_REPLICA", "Tokens de salida por segundo", "=CAP_REPLICA*TOK_SALIDA", "tokens/s", "#,##0",
     "Tampoco es un máximo aparte de decode: los dos ocurren a la vez"),
    ("CAP_GPU", "Peticiones por segundo por GPU", "=CAP_REPLICA/GPUS_REPLICA", "peticiones/s", "0.00", ""),
    ("Prueba de orden de magnitud: cómputo", None),
    ("FLOP_TOKEN", "Operaciones por token (aprox.)", "=2*PARAM_TOTAL", "FLOP/token", "#,##0",
     "Regla de 2 operaciones por parámetro; deja fuera la atención, que con 4.500 tokens suma poco"),
    ("FLOPS_USADOS", "Cómputo que exige el punto de operación",
     "=(TOKS_ENT_REPLICA+TOKS_SAL_REPLICA)*FLOP_TOKEN/1E12", "TFLOPS", "#,##0", "Tokens/s × operaciones por token"),
    ("FLOPS_PICO", "Cómputo máximo de la réplica (denso)",
     "=GPUS_REPLICA*H100_BF16_SPARSE/FACTOR_DISPERSION", "TFLOPS", "#,##0", "GPUs × valor de NVIDIA ÷ 2"),
    ("USO_COMPUTO", "Fracción del cómputo que se usaría", "=FLOPS_USADOS/FLOPS_PICO", "%", "0%",
     "Por encima de 100 % sería imposible; muy cerca de 100 % sería sospechoso"),
    ("Prueba de orden de magnitud: memoria en la generación", None),
    ("SEC_DECODE", "Peticiones generando a la vez en la réplica",
     "=TOKS_SAL_REPLICA*SLO_TPOT/1000", "peticiones", "0",
     "Ley de Little con 40 ms por token: tokens/s × tiempo por token"),
    ("CABE_DECODE", "¿Caben en el KV cache de la réplica?",
     '=IF(SEC_DECODE<=PET_MAX_PROM_BF16,"Sí","No")', "-", "@", "Compara con la hoja Modelo y memoria"),
    ("KV_LEIDO", "KV cache que se lee en cada paso de generación",
     "=SEC_DECODE*(TOK_ENTRADA+TOK_SALIDA/2)*KV_TOKEN_BF16/BYTES_GB", "GB", "0.0",
     "Cada petición va en promedio a mitad de su respuesta"),
    ("T_PASO_MIN", "Tiempo mínimo de un paso de generación",
     "=(MEM_PESOS_BF16+KV_LEIDO)/(GPUS_REPLICA*H100_BW*1000)*1000", "ms", "0.0",
     "Pesos + KV leídos al ancho de banda máximo; el real es mayor"),
    ("CUMPLE_TPOT", "¿Queda por debajo de 40 ms por token?", '=IF(T_PASO_MIN<SLO_TPOT,"Sí","No")', "-", "@",
     "Si no cumpliera ni en teoría, el dato sintético sería imposible"),
    ("Sensibilidad (§14.4)", None),
    ("CAP_REPLICA_BAJA", "Peticiones por segundo si la réplica rinde menos", "=CAP_REPLICA*(1-CAIDA_SENS)",
     "peticiones/s", "0.00", "Su efecto en réplicas y costos se ve en los pasos 4 y 9"),
]

VALIDATION = [
    "Qué es este dato: un supuesto del enunciado (§8.4), no una medición ni un benchmark de NVIDIA, vLLM o Meta. "
    "No se puede escribir \"medimos\" ni \"demostramos\".",
    "Qué muestran las pruebas de arriba: que 2 peticiones/s no rompe la física de la H100. No prueban que se alcance.",
    "Cómo se validaría: fijar modelo y revisión, vLLM y su versión, 4×H100 con paralelismo tensorial 4 y el perfil "
    "4.000/500; calentar; subir la tasa de llegada en lazo abierto hasta que se rompa algún SLO; repetir 3 veces.",
    "Qué registrar en cada prueba: TTFT, tiempo por token y tiempo total en p95, peticiones completadas, rechazadas y "
    "fallidas, estado del caché de prefijos y consumo eléctrico.",
    "Después: repetir con la mezcla de peticiones cortas y largas, con arranque en frío y con la pérdida de un nodo.",
    "Variante FP8: no hay datos de su velocidad ni de su calidad en español. Queda como ahorro posible a futuro, "
    "condicionado a medir ambas cosas.",
    "Referencia externa: MLPerf Inference v5.0 tiene una prueba interactiva de Llama 2 70B con TPOT p99 de 40 ms y "
    "TTFT p99 de 450 ms. Muestra que el límite de 40 ms es una meta realista en la industria, pero usa otro modelo "
    "y textos más cortos, así que sus cifras no se trasladan a este caso.",
]


def build_performance(wb):
    ws = wb.create_sheet("Rendimiento")
    title(ws, "Rendimiento por réplica",
          "Diseño con Llama 3.1 70B BF16 en 4×H100. El dato de capacidad es sintético; las pruebas solo revisan "
          "que sea físicamente posible.")
    headers = ["Nombre", "Concepto", "Valor", "Unidad", "Cómo se obtiene"]
    ncols = len(headers)
    ws.append([])
    ws.append(headers)
    style_header(ws, ws.max_row, ncols)

    for item in PERFORMANCE:
        if item[1] is None:
            ws.append([item[0]])
            style_section(ws, ws.max_row, ncols)
            continue
        key, label, formula, unit, fmt, note = item
        ws.append([key, label, formula, unit, note])
        r = ws.max_row
        ws.cell(row=r, column=3).number_format = fmt
        style_body(ws, r, ncols)
        add_name(wb, key, "Rendimiento", "C", r)

    ws.append([])
    ws.append(["Límites del dato y plan de validación (§8)"])
    style_section(ws, ws.max_row, ncols)
    for text in VALIDATION:
        ws.append(["", text])
        r = ws.max_row
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
        ws.cell(row=r, column=2).alignment = WRAP
        ws.row_dimensions[r].height = 32

    for col, width in zip("ABCDE", [20, 50, 16, 14, 70]):
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A5"


INVENTORY = [
    ("Dimensionamiento", None),
    ("REPL_POR_NODO", "Réplicas por nodo", "=INT(GPUS_NODO/GPUS_REPLICA)", "réplicas", "0",
     "Cada réplica usa GPUs de un solo nodo, así el paralelismo tensorial no cruza la red"),
    ("CAP_UTIL_REPL", "Carga admitida por réplica", "=CAP_REPLICA*UTIL_DISENO", "peticiones/s", "0.00",
     "Capacidad × factor de utilización"),
    ("REPL_MIN", "Réplicas mínimas sin contar fallas", "=ROUNDUP(RPS_PICO/CAP_UTIL_REPL,0)", "réplicas", "0",
     "Pico ÷ carga admitida por réplica, redondeado hacia arriba"),
    ("NODOS_SIN_FALLA", "Nodos para cubrir el pico sin fallas", "=ROUNDUP(RPS_PICO/(CAP_UTIL_REPL*REPL_POR_NODO),0)",
     "nodos", "0", ""),
    ("NODOS", "Nodos instalados", "=NODOS_SIN_FALLA+NODOS_PERDIDOS", "nodos", "0",
     "Se suma el nodo que se puede perder (§4.3)"),
    ("Inventario", None),
    ("GPUS_TOTAL", "GPUs", "=NODOS*GPUS_NODO", "GPUs", "0", ""),
    ("REPL_TOTAL", "Réplicas", "=NODOS*REPL_POR_NODO", "réplicas", "0", ""),
    ("RACKS_COMPUTO", "Racks de cómputo", "=ROUNDUP(NODOS/NODOS_POR_RACK,0)", "racks", "0",
     "Falta el rack de red y almacenamiento (paso 7)"),
    ("Uso de la capacidad instalada", None),
    ("USO_NORMAL", "Pico frente a la capacidad instalada", "=RPS_PICO/(REPL_TOTAL*CAP_REPLICA)", "%", "0.0%",
     "El enunciado da 57,9 %"),
    ("USO_FALLA_NODO", "Pico frente a la capacidad tras perder un nodo",
     "=RPS_PICO/((REPL_TOTAL-REPL_POR_NODO)*CAP_REPLICA)", "%", "0.0%", "El enunciado da 65,1 %"),
    ("Carga en cada réplica durante el pico", None),
    ("CONC_REPL_NORMAL", "Peticiones en curso por réplica, sin fallas",
     "=RPS_PICO/REPL_TOTAL*TIEMPO_SISTEMA", "peticiones", "0", "Ley de Little por réplica"),
    ("CONC_REPL_FALLA", "Peticiones en curso por réplica, tras perder un nodo",
     "=RPS_PICO/(REPL_TOTAL-REPL_POR_NODO)*TIEMPO_SISTEMA", "peticiones", "0",
     "Debe quedar por debajo de lo que cabe en memoria"),
    ("CABE_REPL_FALLA", "¿Cabe en el KV cache de la réplica?", '=IF(CONC_REPL_FALLA<=PET_MAX_PROM_BF16,"Sí","No")',
     "-", "@", "Compara con la hoja Modelo y memoria (50 peticiones promedio)"),
]

# (label, replicas lost formula)
FAILURES = [
    ("Ninguna", "0"),
    ("Una GPU (cae su réplica)", "1"),
    ("Un nodo", "REPL_POR_NODO"),
    ("Un rack de cómputo", "MIN(NODOS_POR_RACK,NODOS)*REPL_POR_NODO"),
]

# (label, demand multiplier, capacity multiplier)
SCENARIOS = [
    ("Base", "1", "1"),
    ("Demanda ×2", "FACTOR_CRECIMIENTO", "1"),
    ("Rendimiento −30 %", "1", "(1-CAIDA_SENS)"),
    ("Demanda ×2 y rendimiento −30 %", "FACTOR_CRECIMIENTO", "(1-CAIDA_SENS)"),
]


def build_inventory(wb):
    ws = wb.create_sheet("Inventario")
    title(ws, "Réplicas, nodos y holgura",
          "Una réplica es una copia del modelo (4 GPUs). Un nodo es un servidor físico (8 GPUs, 2 réplicas).")
    ncols = 7
    ws.append([])
    ws.append(["Nombre", "Concepto", "Valor", "Unidad", "Cómo se obtiene", "", ""])
    style_header(ws, ws.max_row, ncols)

    for item in INVENTORY:
        if item[1] is None:
            ws.append([item[0]])
            style_section(ws, ws.max_row, ncols)
            continue
        key, label, formula, unit, fmt, note = item
        ws.append([key, label, formula, unit, note])
        r = ws.max_row
        ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=7)
        ws.cell(row=r, column=3).number_format = fmt
        style_body(ws, r, ncols)
        add_name(wb, key, "Inventario", "C", r)

    # ---- Failure table
    ws.append([])
    ws.append(["Capacidad que queda ante cada falla (demanda base)"])
    style_section(ws, ws.max_row, ncols)
    ws.append(["", "Falla", "Réplicas que quedan", "Capacidad con holgura (peticiones/s)",
               "¿Cubre el pico con holgura?", "Capacidad al 100 % (peticiones/s)", "¿Cubre el pico al 100 %?"])
    style_header(ws, ws.max_row, ncols)
    ws.row_dimensions[ws.max_row].height = 45
    for label, lost in FAILURES:
        r = ws.max_row + 1
        ws.cell(row=r, column=2, value=label)
        ws.cell(row=r, column=3, value=f"=REPL_TOTAL-{lost}")
        ws.cell(row=r, column=4, value=f"=C{r}*CAP_UTIL_REPL").number_format = "0.0"
        ws.cell(row=r, column=5, value=f'=IF(D{r}>=RPS_PICO,"Sí","No")')
        ws.cell(row=r, column=6, value=f"=C{r}*CAP_REPLICA").number_format = "0.0"
        ws.cell(row=r, column=7, value=f'=IF(F{r}>=RPS_PICO,"Sí","No")')
        style_body(ws, r, ncols)
    for text in [
        "\"Al 100 %\" significa trabajar justo en el punto donde aún se cumplen los SLOs, sin margen. Si la carga "
        "varía un poco, la latencia se dispara.",
        "El diseño promete tolerar un nodo. No promete tolerar un rack: para eso harían falta más nodos "
        "(ver la tabla de escenarios).",
        "Que la capacidad alcance no prueba que se cumpla la latencia durante la falla: faltan la detección, las "
        "peticiones en curso, los reintentos y el calentamiento de réplicas (paso 10).",
    ]:
        ws.append(["", text])
        r = ws.max_row
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
        ws.cell(row=r, column=2).alignment = WRAP
        ws.cell(row=r, column=2).font = NOTE_FONT
        ws.row_dimensions[r].height = 30

    # ---- Scenarios
    ws.append([])
    ws.append(["Escenarios de sensibilidad (§14.4)"])
    style_section(ws, ws.max_row, ncols)
    ws.append(["", "Escenario", "Pico (peticiones/s)", "Capacidad por réplica (peticiones/s)",
               "Nodos (tolera 1 nodo)", "GPUs", "Nodos si se quiere tolerar un rack"])
    style_header(ws, ws.max_row, ncols)
    ws.row_dimensions[ws.max_row].height = 45
    for label, dm, cm in SCENARIOS:
        r = ws.max_row + 1
        ws.cell(row=r, column=2, value=label)
        ws.cell(row=r, column=3, value=f"=RPS_PICO*{dm}").number_format = "0.00"
        ws.cell(row=r, column=4, value=f"=CAP_REPLICA*{cm}").number_format = "0.00"
        base = f"ROUNDUP(C{r}/(D{r}*UTIL_DISENO*REPL_POR_NODO),0)"
        ws.cell(row=r, column=5, value=f"={base}+NODOS_PERDIDOS")
        ws.cell(row=r, column=6, value=f"=E{r}*GPUS_NODO")
        ws.cell(row=r, column=7, value=f"={base}+NODOS_POR_RACK")
        style_body(ws, r, ncols)

    for col, width in zip("ABCDEFG", [20, 46, 16, 20, 20, 18, 22]):
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A5"


def build_log(wb):
    ws = wb.create_sheet("Registro")
    ws["A1"] = "Registro de decisiones y datos"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = ("Clasificación: Documentado (fuente identificable), Medido (experimento reproducible), "
                "Calculado (fórmulas visibles), Supuesto (requiere validación). "
                "Una cifra dada por una IA sigue siendo Supuesto hasta verificarla.")
    ws["A2"].font = NOTE_FONT
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
        style_body(ws, ws.max_row, len(headers))

    for col, width in zip("ABCDEFG", [5, 34, 30, 36, 36, 22, 15]):
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A5"


def main():
    wb = Workbook()
    build_inputs(wb)
    build_demand(wb)
    build_memory(wb)
    build_performance(wb)
    build_inventory(wb)
    build_log(wb)
    for ws in wb.worksheets:
        fit_page(ws)
    wb.calculation = CalcProperties(fullCalcOnLoad=True)
    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT)
    print(f"Saved {OUT.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
