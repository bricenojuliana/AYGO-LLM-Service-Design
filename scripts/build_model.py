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
    ("Sobrecarga y límites", None),
    ("RED_IDA_VUELTA", "Viaje de red de ida y vuelta hasta el usuario", 0.2, "s", "0.00",
     "Estimación propia", "Supuesto", "Depende de la ubicación (paso 7)"),
    ("PREFILL_EST", "Tiempo de procesar la entrada (4.000 tokens) en una réplica cargada", 0.5, "s", "0.00",
     "Estimación propia", "Supuesto", "El mínimo teórico está en la hoja Sobrecarga"),
    ("ESPERA_FILA_MAX", "Espera máxima en la fila antes de rechazar", 1, "s", "0.0",
     "Decisión de Juliana", "Documentado", "Debe caber en el presupuesto del TTFT"),
    ("LIM_CONC_USUARIO", "Peticiones simultáneas por usuario", 3, "peticiones", "0",
     "Decisión de Juliana", "Supuesto", "Regla contra abuso; no cambia el dimensionamiento"),
    ("LIM_MIN_USUARIO", "Peticiones por minuto por usuario", 20, "peticiones/min", "0",
     "Decisión de Juliana", "Supuesto", "Regla contra abuso"),
    ("REINTENTOS_MAX", "Reintentos automáticos por petición", 1, "reintentos", "0",
     "Decisión de Juliana", "Supuesto", "Solo si aún no se envió ninguna palabra al usuario"),
    ("Red y almacenamiento", None),
    ("BYTES_TOKEN", "Bytes de texto por token", 4, "bytes", "0",
     SRC.format("11") + " (aproximación didáctica)", "Supuesto", ""),
    ("BYTES_EVENTO", "Bytes por token enviado al usuario, con su envoltorio", 150, "bytes", "0",
     "Estimación propia", "Supuesto",
     "Cada palabra viaja en un pequeño mensaje con encabezados; pesa más que el texto solo"),
    ("NVLINK_BW", "Ancho de banda NVLink entre GPUs del mismo nodo", 900, "GB/s", "0",
     "Guía de usuario DGX H100/H200 de NVIDIA", "Documentado", ""),
    ("ENLACE_EJEMPLO", "Enlace de red de ejemplo", 25, "Gb/s", "0",
     SRC.format("11"), "Documentado", "Para comparar tiempos de carga"),
    ("ALMAC_LECTURA", "Lectura sostenida del almacenamiento central", 10, "GB/s", "0",
     "Estimación propia", "Supuesto", "Se confirma con el equipo que se cotice (paso 9)"),
    ("CACHE_LOCAL_DGX", "Disco local de datos de cada DGX (8 × 3,84 TB)", 30.72, "TB", "0.00",
     "Guía de usuario DGX H100/H200 de NVIDIA", "Documentado", "En RAID 0: rápido, pero sin protección"),
    ("VERSIONES_MODELO", "Versiones del modelo que se guardan", 3, "versiones", "0",
     "Estimación propia", "Supuesto", "La actual, la anterior (para volver atrás) y una en prueba"),
    ("IMAGENES_GB", "Imágenes de contenedores y artefactos", 200, "GB", "0",
     "Estimación propia", "Supuesto", ""),
    ("DATOS_APP_GB", "Datos de la aplicación (cuentas, cuotas, configuración)", 10, "GB", "0",
     "Estimación propia", "Supuesto", "Sin conversaciones: no se guardan"),
    ("LOG_BYTES_PET", "Registro operativo por petición (sin contenido)", 1_000, "bytes", "#,##0",
     "Estimación propia", "Supuesto", "Tiempos, tokens, errores, identificador anónimo"),
    ("RETENCION_DIAS", "Retención de registros y métricas", 30, "días", "0",
     "Estimación propia", "Supuesto", ""),
    ("REPLICACION", "Copias de cada dato en el almacenamiento central", 3, "copias", "0",
     "Estimación propia", "Supuesto", "Aguanta perder un disco o un servidor de almacenamiento"),
    ("HOLGURA_ALMAC", "Holgura de crecimiento del almacenamiento", 2, "x", "0",
     "Estimación propia", "Supuesto", ""),
    ("ENLACE_RESPALDO", "Enlace hacia la copia fuera del sitio", 1, "Gb/s", "0",
     "Estimación propia", "Supuesto", "Define cuánto tarda restaurar desde fuera"),
    ("Energía (base para la ubicación; se detalla en el paso 8)", None),
    ("P_NODO_MAX", "Potencia máxima de un DGX H100", 10.2, "kW", "0.0",
     "Guía de usuario DGX H100/H200 de NVIDIA", "Documentado",
     "La guía la lista como entrada máxima; ya incluye las GPUs, no se suman aparte"),
    ("P_NODO_PROM", "Potencia promedio de un DGX H100 en operación", 6.8, "kW", "0.0",
     SRC.format("13.4") + " (ejemplo)", "Supuesto", "Hay que reemplazarlo por un perfil medido"),
    ("P_AUX_MAX", "Red, almacenamiento y servicios: potencia máxima", 12, "kW", "0.0",
     SRC.format("13.2") + " (ejemplo)", "Supuesto", ""),
    ("P_AUX_PROM", "Red, almacenamiento y servicios: potencia promedio", 9, "kW", "0.0",
     SRC.format("13.4") + " (ejemplo)", "Supuesto", ""),
    ("PUE", "PUE anual del centro de datos", 1.30, "x", "0.00",
     SRC.format("13.4") + " (ejemplo)", "Supuesto", "Se confirma con el proveedor de colocation"),
    ("P_NO_TI", "Cargas que no son de TI en el punto de diseño (enfriamiento, UPS, luces)", 35, "kW", "0.0",
     SRC.format("13.3") + " (ejemplo)", "Supuesto", "Lo pone el proveedor; hay que pedirle su valor"),
    ("ACOMETIDA", "Potencia eléctrica contratada para nuestra zona", 160, "kW", "0",
     SRC.format("13.3") + " (ejemplo)", "Supuesto", ""),
    ("DGX_BTU", "Calor que declara NVIDIA por DGX H100", 38_557, "BTU/h", "#,##0",
     "Guía de usuario DGX H100/H200 de NVIDIA", "Documentado", "Equivale a más que los 10,2 kW: se revisa"),
    ("MODULO_FRIO", "Capacidad de cada módulo de enfriamiento", 60, "kW térmicos", "0",
     SRC.format("13.5") + " (ejemplo)", "Supuesto", ""),
    ("N_MODULOS", "Módulos de enfriamiento", 3, "módulos", "0",
     SRC.format("13.5") + " (ejemplo)", "Supuesto", "N+1: aguanta uno fuera de servicio"),
    ("DELTA_T", "Subida de temperatura del agua en el circuito", 10, "°C", "0",
     SRC.format("13.5") + " (ejemplo)", "Supuesto", ""),
    ("CP_AGUA", "Calor específico del agua", 4_180, "J/(kg·°C)", "#,##0",
     "Propiedad física", "Documentado", "Con glicol sería distinto"),
    ("UPS_KW", "Carga protegida por la UPS", 120, "kW", "0",
     SRC.format("13.6") + " (ejemplo)", "Supuesto", "TI + enfriamiento crítico"),
    ("UPS_MIN", "Autonomía de la UPS", 10, "min", "0",
     SRC.format("13.6") + " (ejemplo)", "Supuesto", "Tiempo para que arranquen los generadores"),
    ("UPS_EFIC", "Eficiencia de la UPS", 0.90, "%", "0%", SRC.format("13.6"), "Supuesto", ""),
    ("UPS_DOD", "Profundidad de descarga permitida", 0.80, "%", "0%", SRC.format("13.6"), "Supuesto", ""),
    ("UPS_ENVEJ", "Capacidad que conserva la batería al final de su vida", 0.80, "%", "0%",
     SRC.format("13.6"), "Supuesto", ""),
    ("WUE", "Agua consumida por kWh de TI", 0, "L/kWh", "0.0",
     "Supuesto: enfriamiento en circuito cerrado sin torres evaporativas", "Supuesto",
     "Hay que pedírselo al proveedor; la hoja Energía muestra qué pasa con 1 y 2 L/kWh"),
    ("VIDA_EQUIPO", "Vida útil de los servidores", 5, "años", "0",
     "Estimación propia", "Supuesto", "Para estimar residuos electrónicos"),
    ("Ubicación", None),
    ("UBICACION", "Ubicación elegida", "São Paulo", "-", "@",
     "Decisión de Juliana (13-nov-2026)", "Documentado", "Debe coincidir con un encabezado de la hoja Ubicación"),
    ("ULTIMA_MILLA", "Tiempo extra de la conexión del usuario (ida y vuelta)", 30, "ms", "0",
     "Estimación propia", "Supuesto", "Se suma a los pings entre servidores"),
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
    ("Componentes de la arquitectura", "Por función, con un producto de ejemplo entre paréntesis",
     "Decisión de Juliana", "Baja", "Documentado"),
    ("RAG", "No se incluye; queda como posible extensión", "Decisión de Juliana; §4.1 lo deja opcional", "Baja",
     "Documentado"),
    ("Espera máxima en la fila", '=ESPERA_FILA_MAX&" s (caben "&ROUND(ESPERA_DISPONIBLE,2)&" s)"',
     "Sobrecarga: TTFT − red − procesar la entrada", "Media: red y prefill son estimaciones", "Calculado"),
    ("Red y procesamiento de la entrada", '=RED_IDA_VUELTA&" s de red, "&PREFILL_EST&" s de entrada"',
     "Estimación propia", "Alta: se revisa con la ubicación y con una medición", "Supuesto"),
    ("Límites por usuario", '=LIM_CONC_USUARIO&" a la vez, "&LIM_MIN_USUARIO&" por minuto, "&REINTENTOS_MAX&" reintento"',
     "Decisión de Juliana", "Baja: no cambian el dimensionamiento", "Supuesto"),
    ("Decisiones de operación de la sección 7", "Revisadas una por una",
     "Aprobadas por Juliana (respuestas a medias se pierden, caché solo por cliente, réplicas siempre "
     "encendidas, etc.)", "Baja", "Documentado"),
    ("Tráfico de aplicación en el pico", '=ROUND(APP_TOTAL,1)&" Mb/s"',
     "Red y almacenamiento: entrada × 4 bytes + salida × 150 bytes", "Media: el envoltorio es estimado",
     "Calculado"),
    ("Comunicación entre GPUs de una réplica", '=ROUND(TP_TRAFICO,1)&" GB/s ("&ROUND(TP_USO_NVLINK*100,1)&" % de NVLink)"',
     "Aproximación de 2 sincronizaciones por capa", "Media: aproximación de orden de magnitud", "Calculado"),
    ("Carga de todas las réplicas", '=ROUND(CARGA_TODAS,0)&" s desde el almacenamiento central"',
     "18 × pesos ÷ lectura del almacenamiento", "Alta: la lectura de 10 GB/s es supuesta", "Calculado"),
    ("Almacenamiento central", '=ROUND(ALM_BRUTO,1)&" TB brutos"',
     "Modelos + imágenes + datos + registros, × 3 copias × 2 de holgura", "Media", "Calculado"),
    ("Supuestos de red y almacenamiento", "11 valores estimados (lectura 10 GB/s, 3 versiones, 30 días, etc.)",
     "Aprobados por Juliana", "Media", "Supuesto"),
    ("Tipo de instalación", "Colocation", "Decisión de Juliana", "Baja", "Documentado"),
    ("Ubicación", '=UBICACION&" (comparada con Querétaro y Santiago)"', "Decisión de Juliana; hoja Ubicación",
     "Baja", "Documentado"),
    ("Precio de la electricidad", '=PRECIO_KWH&" USD/kWh"', "GlobalPetrolPrices, empresas, mar-2026, con impuestos",
     "Media: precio minorista; un contrato grande podría ser menor", "Documentado"),
    ("Emisiones de la red", '=FACTOR_EMISION&" kg CO₂e/kWh"', "Ember 2025 vía agregadores",
     "Media: fuente secundaria", "Documentado"),
    ("Renta de espacio", '=RENTA_KW_MES&" USD/kW/mes"', "CBRE 2026, punto medio del rango",
     "Media: el rango es para 250–500 kW; el nuestro es más chico", "Documentado"),
    ("Peor tiempo de red", '=UB_RTT&" ms"', "WonderNetwork + 30 ms de última milla", "Media", "Calculado"),
    ("Potencia por rack de cómputo", '=RACK_KW&" kW"', "3 DGX × 10,2 kW", "Hay que confirmarla con el proveedor",
     "Calculado"),
    ("Área de radiadores en órbita", '=ROUND(ORB_AREA,0)&" m²"', "Stefan-Boltzmann ideal (§12.4)",
     "Baja como cota mínima", "Calculado"),
    ("Potencia de TI máxima", '=ROUND(E_P_TI_M,1)&" kW; sitio "&ROUND(E_P_SITIO_M,1)&" kW"',
     "Energía: 9 × 10,2 kW + 12 kW; + 35 kW no TI", "Media: las cargas no TI son del ejemplo", "Calculado"),
    ("Energía anual", '=ROUND(E_ANUAL/1000,0)&" MWh/año, USD "&ROUND(E_COSTO,0)',
     "Energía: TI promedio × PUE × 8.760 h", "Alta: el consumo promedio es supuesto", "Calculado"),
    ("Calor declarado por DGX", '=ROUND(DGX_KW_CALOR,1)&" kW (38.557 BTU/h)"',
     "Guía DGX H100; no coincide con los 10,2 kW", "El enfriamiento N+1 alcanza igual", "Calculado"),
    ("Agua directa (WUE)", '=WUE&" L/kWh"', "Supuesto de circuito cerrado", "Alta: pedir el dato al proveedor",
     "Supuesto"),
    ("Valores eléctricos y térmicos del ejemplo", "35 kW no TI, 160 kW contratados, 3 × 60 kW, UPS 120 kW × 10 min",
     "Enunciado §13", "Media: los define el proveedor de colocation", "Supuesto"),
    ("Precio del DGX H100", "USD 350.000", "Decisión de Juliana; enunciado §14.1, dentro del rango público "
     "USD 300–400 mil", "Media: NVIDIA no publica precio oficial", "Supuesto"),
    ("Ajuste por colocation", "Sin obra eléctrica, UPS, enfriamiento ni edificio; +USD 100.000 de adecuación",
     "Decisión de Juliana: el proveedor lo cobra en la renta", "Media", "Supuesto"),
    ("Impuestos de importación", '=ROUND(IMPUESTO_IMPORT*100,0)&" % sobre el equipo importado"',
     "Decisión de Juliana; fuentes: >50 % antes de ReData, 18–25 % de arancel", "Alta: depende de ReData y de la "
     "reforma de 2027", "Supuesto"),
    ("CAPEX / OPEX / TCO", '="USD "&ROUND(CAPEX/1E6,2)&" M / "&ROUND(OPEX/1E6,2)&" M / "&ROUND(TCO/1E6,2)&" M"',
     "Hoja Presupuesto", "Hereda la de precios y rendimiento", "Calculado"),
    ("Costo por petición", '="USD "&ROUND(COSTO_PET,4)&" (a tiempo y sin error: "&ROUND(COSTO_PET_BUENA,4)&")"',
     "TCO ÷ 328,5 millones de peticiones", "Media", "Calculado"),
    ("Escenario que más pesa en el riesgo", "Rendimiento −30 % y hardware +30 %: los dos rompen el CAPEX",
     "Hoja Escenarios", "Baja", "Calculado"),
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
        key = {"Base": "BASE", "Demanda ×2": "X2", "Rendimiento −30 %": "BAJA",
               "Demanda ×2 y rendimiento −30 %": "X2_BAJA"}[label]
        add_name(wb, f"ESC_NODOS_{key}", "Inventario", "E", r)

    for col, width in zip("ABCDEFG", [20, 46, 16, 20, 20, 18, 22]):
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A5"


OVERLOAD = [
    ("Presupuesto de tiempo para la primera palabra (TTFT)", None),
    ("PREFILL_MIN", "Mínimo teórico para procesar la entrada", "=TOK_ENTRADA*FLOP_TOKEN/(FLOPS_PICO*1E12)", "s",
     "0.00", "Con toda la réplica para una sola petición; en la práctica comparte con otras"),
    ("ESPERA_DISPONIBLE", "Tiempo que queda para esperar en la fila", "=SLO_TTFT-RED_IDA_VUELTA-PREFILL_EST", "s",
     "0.00", "TTFT prometido − red − procesar la entrada"),
    ("CABE_ESPERA", "¿La espera máxima elegida cabe en ese tiempo?", '=IF(ESPERA_FILA_MAX<=ESPERA_DISPONIBLE,"Sí","No")',
     "-", "@", "Si no cabe, la fila dejaría pasar peticiones que igual incumplen"),
    ("MARGEN_ESPERA", "Margen que sobra", "=ESPERA_DISPONIBLE-ESPERA_FILA_MAX", "s", "0.00", ""),
    ("Tamaño de la fila", None),
    ("FILA_TOTAL", "Fila máxima en todo el servicio", "=ROUNDUP(RPS_PICO*ESPERA_FILA_MAX,0)", "peticiones", "0",
     "Peticiones que llegan en el pico durante la espera máxima"),
    ("FILA_REPL", "Fila máxima por réplica", "=ROUNDUP(FILA_TOTAL/REPL_TOTAL,0)", "peticiones", "0",
     "Si la fila es por réplica en lugar de central"),
    ("Control de admisión por memoria", None),
    ("TOK_PET_RESERVA", "Tokens que se reservan al admitir una petición", "=TOK_ENTRADA+SALIDA_MAX", "tokens", "#,##0",
     "Entrada real + salida máxima permitida, para no quedarse sin memoria a mitad de la respuesta"),
    ("PET_ADMITIDAS_REPL", "Peticiones promedio que se pueden admitir por réplica",
     "=INT(TOK_KV_MAX_BF16/TOK_PET_RESERVA)", "peticiones", "0",
     "Presupuesto de tokens del KV cache ÷ reserva por petición"),
    ("CABE_ADMISION", "¿Alcanza para la carga tras perder un nodo?",
     '=IF(PET_ADMITIDAS_REPL>=CONC_REPL_FALLA,"Sí","No")', "-", "@", "Compara con la hoja Inventario"),
    ("Límites por usuario", None),
    ("PICO_USUARIO", "Tasa máxima de un usuario", "=LIM_MIN_USUARIO/60", "peticiones/s", "0.00", ""),
    ("PESO_USUARIO", "Parte del pico que podría ocupar un solo usuario", "=PICO_USUARIO/RPS_PICO", "%", "0.0%",
     "Un abusador aislado no puede saturar el servicio"),
    ("TOK_USUARIO", "Tokens de KV que puede ocupar un usuario a la vez", "=LIM_CONC_USUARIO*CONTEXTO_MAX", "tokens",
     "#,##0", "Peor caso: todas sus peticiones con la ventana completa"),
]


def build_overload(wb):
    ws = wb.create_sheet("Sobrecarga")
    title(ws, "Sobrecarga, fila y límites",
          "Qué hace el servicio cuando llegan más peticiones de las que puede atender a tiempo.")
    headers = ["Nombre", "Concepto", "Valor", "Unidad", "Cómo se obtiene"]
    ncols = len(headers)
    ws.append([])
    ws.append(headers)
    style_header(ws, ws.max_row, ncols)
    for item in OVERLOAD:
        if item[1] is None:
            ws.append([item[0]])
            style_section(ws, ws.max_row, ncols)
            continue
        key, label, formula, unit, fmt, note = item
        ws.append([key, label, formula, unit, note])
        r = ws.max_row
        ws.cell(row=r, column=3).number_format = fmt
        style_body(ws, r, ncols)
        add_name(wb, key, "Sobrecarga", "C", r)
    for col, width in zip("ABCDE", [20, 52, 14, 14, 70]):
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A5"


NETWORK = [
    ("1. Tráfico de aplicación (usuarios ↔ servicio), en el pico", None),
    ("APP_ENTRADA", "Texto que suben los usuarios", "=RPS_PICO*TOK_ENTRADA*BYTES_TOKEN*8/1E6", "Mb/s", "0.0",
     "Peticiones/s × tokens × bytes × 8 bits. El enunciado da ~375 kB/s para entrada + salida"),
    ("APP_SALIDA", "Respuestas que bajan palabra a palabra", "=TOKS_SAL_PICO*BYTES_EVENTO*8/1E6", "Mb/s", "0.0",
     "Tokens de salida/s × bytes por mensaje"),
    ("APP_TOTAL", "Total de aplicación", "=APP_ENTRADA+APP_SALIDA", "Mb/s", "0.0",
     "Muy poco: cualquier enlace de 1 Gb/s sobra"),
    ("APP_TOTAL_ESTRES", "Total con la demanda ×2", "=APP_TOTAL*FACTOR_CRECIMIENTO", "Mb/s", "0.0", ""),
    ("2. Comunicación entre GPUs de una réplica", None),
    ("TP_BYTES_TOKEN", "Datos que intercambian las 4 GPUs por cada token",
     "=2*CAPAS*HIDDEN*2", "bytes", "#,##0",
     "Aproximación: 2 sincronizaciones por capa × dimensión oculta × 2 bytes"),
    ("TP_TRAFICO", "Intercambio entre GPUs en el punto de operación",
     "=(TOKS_ENT_REPLICA+TOKS_SAL_REPLICA)*TP_BYTES_TOKEN/BYTES_GB", "GB/s", "0.0",
     "Tokens/s de la réplica × datos por token"),
    ("TP_USO_NVLINK", "Fracción de NVLink que usa", "=TP_TRAFICO/NVLINK_BW", "%", "0.0%", "Holgado dentro del nodo"),
    ("TP_USO_RED", "Fracción que usaría si cruzara un enlace de ejemplo", "=TP_TRAFICO/(ENLACE_EJEMPLO/8)", "%",
     "0%", "Por eso la réplica no se reparte entre servidores"),
    ("3. Tráfico de operación: cargar el modelo", None),
    ("CARGA_25", "Cargar los pesos de una réplica por el enlace de ejemplo",
     "=MEM_PESOS_BF16*8/ENLACE_EJEMPLO", "s", "0", "Ideal; el enunciado da 44,8 s con 140 GB"),
    ("CARGA_TODAS", "Cargar las 18 réplicas a la vez desde el almacenamiento central",
     "=REPL_TOTAL*MEM_PESOS_BF16/ALMAC_LECTURA", "s", "0",
     "Arranque en frío de todo el servicio; manda la lectura del almacenamiento"),
    ("CARGA_LOCAL", "¿Caben las versiones del modelo en el disco local de cada DGX?",
     '=IF(VERSIONES_MODELO*MEM_PESOS_BF16/1000<=CACHE_LOCAL_DGX,"Sí","No")', "-", "@",
     "Guardar los pesos en local evita depender de la red al reiniciar un nodo"),
    ("RESTAURAR_FUERA", "Traer una versión del modelo desde la copia fuera del sitio",
     "=MEM_PESOS_BF16*8/ENLACE_RESPALDO/60", "min", "0",
     "Solo si se pierden el almacenamiento central y las copias locales"),
    ("4. Presupuesto de almacenamiento central", None),
    ("ALM_MODELOS", "Versiones del modelo", "=VERSIONES_MODELO*MEM_PESOS_BF16", "GB", "#,##0", ""),
    ("ALM_IMAGENES", "Imágenes y artefactos", "=IMAGENES_GB", "GB", "#,##0", ""),
    ("ALM_DATOS", "Datos de la aplicación", "=DATOS_APP_GB", "GB", "#,##0", ""),
    ("ALM_LOGS", "Registros y métricas", "=PET_DIA*LOG_BYTES_PET*RETENCION_DIAS/BYTES_GB", "GB", "#,##0.0",
     "Peticiones/día × bytes × días de retención"),
    ("ALM_UTIL", "Capacidad útil necesaria", "=ALM_MODELOS+ALM_IMAGENES+ALM_DATOS+ALM_LOGS", "GB", "#,##0", ""),
    ("ALM_BRUTO", "Capacidad bruta con copias y holgura", "=ALM_UTIL*REPLICACION*HOLGURA_ALMAC/1000", "TB", "0.0",
     "Útil × copias × holgura"),
    ("ALM_RESPALDO", "Copia fuera del sitio (modelos, configuración, datos)",
     "=(ALM_MODELOS+ALM_DATOS+ALM_IMAGENES)/1000", "TB", "0.00",
     "Los registros no se respaldan fuera: se pueden perder sin afectar el servicio"),
]

STORAGE_TABLE_LINKS = [10, 25, 100, 400]


def build_network(wb):
    ws = wb.create_sheet("Red y almacenamiento")
    title(ws, "Red y almacenamiento",
          "Tres tipos de tráfico muy distintos: el de los usuarios es pequeño; el pesado es mover el modelo y "
          "comunicar las GPUs.")
    headers = ["Nombre", "Concepto", "Valor", "Unidad", "Cómo se obtiene"]
    ncols = len(headers)
    ws.append([])
    ws.append(headers)
    style_header(ws, ws.max_row, ncols)
    for item in NETWORK:
        if item[1] is None:
            ws.append([item[0]])
            style_section(ws, ws.max_row, ncols)
            continue
        key, label, formula, unit, fmt, note = item
        ws.append([key, label, formula, unit, note])
        r = ws.max_row
        ws.cell(row=r, column=3).number_format = fmt
        style_body(ws, r, ncols)
        add_name(wb, key, "'Red y almacenamiento'", "C", r)

    ws.append([])
    ws.append(["Tiempo ideal para cargar una réplica según el enlace"])
    style_section(ws, ws.max_row, ncols)
    ws.append(["", "Enlace (Gb/s)", "Segundos", "Minutos", "Nota"])
    style_header(ws, ws.max_row, ncols)
    for gbps in STORAGE_TABLE_LINKS:
        r = ws.max_row + 1
        ws.cell(row=r, column=2, value=gbps).font = INPUT_FONT
        ws.cell(row=r, column=3, value=f"=MEM_PESOS_BF16*8/B{r}").number_format = "0"
        ws.cell(row=r, column=4, value=f"=C{r}/60").number_format = "0.0"
        style_body(ws, r, ncols)
    ws.cell(row=ws.max_row, column=5, value="La DGX H100 trae puertos de hasta 400 Gb/s")
    ws.append(["", "Son mínimos teóricos: el tiempo real suma la lectura del disco, los protocolos, cargar en las "
               "GPUs y calentar el motor."])
    ws.cell(row=ws.max_row, column=2).font = NOTE_FONT

    for col, width in zip("ABCDE", [20, 56, 14, 12, 70]):
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A5"


CITIES = ["Querétaro", "São Paulo", "Santiago"]
GPP = "GlobalPetrolPrices, precio para empresas, mar-2026, con impuestos"
EMB = "Ember 2025 vía Carbon Recycling / GreenCalculus (fuente secundaria)"
CBRE = "CBRE Global Data Center Trends 2026 (Q1 2026, contratos de 250–500 kW)"
WN = "WonderNetwork, promedio entre servidores, 13-nov-2026"

# (key, label, values per city or template, unit, fmt, source)
LOCATION = [
    ("Datos de cada ciudad (azul = dato de fuente)", None),
    ("in_precio", "Precio de la electricidad", [0.199, 0.146, 0.197], "USD/kWh", "0.000", GPP),
    ("in_emision", "Emisiones de la red eléctrica", [0.474, 0.110, 0.291], "kg CO₂e/kWh", "0.000", EMB),
    ("in_renta", "Renta de espacio (punto medio del rango)", [260, 160, 200], "USD/kW/mes", "0", CBRE),
    ("in_disp", "Capacidad disponible en el mercado", [31.5, 51.5, 5.4], "MW", "0.0", CBRE),
    ("in_ping", "Peor ping a una capital hispana de la región", [193, 158, 178], "ms", "0",
     WN + "; QRO usa Ciudad de México como referencia"),
    ("in_peor", "Ciudad del peor caso", ["Bogotá", "Ciudad de México", "Ciudad de México"], "-", "@", WN),
    ("in_sismo", "Riesgo sísmico", ["Medio", "Bajo", "Alto"], "-", "@", "Apreciación general; verificar con mapas oficiales"),
    ("Resultados", None),
    ("UB_ENERGIA", "Energía al año", "=(NODOS*P_NODO_PROM+P_AUX_PROM)*PUE*HORAS_ANIO/1000", "MWh/año", "#,##0",
     "Igual en las tres: depende del equipo, no de la ciudad"),
    ("UB_COSTO_ELEC", "Costo de electricidad", "={c}{UB_ENERGIA}*1000*{c}{in_precio}", "USD/año", "#,##0", ""),
    ("UB_KW_CONTRATO", "Potencia a contratar", "=NODOS*P_NODO_MAX+P_AUX_MAX", "kW", "0.0", "Potencia máxima de TI"),
    ("UB_RENTA", "Renta de espacio", "={c}{UB_KW_CONTRATO}*{c}{in_renta}*12", "USD/año", "#,##0",
     "Puede ser mayor: nuestro contrato es más chico que 250 kW"),
    ("UB_TOTAL", "Electricidad + renta", "={c}{UB_COSTO_ELEC}+{c}{UB_RENTA}", "USD/año", "#,##0", ""),
    ("UB_CO2", "Emisiones de operación", "={c}{UB_ENERGIA}*{c}{in_emision}", "t CO₂e/año", "#,##0",
     "MWh × kg/kWh = toneladas"),
    ("UB_RTT", "Peor tiempo de red para un usuario", "={c}{in_ping}+ULTIMA_MILLA", "ms", "0", "Ping + última milla"),
    ("UB_RTT_OK", "¿Cabe en la red supuesta para el TTFT?", '=IF({c}{UB_RTT}<=RED_IDA_VUELTA*1000,"Sí","No")',
     "-", "@", "Hoja Sobrecarga: 0,2 s"),
    ("UB_DISP_OK", "¿Hay capacidad disponible para nosotros?", '=IF({c}{in_disp}*1000>={c}{UB_KW_CONTRATO}*PUE,"Sí","No")',
     "-", "@", "Necesitamos ~0,14 MW con enfriamiento"),
]

PHYSICAL = [
    ("Diseño físico en el sitio elegido", None),
    ("RACK_U_NODO", "Altura de un DGX H100", "=8", "U", "0", "Guía de usuario DGX H100: 8U"),
    ("RACK_U_OCUP", "Altura ocupada por rack de cómputo", "=NODOS_POR_RACK*RACK_U_NODO", "U de 42", "0",
     "Sobra espacio: el límite no es el espacio sino la energía y el calor"),
    ("RACK_KW", "Potencia por rack de cómputo", "=NODOS_POR_RACK*P_NODO_MAX", "kW", "0.0",
     "Hay que confirmar que el proveedor soporte esta densidad con aire"),
    ("RACK_KG", "Peso de los servidores por rack", "=NODOS_POR_RACK*130.45", "kg", "0",
     "Guía DGX H100: 130,45 kg máx. por equipo; falta sumar el rack"),
    ("RACK_CFM", "Aire que mueven los servidores de un rack", "=NODOS_POR_RACK*1105", "CFM", "#,##0",
     "Guía DGX H100: 1.105 CFM por equipo, de frente a atrás"),
    ("RACK_TOTAL", "Racks en total", "=RACKS_COMPUTO+1", "racks", "0", "3 de cómputo + 1 de red y almacenamiento"),
    ("RACK_ALT", "Si el proveedor solo da ~25 kW por rack", "=ROUNDUP(NODOS/INT(25/P_NODO_MAX),0)+1", "racks",
     "0", "2 servidores por rack: más racks, pero perder uno quita menos réplicas"),
]

EXOTIC = [
    ("Opciones descartadas con números (§12.3–12.5)", None),
    ("MARTE_RTT", "Marte: ida y vuelta mínima de la señal", "=2*3*60", "s", "#,##0",
     "NASA: 3 a 22 min solo de ida. Contra un TTFT de 2 s: imposible para usuarios en la Tierra"),
    ("MARTE_RTT_MAX", "Marte: ida y vuelta máxima", "=2*22*60", "s", "#,##0", ""),
    ("ORB_P", "Órbita: calor a disipar (TI máxima)", "=NODOS*P_NODO_MAX+P_AUX_MAX", "kW", "0.0", ""),
    ("ORB_AREA", "Órbita: área mínima de radiadores",
     "=ORB_P*1000/(0.9*5.670374E-8*(300^4-3^4))", "m²", "0",
     "Stefan-Boltzmann ideal, ε = 0,9, 300 K, espacio profundo. El enunciado da ~251 m²; sin contar sol, Tierra ni masa"),
]

EXOTIC_TEXT = [
    "Submarino: el mar recibe el calor, pero hay que diseñar toda la cadena de intercambio, sellado y corrosión. "
    "Con 72 GPUs que fallan de vez en cuando, cada reparación obliga a sacar la cápsula. Project Natick es un "
    "experimento, no una garantía.",
    "Órbita: además de los radiadores, hay radiación que daña la electrónica, eclipses sin sol, enlaces limitados con "
    "tierra y costo de lanzamiento y reemplazo. Nada de eso cabe en el presupuesto de USD 8 millones.",
    "Marte: descartado por física. Solo serviría para usuarios en Marte o trabajos por lotes, no para este chat.",
]


def build_location(wb):
    ws = wb.create_sheet("Ubicación")
    title(ws, "Ubicación y centro de datos",
          "Colocation: el proveedor pone edificio, electricidad, enfriamiento y seguridad; nosotros los servidores.")
    ncols = 7
    ws.append([])
    ws.append(["Nombre", "Concepto"] + CITIES + ["Unidad", "Fuente o cómo se obtiene"])
    style_header(ws, ws.max_row, ncols)
    hdr = ws.max_row
    cols = ["C", "D", "E"]

    rows, r = {}, ws.max_row
    for item in LOCATION:
        r += 1
        if item[1] is not None:
            rows[item[0]] = r
    for item in LOCATION:
        if item[1] is None:
            ws.append([item[0]])
            style_section(ws, ws.max_row, ncols)
            continue
        key, label, vals, unit, fmt, src = item
        r = ws.max_row + 1
        ws.cell(row=r, column=1, value="" if key.startswith("in_") else key)
        ws.cell(row=r, column=2, value=label)
        for i, col in enumerate(cols):
            cell = ws.cell(row=r, column=3 + i)
            if isinstance(vals, list):
                cell.value = vals[i]
                cell.font = INPUT_FONT
            else:
                cell.value = vals.format(c=col, **rows)
            cell.number_format = fmt
        ws.cell(row=r, column=6, value=unit)
        ws.cell(row=r, column=7, value=src)
        style_body(ws, r, ncols)
        if not key.startswith("in_"):
            # value for the chosen city
            add_name(wb, key, "Ubicación", "C", r)
            wb.defined_names[key] = DefinedName(
                key, attr_text=f"INDEX(Ubicación!$C${r}:$E${r},MATCH(UBICACION,Ubicación!$C${hdr}:$E${hdr},0))")
    # chosen-city inputs too
    for key, name in [("in_precio", "PRECIO_KWH"), ("in_emision", "FACTOR_EMISION"), ("in_renta", "RENTA_KW_MES")]:
        rr = rows[key]
        wb.defined_names[name] = DefinedName(
            name, attr_text=f"INDEX(Ubicación!$C${rr}:$E${rr},MATCH(UBICACION,Ubicación!$C${hdr}:$E${hdr},0))")

    for block in (PHYSICAL, EXOTIC):
        ws.append([])
        for item in block:
            if item[1] is None:
                ws.append([item[0]])
                style_section(ws, ws.max_row, ncols)
                continue
            key, label, formula, unit, fmt, note = item
            ws.append([key, label, formula, "", "", unit, note])
            r = ws.max_row
            ws.cell(row=r, column=3).number_format = fmt
            style_body(ws, r, ncols)
            add_name(wb, key, "Ubicación", "C", r)
    for text in EXOTIC_TEXT:
        ws.append(["", text])
        r = ws.max_row
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
        ws.cell(row=r, column=2).alignment = WRAP
        ws.row_dimensions[r].height = 32

    ws.append([])
    ws.append(["", "Fuentes: GlobalPetrolPrices (mar-2026), CBRE Global Data Center Trends 2026, WonderNetwork "
               "(13-nov-2026), Ember 2025 vía agregadores, guía de usuario DGX H100/H200 y NASA. Consultadas el "
               "13-nov-2026."])
    ws.cell(row=ws.max_row, column=2).font = NOTE_FONT

    for col, width in zip("ABCDEFG", [16, 46, 16, 16, 16, 14, 62]):
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "C5"


# Three operating cases side by side: normal (average), maximum, degraded (one node and one cooling module out).
ENERGY_CASES = ["Normal (promedio)", "Máxima", "Degradada (−1 nodo, −1 módulo)"]
# (key, label, [normal, max, degraded] formulas or single formula, unit, fmt, note)
ENERGY = [
    ("Balance eléctrico", None),
    ("E_NODOS_ON", "Nodos encendidos", ["=NODOS", "=NODOS", "=NODOS-NODOS_PERDIDOS"], "nodos", "0", ""),
    ("E_P_NODO", "Potencia por nodo", ["=P_NODO_PROM", "=P_NODO_MAX", "=P_NODO_MAX"], "kW", "0.0",
     "En la degradada, los nodos que quedan trabajan más"),
    ("E_P_AUX", "Red, almacenamiento y servicios", ["=P_AUX_PROM", "=P_AUX_MAX", "=P_AUX_MAX"], "kW", "0.0", ""),
    ("E_P_TI", "Potencia de TI", "={c}{E_NODOS_ON}*{c}{E_P_NODO}+{c}{E_P_AUX}", "kW", "0.0",
     "La potencia del servidor ya incluye sus GPUs: no se suman otra vez"),
    ("E_P_SITIO", "Potencia total del sitio",
     ["=E_P_TI_N*PUE", "=E_P_TI_M+P_NO_TI", "=E_P_TI_D+P_NO_TI"], "kW", "0.0",
     "Normal: TI × PUE. Máxima y degradada: TI + cargas no TI del punto de diseño"),
    ("E_HOLG_ACOM", "Holgura frente a la potencia contratada", "=1-{c}{E_P_SITIO}/ACOMETIDA", "%", "0%",
     "Falta revisar arranques, derrateo y mantenimiento"),
    ("Balance térmico", None),
    ("E_CALOR", "Calor a retirar (≈ potencia de TI)", "={c}{E_P_TI}", "kW", "0.0",
     "Casi toda la electricidad de TI termina como calor"),
    ("E_FRIO_DISP", "Enfriamiento disponible",
     ["=N_MODULOS*MODULO_FRIO", "=N_MODULOS*MODULO_FRIO", "=(N_MODULOS-1)*MODULO_FRIO"], "kW", "0", ""),
    ("E_FRIO_OK", "¿Alcanza el enfriamiento?", '=IF({c}{E_FRIO_DISP}>={c}{E_CALOR},"Sí","No")', "-", "@", ""),
    ("E_FRIO_USO", "Uso del enfriamiento", "={c}{E_CALOR}/{c}{E_FRIO_DISP}", "%", "0%", ""),
    ("E_CAUDAL", "Caudal de agua del circuito", "={c}{E_CALOR}*1000/(CP_AGUA*DELTA_T)", "kg/s", "0.00",
     "Calor ÷ (calor específico × subida de temperatura)"),
    ("E_CAUDAL_LMIN", "Caudal en litros por minuto", "={c}{E_CAUDAL}*60", "L/min", "0",
     "Agua: 1 kg ≈ 1 L. El enunciado da ~149 L/min"),
]

ENERGY_SINGLE = [
    ("Energía y costo anual (perfil promedio)", None),
    ("E_ANUAL", "Energía del sitio al año", "=E_P_SITIO_N*HORAS_ANIO", "kWh/año", "#,##0",
     "El enunciado da ~799.438 kWh"),
    ("E_ANUAL_TI", "Energía de TI al año", "=E_P_TI_N*HORAS_ANIO", "kWh/año", "#,##0", ""),
    ("E_COSTO", "Costo de la electricidad", "=E_ANUAL*PRECIO_KWH", "USD/año", "#,##0",
     "Con el precio de la ubicación elegida"),
    ("Respaldo (UPS)", None),
    ("UPS_UTIL", "Energía útil necesaria", "=UPS_KW*UPS_MIN/60", "kWh", "0.0", "Carga × tiempo"),
    ("UPS_NOMINAL", "Capacidad nominal de baterías", "=UPS_UTIL/(UPS_EFIC*UPS_DOD*UPS_ENVEJ)", "kWh", "0.0",
     "El enunciado da ~35 kWh"),
    ("UPS_CUBRE", "¿La UPS cubre la TI máxima?", '=IF(UPS_KW>=E_P_TI_M,"Sí","No")', "-", "@",
     "Las baterías también deben poder entregar esa potencia, no solo la energía"),
    ("Revisión del dato de potencia", None),
    ("DGX_KW_CALOR", "Calor declarado por DGX convertido a kW", "=DGX_BTU/3412.14", "kW", "0.0",
     "1 kW = 3.412,14 BTU/h. Da más que los 10,2 kW de la entrada"),
    ("TI_MAX_CALOR", "TI máxima si cada DGX disipara ese valor", "=NODOS*DGX_KW_CALOR+P_AUX_MAX", "kW", "0.0", ""),
    ("FRIO_OK_CALOR", "¿Alcanza el enfriamiento con un módulo fuera?",
     '=IF((N_MODULOS-1)*MODULO_FRIO>=TI_MAX_CALOR,"Sí","No")', "-", "@", "Sigue alcanzando: el diseño aguanta la duda"),
    ("Impacto ambiental", None),
    ("AMB_CO2", "Emisiones de operación", "=E_ANUAL*FACTOR_EMISION/1000", "t CO₂e/año", "#,##0",
     "Energía × factor de la red de la ubicación"),
    ("AMB_AGUA", "Agua directa", "=E_ANUAL_TI*WUE/1000", "m³/año", "#,##0", "Con el WUE de Entradas"),
    ("AMB_AGUA_1", "Agua directa si WUE = 1 L/kWh", "=E_ANUAL_TI*1/1000", "m³/año", "#,##0", "Sensibilidad"),
    ("AMB_AGUA_2", "Agua directa si WUE = 2 L/kWh", "=E_ANUAL_TI*2/1000", "m³/año", "#,##0", "Sensibilidad"),
    ("AMB_RESIDUOS", "Equipo de cómputo a reemplazar al final de su vida", "=NODOS*130.45/1000", "t", "0.00",
     "9 DGX × 130,45 kg, cada VIDA_EQUIPO años; hay que reciclarlo"),
    ("Crecimiento (demanda ×2)", None),
    ("E_TI_X2", "TI máxima con demanda ×2", "=ESC_NODOS_X2*P_NODO_MAX+P_AUX_MAX", "kW", "0.0",
     "Con los nodos del escenario ×2 de la hoja Inventario"),
    ("E_SITIO_X2", "Sitio máximo con demanda ×2", "=E_TI_X2+P_NO_TI*E_TI_X2/E_P_TI_M", "kW", "0.0",
     "Supone que las cargas no TI crecen en proporción"),
    ("E_ACOM_X2", "¿Cabe en la potencia contratada?", '=IF(E_SITIO_X2<=ACOMETIDA,"Sí","No")', "-", "@",
     "Si no cabe, crecer exige más potencia y enfriamiento, no solo servidores"),
]


def build_energy(wb):
    ws = wb.create_sheet("Energía")
    title(ws, "Electricidad, energía y enfriamiento",
          "kW es potencia (cuánto en un instante). kWh es energía (potencia × tiempo). Tres casos de operación.")
    ncols = 7
    ws.append([])
    ws.append(["Nombre", "Concepto"] + ENERGY_CASES + ["Unidad", "Cómo se obtiene"])
    style_header(ws, ws.max_row, ncols)
    ws.row_dimensions[ws.max_row].height = 32
    cols = ["C", "D", "E"]
    suffix = ["_N", "_M", "_D"]

    rows, r = {}, ws.max_row
    for item in ENERGY:
        r += 1
        if item[1] is not None:
            rows[item[0]] = r
    for item in ENERGY:
        if item[1] is None:
            ws.append([item[0]])
            style_section(ws, ws.max_row, ncols)
            continue
        key, label, f, unit, fmt, note = item
        r = ws.max_row + 1
        ws.cell(row=r, column=1, value=key)
        ws.cell(row=r, column=2, value=label)
        for i, col in enumerate(cols):
            val = f[i] if isinstance(f, list) else f.format(c=col, **rows)
            ws.cell(row=r, column=3 + i, value=val).number_format = fmt
            add_name(wb, key + suffix[i], "Energía", col, r)
        ws.cell(row=r, column=6, value=unit)
        ws.cell(row=r, column=7, value=note)
        style_body(ws, r, ncols)

    ws.append([])
    for item in ENERGY_SINGLE:
        if item[1] is None:
            ws.append([item[0]])
            style_section(ws, ws.max_row, ncols)
            continue
        key, label, formula, unit, fmt, note = item
        ws.append([key, label, formula, "", "", unit, note])
        r = ws.max_row
        ws.cell(row=r, column=3).number_format = fmt
        style_body(ws, r, ncols)
        add_name(wb, key, "Energía", "C", r)

    for text in [
        "Los DGX H100 se enfrían por aire. El aire caliente pasa a las unidades de enfriamiento del proveedor, que "
        "lo llevan al circuito de agua y de ahí al exterior. Con ~31 kW por rack hace falta pasillo caliente "
        "contenido o intercambiadores en la puerta trasera.",
        "N+1 en módulos no prueba redundancia en bombas, tuberías, controles ni alimentación: eso se le exige al "
        "proveedor por contrato y se revisa en la visita.",
        "Un PUE bajo no significa bajas emisiones ni poca agua: son tres indicadores distintos.",
    ]:
        ws.append(["", text])
        r = ws.max_row
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
        ws.cell(row=r, column=2).alignment = WRAP
        ws.cell(row=r, column=2).font = NOTE_FONT
        ws.row_dimensions[r].height = 32

    for col, width in zip("ABCDEFG", [16, 48, 16, 16, 18, 12, 62]):
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "C5"


EJ = "Enunciado §14 (valor hipotético, no es cotización)"
PRICE_NOTE = "USD, nov-2026, sin impuestos de importación (se suman aparte)"

# CAPEX lines: (name, label, qty formula, unit price, imported?, source, scope note)
CAPEX = [
    ("CX_DGX", "Servidores DGX H100 (8 GPUs)", "=NODOS", 350_000, True,
     "Enunciado §14.1; rango público USD 300–400 mil (sin precio oficial de NVIDIA)", "Incluye garantía estándar"),
    ("CX_RED", "Red y ópticas", "=1", 350_000, True, EJ, "Switches, cables y transceptores"),
    ("CX_ALM", "Almacenamiento y servidores de soporte", "=1", 250_000, True, EJ,
     "~3,9 TB brutos + 3 servidores de servicios"),
    ("CX_RACK", "Racks y contención", "=RACK_TOTAL", 37_500, False, EJ + " (150.000 ÷ 4 racks)",
     "Pasillo contenido para ~31 kW por rack"),
    ("CX_ADEC", "Adecuación del espacio en colocation", "=1", 100_000, False, "Estimación propia",
     "Jaula, PDUs de rack, cableado; reemplaza la obra eléctrica, UPS, enfriamiento y edificio del ejemplo"),
    ("CX_INTEG", "Integración y puesta en marcha", "=1", 200_000, False, EJ, "Servicio local"),
    ("CX_REP", "Repuestos iniciales", "=1", 150_000, True, EJ, "GPUs, fuentes, discos, ópticas"),
]

# OPEX lines: (name, label, formula, source)
OPEX = [
    ("OX_ELEC", "Electricidad", "=E_COSTO", "Hoja Energía: 799 MWh × precio de São Paulo (con impuestos)"),
    ("OX_RENTA", "Renta de colocation (espacio, enfriamiento, UPS)", "=UB_RENTA",
     "Hoja Ubicación: 103,8 kW × USD/kW/mes de CBRE"),
    ("OX_PERSONAL", "Personal y servicios de operación", 600_000, EJ + "; guardia 24/7 con turnos y soporte remoto"),
    ("OX_CONECT", "Conectividad (2 proveedores de internet)", 60_000, EJ),
    ("OX_MANT", "Mantenimiento no cubierto por la garantía", 120_000, EJ),
    ("OX_SOFT", "Software, observabilidad y seguridad", 60_000, EJ),
    ("OX_SEGURO", "Seguros y cumplimiento", 36_000, EJ + " (sin agua: la pone el proveedor)"),
    ("OX_RESP", "Copia fuera del sitio y reserva de recuperación", 48_000, EJ),
    ("OX_CONT", "Contingencia operativa", 120_000, EJ),
]


def build_budget(wb):
    ws = wb.create_sheet("Presupuesto")
    title(ws, "Presupuesto: CAPEX, OPEX y TCO",
          "USD, precios hipotéticos del enunciado salvo donde se indica. Azul = precio escrito a mano.")
    ncols = 7
    ws.append([])
    ws.append(["Nombre", "Partida", "Cantidad", "Precio unitario", "Total", "¿Importado?", "Fuente y alcance"])
    style_header(ws, ws.max_row, ncols)

    ws.append(["CAPEX (compra inicial)"])
    style_section(ws, ws.max_row, ncols)
    first = ws.max_row + 1
    for name, label, qty, price, imported, src, scope in CAPEX:
        r = ws.max_row + 1
        ws.cell(row=r, column=1, value=name)
        ws.cell(row=r, column=2, value=label)
        ws.cell(row=r, column=3, value=qty).number_format = "0"
        c = ws.cell(row=r, column=4, value=price)
        c.font = INPUT_FONT
        c.number_format = "#,##0"
        ws.cell(row=r, column=5, value=f"=C{r}*D{r}").number_format = "#,##0"
        ws.cell(row=r, column=6, value="Sí" if imported else "No")
        ws.cell(row=r, column=7, value=f"{src}. {scope}. {PRICE_NOTE}.")
        style_body(ws, r, ncols)
        add_name(wb, name, "Presupuesto", "E", r)
        add_name(wb, name + "_P", "Presupuesto", "D", r)
    last = ws.max_row

    def line(name, label, formula, fmt="#,##0", note="", input_val=False):
        r = ws.max_row + 1
        ws.cell(row=r, column=1, value=name)
        ws.cell(row=r, column=2, value=label)
        c = ws.cell(row=r, column=5, value=formula)
        c.number_format = fmt
        if input_val:
            c.font = INPUT_FONT
        ws.cell(row=r, column=7, value=note)
        style_body(ws, r, ncols)
        add_name(wb, name, "Presupuesto", "E", r)

    line("CX_IMPORTADO", "Equipo importado", f'=SUMIF(F{first}:F{last},"Sí",E{first}:E{last})')
    line("IMPUESTO_IMPORT", "Carga de impuestos de importación en Brasil", 0.50, "0%",
         "Supuesto prudente: fuentes hablan de >50 % antes de ReData; arancel de 18–25 % en servidores grandes",
         input_val=True)
    line("CX_IMPUESTOS", "Impuestos de importación", "=CX_IMPORTADO*IMPUESTO_IMPORT")
    line("CX_SUBTOTAL", "Subtotal", f"=SUM(E{first}:E{last})+CX_IMPUESTOS")
    line("CONTINGENCIA", "Contingencia", 0.20, "0%", "Enunciado §14.1", input_val=True)
    line("CX_CONT", "Monto de contingencia", "=CX_SUBTOTAL*CONTINGENCIA")
    line("CAPEX", "CAPEX total", "=CX_SUBTOTAL+CX_CONT")
    for c in range(1, ncols + 1):
        ws.cell(row=ws.max_row, column=c).font = Font(bold=True)

    ws.append([])
    ws.append(["OPEX (cada año)"])
    style_section(ws, ws.max_row, ncols)
    o_first = ws.max_row + 1
    for name, label, val, src in OPEX:
        r = ws.max_row + 1
        ws.cell(row=r, column=1, value=name)
        ws.cell(row=r, column=2, value=label)
        c = ws.cell(row=r, column=5, value=val)
        c.number_format = "#,##0"
        if not isinstance(val, str):
            c.font = INPUT_FONT
        ws.cell(row=r, column=7, value=f"{src}. USD/año, nov-2026.")
        style_body(ws, r, ncols)
        add_name(wb, name, "Presupuesto", "E", r)
    o_last = ws.max_row
    line("OX_FIJO", "Parte fija del OPEX (todo menos electricidad y renta)", f"=SUM(E{o_first+2}:E{o_last})")
    line("OPEX", "OPEX anual", f"=SUM(E{o_first}:E{o_last})")
    for c in range(1, ncols + 1):
        ws.cell(row=ws.max_row, column=c).font = Font(bold=True)

    ws.append([])
    ws.append(["TCO e indicadores (demanda base constante, sin descuento financiero ni valor residual)"])
    style_section(ws, ws.max_row, ncols)
    line("TCO", "TCO a 3 años", "=CAPEX+ANIOS_TCO*OPEX", note="CAPEX + años × OPEX")
    line("PET_PERIODO", "Peticiones en el periodo", "=PET_DIA*365*ANIOS_TCO", note="El enunciado da 328,5 millones")
    line("COSTO_PET", "Costo por petición", "=TCO/PET_PERIODO", "0.0000", "Supone que todas se completan")
    line("FRAC_BUENAS", "Peticiones a tiempo y sin error (cota)", "=SLO_DISPONIBILIDAD*(1-SLO_ERRORES)*0.95", "0.0%",
         "Disponibilidad × (1 − errores) × 95 % dentro del p95. La calidad de las respuestas está pendiente de medir")
    line("COSTO_PET_BUENA", "Costo por petición correcta y a tiempo", "=COSTO_PET/FRAC_BUENAS", "0.0000", "")
    line("COSTO_MTOK", "Costo por millón de tokens (entrada + salida)",
         "=TCO/(PET_PERIODO*(TOK_ENTRADA+TOK_SALIDA)/1E6)", "0.00",
         "Denominador: tokens de entrada y salida. No comparar con tarifas de API que cobran solo una parte")
    line("COSTO_MTOK_SAL", "Costo por millón de tokens de salida (todo el costo cargado a la salida)",
         "=TCO/(PET_PERIODO*TOK_SALIDA/1E6)", "0.00", "")

    ws.append([])
    ws.append(["Límites del enunciado (§4.5)"])
    style_section(ws, ws.max_row, ncols)
    for name, label, val, lim in [("OK_CAPEX", "CAPEX", "CAPEX", "CAPEX_MAX"),
                                  ("OK_OPEX", "OPEX anual", "OPEX", "OPEX_MAX"),
                                  ("OK_TCO", "TCO a 3 años", "TCO", "TCO_MAX")]:
        r = ws.max_row + 1
        ws.cell(row=r, column=1, value=name)
        ws.cell(row=r, column=2, value=label)
        ws.cell(row=r, column=3, value=f"={val}").number_format = "#,##0"
        ws.cell(row=r, column=4, value=f"={lim}").number_format = "#,##0"
        ws.cell(row=r, column=5, value=f"={val}/{lim}").number_format = "0%"
        ws.cell(row=r, column=6, value=f'=IF({val}<={lim},"Cumple","No cumple")')
        ws.cell(row=r, column=7, value="Valor · límite · uso del límite · resultado")
        style_body(ws, r, ncols)

    for col, width in zip("ABCDEFG", [17, 48, 12, 15, 15, 12, 70]):
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "C5"


# Scenarios: (label, nodes, hw factor, tax, energy factor, demand factor, note)
SCEN = [
    ("Base", "=NODOS", 1, "=IMPUESTO_IMPORT", 1, 1, ""),
    ("Demanda ×2", "=ESC_NODOS_X2", 1, "=IMPUESTO_IMPORT", 1, "=FACTOR_CRECIMIENTO",
     "Además exige más potencia contratada (hoja Energía)"),
    ("Rendimiento −30 %", "=ESC_NODOS_BAJA", 1, "=IMPUESTO_IMPORT", 1, 1, ""),
    ("Energía +50 %", "=NODOS", 1, "=IMPUESTO_IMPORT", 1.5, 1, ""),
    ("Hardware +30 %", "=NODOS", 1.3, "=IMPUESTO_IMPORT", 1, 1, ""),
    ("Entrada ×2", "=ROUNDUP(RPS_PICO/(CAP_REPLICA*TOK_POR_PET/(2*TOK_ENTRADA+TOK_SALIDA)*UTIL_DISENO*REPL_POR_NODO),0)+NODOS_PERDIDOS",
     1, "=IMPUESTO_IMPORT", 1, 1,
     "Supuesto: capacidad inversa a los tokens. Además no cabe en la ventana de 8.192"),
    ("Tolerar un rack", "=ESC_NODOS_BASE-NODOS_PERDIDOS+NODOS_POR_RACK", 1, "=IMPUESTO_IMPORT", 1, 1,
     "Para promesa de rack completo (hoja Inventario)"),
    ("Con ReData (20 %)", "=NODOS", 1, 0.20, 1, 1, "Si se califica al régimen especial"),
]


def build_scenarios(wb):
    ws = wb.create_sheet("Escenarios")
    title(ws, "Escenarios de costo (§14.4)",
          "Cada columna recalcula nodos, CAPEX, OPEX y TCO. Azul = lo que cambia en ese escenario.")
    n = len(SCEN)
    cols = [chr(ord("C") + i) for i in range(n)]
    ncols = 2 + n
    ws.append([])
    ws.append(["Nombre", "Concepto"] + [s[0] for s in SCEN])
    style_header(ws, ws.max_row, ncols)
    ws.row_dimensions[ws.max_row].height = 32

    rows = {}
    spec = [
        ("Lo que cambia", None),
        ("S_NODOS", "Nodos", [s[1] for s in SCEN], "0"),
        ("S_FHW", "Factor de precio del hardware", [s[2] for s in SCEN], "0.00"),
        ("S_IMP", "Impuestos de importación", [s[3] for s in SCEN], "0%"),
        ("S_FENER", "Factor de precio de la energía", [s[4] for s in SCEN], "0.00"),
        ("S_FDEM", "Factor de demanda", [s[5] for s in SCEN], "0.00"),
        ("Resultados", None),
        ("S_RACKS", "Racks", "=ROUNDUP({c}{S_NODOS}/NODOS_POR_RACK,0)+1", "0"),
        ("S_CAPEX", "CAPEX",
         "=(1+CONTINGENCIA)*(({c}{S_NODOS}*CX_DGX_P+CX_RED+CX_ALM+CX_REP)*{c}{S_FHW}*(1+{c}{S_IMP})"
         "+{c}{S_RACKS}*CX_RACK_P*{c}{S_FHW}+CX_ADEC+CX_INTEG)", "#,##0"),
        ("S_ELEC", "Electricidad al año",
         "=({c}{S_NODOS}*P_NODO_PROM+P_AUX_PROM)*PUE*HORAS_ANIO*PRECIO_KWH*{c}{S_FENER}", "#,##0"),
        ("S_RENTA", "Renta al año", "=({c}{S_NODOS}*P_NODO_MAX+P_AUX_MAX)*RENTA_KW_MES*12", "#,##0"),
        ("S_OPEX", "OPEX anual", "={c}{S_ELEC}+{c}{S_RENTA}+OX_FIJO", "#,##0"),
        ("S_TCO", "TCO a 3 años", "={c}{S_CAPEX}+ANIOS_TCO*{c}{S_OPEX}", "#,##0"),
        ("S_COSTO_PET", "Costo por petición", "={c}{S_TCO}/(PET_PERIODO*{c}{S_FDEM})", "0.0000"),
        ("S_DIF", "Diferencia de TCO con la base", "={c}{S_TCO}-$C${S_TCO}", "#,##0"),
        ("Contra los límites", None),
        ("S_OK_CAPEX", "¿CAPEX ≤ USD 8 M?", '=IF({c}{S_CAPEX}<=CAPEX_MAX,"Sí","No")', "@"),
        ("S_OK_OPEX", "¿OPEX ≤ USD 1,5 M?", '=IF({c}{S_OPEX}<=OPEX_MAX,"Sí","No")', "@"),
        ("S_OK_TCO", "¿TCO ≤ USD 12,5 M?", '=IF({c}{S_TCO}<=TCO_MAX,"Sí","No")', "@"),
    ]
    r = ws.max_row
    for item in spec:
        r += 1
        if item[1] is not None:
            rows[item[0]] = r
    for item in spec:
        if item[1] is None:
            ws.append([item[0]])
            style_section(ws, ws.max_row, ncols)
            continue
        key, label, f, fmt = item
        r = ws.max_row + 1
        ws.cell(row=r, column=1, value=key)
        ws.cell(row=r, column=2, value=label)
        for i, col in enumerate(cols):
            if isinstance(f, list):
                v = f[i]
                c = ws.cell(row=r, column=3 + i, value=v)
                if i > 0 and str(v) != str(f[0]):
                    c.font = INPUT_FONT
            else:
                c = ws.cell(row=r, column=3 + i, value=f.format(c=col, **rows))
            c.number_format = fmt
        style_body(ws, r, ncols)
        add_name(wb, key + "_BASE", "Escenarios", "C", r)

    ws.append([])
    ws.append(["", "Variable que más mueve el TCO:"])
    r = ws.max_row
    rng = f"D{rows['S_DIF']}:{cols[-1]}{rows['S_DIF']}"
    ws.cell(row=r, column=3, value=f"=INDEX(D4:{cols[-1]}4,MATCH(MAX({rng}),{rng},0))")
    ws.cell(row=r, column=2).font = Font(bold=True)
    notes = [s for s in SCEN if s[6]]
    for label, *_, note in notes:
        ws.append(["", f"{label}: {note}."])
        ws.cell(row=ws.max_row, column=2).font = NOTE_FONT
    ws.append(["", "Pérdida de un nodo o de un rack: no cambia el costo; cambia la capacidad (hoja Inventario). "
               "El escenario \"Tolerar un rack\" muestra cuánto cuesta prometerlo."])
    ws.cell(row=ws.max_row, column=2).font = NOTE_FONT

    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 34
    for col in cols:
        ws.column_dimensions[col].width = 14
    ws.freeze_panes = "C5"


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
    build_overload(wb)
    build_network(wb)
    build_location(wb)
    build_energy(wb)
    build_budget(wb)
    build_scenarios(wb)
    build_log(wb)
    for ws in wb.worksheets:
        fit_page(ws)
    wb.calculation = CalcProperties(fullCalcOnLoad=True)
    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT)
    print(f"Saved {OUT.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
