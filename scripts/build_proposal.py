"""Builds propuesta/propuesta_tecnica.docx.

The prose lives here. Tables are read from the Excel model after LibreOffice
recalculates it, so their numbers always match the spreadsheet.
Needs LibreOffice (flatpak org.libreoffice.LibreOffice or `soffice`).
"""
import re
import shutil
import subprocess
import tempfile
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent.parent
XLSX = ROOT / "modelo" / "modelo_calculo.xlsx"
OUT = ROOT / "propuesta" / "propuesta_tecnica.docx"
DIAG = ROOT / "diagramas"

INK = RGBColor(0x2E, 0x2E, 0x2E)
ACCENT = RGBColor(0x1F, 0x38, 0x64)
MUTED = RGBColor(0x6B, 0x64, 0x59)
HEADER_HEX = "1F3864"
ZEBRA_HEX = "F4F1EC"


# ------------------------------------------------------------------ Excel values
def recalculated_workbook():
    tmp = Path(tempfile.mkdtemp())
    if shutil.which("flatpak"):
        cmd = ["flatpak", "run", f"--filesystem={tmp}", "org.libreoffice.LibreOffice"]
    else:
        cmd = ["soffice"]
    subprocess.run(cmd + ["--headless", "--convert-to", "xlsx", "--outdir", str(tmp), str(XLSX)],
                   check=True, capture_output=True, timeout=180)
    return load_workbook(tmp / XLSX.name, data_only=True)


def fmt_num(v, dec=0):
    if v is None:
        return ""
    if isinstance(v, str):
        return v
    s = f"{v:,.{dec}f}"
    return s.replace(",", "§").replace(".", ",").replace("§", ".")


def fmt_pct(v, dec=0):
    return "" if v is None else fmt_num(v * 100, dec) + " %"


# ------------------------------------------------------------------ docx helpers
def shade(cell, hex_color):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tc_pr.append(shd)


def setup(doc):
    st = doc.styles["Normal"]
    st.font.name = "Calibri"
    st.font.size = Pt(11)
    st.font.color.rgb = INK
    st.paragraph_format.space_after = Pt(6)
    st.paragraph_format.line_spacing = 1.15
    for name, size in (("Heading 1", 16), ("Heading 2", 12.5)):
        h = doc.styles[name]
        h.font.name = "Calibri"
        h.font.size = Pt(size)
        h.font.bold = True
        h.font.color.rgb = ACCENT
        h.paragraph_format.space_before = Pt(14 if name == "Heading 1" else 10)
        h.paragraph_format.space_after = Pt(4)
    sec = doc.sections[0]
    sec.left_margin = sec.right_margin = Cm(2.3)
    sec.top_margin = sec.bottom_margin = Cm(2.2)


def para(doc, text, style=None, italic=False, size=None, color=None, align=None, after=None):
    """Paragraph with **bold** fragments."""
    p = doc.add_paragraph(style=style)
    for i, part in enumerate(re.split(r"\*\*", text)):
        if not part:
            continue
        run = p.add_run(part)
        run.bold = i % 2 == 1
        run.italic = italic
        if size:
            run.font.size = Pt(size)
        if color:
            run.font.color.rgb = color
    if align:
        p.alignment = align
    if after is not None:
        p.paragraph_format.space_after = Pt(after)
    return p


def bullets(doc, items):
    for it in items:
        p = para(doc, it, style="List Bullet")
        p.paragraph_format.space_after = Pt(2)


def table(doc, headers, rows, widths=None, size=9):
    t = doc.add_table(rows=1, cols=len(headers))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.style = "Table Grid"
    for i, h in enumerate(headers):
        c = t.rows[0].cells[i]
        c.text = ""
        run = c.paragraphs[0].add_run(h)
        run.bold = True
        run.font.size = Pt(size)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        shade(c, HEADER_HEX)
    for k, row in enumerate(rows):
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = ""
            run = cells[i].paragraphs[0].add_run(str(v))
            run.font.size = Pt(size)
            if k % 2 == 1:
                shade(cells[i], ZEBRA_HEX)
    if widths:
        for row in t.rows:
            for i, w in enumerate(widths):
                row.cells[i].width = Cm(w)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return t


def figure(doc, path, caption, width=16):
    doc.add_picture(str(path), width=Cm(width))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    para(doc, caption, italic=True, size=9, color=MUTED, align=WD_ALIGN_PARAGRAPH.CENTER)


def landscape_section(doc):
    sec = doc.add_section()
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width, sec.page_height = sec.page_height, sec.page_width
    return sec


def portrait_section(doc):
    sec = doc.add_section()
    sec.orientation = WD_ORIENT.PORTRAIT
    sec.page_width, sec.page_height = sec.page_height, sec.page_width
    return sec


# ------------------------------------------------------------------ content
def cover(doc):
    for _ in range(9):
        doc.add_paragraph()
    para(doc, "Diseño de un servicio de LLM que pueda crecer", size=26, color=ACCENT,
         align=WD_ALIGN_PARAGRAPH.CENTER, after=10).runs[0].bold = True
    para(doc, "Propuesta técnica para un asistente conversacional en español", size=13, color=MUTED,
         align=WD_ALIGN_PARAGRAPH.CENTER, after=40)
    para(doc, "Juliana Briceño", size=14, align=WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


def s1(doc):
    doc.add_heading("1. Resumen ejecutivo", 1)
    para(doc, "La propuesta es instalar **9 servidores NVIDIA DGX H100 (72 GPUs) en un centro de datos de colocation en "
              "São Paulo**, sirviendo Llama 3.1 70B Instruct en BF16 a través de 18 réplicas de 4 GPUs cada una.")
    para(doc, "Con eso se atiende el pico de 20,8 peticiones por segundo dejando un 30 % de margen, y se sigue "
              "atendiendo si se cae un servidor completo. El costo queda dentro de los tres límites: CAPEX de USD "
              "7,56 millones, OPEX de USD 1,36 millones al año y TCO a tres años de USD 11,64 millones. Cada "
              "petición cuesta unos 3,5 centavos de dólar.")
    para(doc, "São Paulo ganó a Querétaro y a Santiago por casi todo lo que se midió: electricidad más barata y más "
              "limpia, espacio más barato, más capacidad libre y menos riesgo sísmico. Pero el argumento que más pesa "
              "es otro: es la única de las tres desde donde el usuario más lejano de la región queda dentro del "
              "presupuesto de red que exige el tiempo de primera palabra.")
    para(doc, "Hay que decir con claridad dónde está el riesgo. La capacidad por réplica es un dato **sintético** del "
              "enunciado, no una medición. Si en la realidad una réplica rinde un 30 % menos, o si el hardware cuesta "
              "un 30 % más, el proyecto se sale del presupuesto. Por eso la recomendación final tiene condiciones: "
              "medir antes de comprar y cotizar con proveedores reales.")


def s2(doc):
    doc.add_heading("2. Requisitos y supuestos", 1)
    para(doc, "El servicio es un asistente de texto en español para usuarios de Latinoamérica. Responde preguntas, "
              "resume documentos y ayuda a escribir. Solo se trata de inferencia: no se entrena ningún modelo y "
              "tampoco se incluye búsqueda en documentos (RAG), que queda como posible extensión.")
    para(doc, "Los requisitos vienen del escenario común del enunciado:")
    table(doc, ["Requisito", "Valor"], [
        ("Usuarios activos por día", "200.000"),
        ("Peticiones por usuario al día", "1,5"),
        ("Factor de pico", "6"),
        ("Tokens por petición (perfil controlado)", "4.000 de entrada, 500 de salida"),
        ("Ventana máxima / salida máxima", "8.192 / 1.000 tokens"),
        ("TTFT p95", "≤ 2 s"),
        ("Tiempo por token de salida p95", "≤ 40 ms"),
        ("Tiempo total p95 (500 tokens)", "≤ 30 s"),
        ("Disponibilidad mensual / tasa de error", "99,9 % / ≤ 0,1 %"),
        ("Tolerancia a fallas", "Atender el pico tras perder un nodo"),
        ("Calidad", "≥ 90 % de Llama 3.1 70B Instruct BF16, sin fallas críticas de privacidad"),
        ("CAPEX / OPEX anual / TCO a 3 años", "USD 8 M / 1,5 M / 12,5 M"),
    ], widths=[8, 8.4])
    para(doc, "Cada dato del Excel está clasificado como **documentado**, **medido**, **calculado** o **supuesto**, "
              "con su fuente y su incertidumbre (anexo A). No hay ningún dato medido: no hubo acceso a GPUs. Los "
              "supuestos que más mueven el resultado son estos:")
    bullets(doc, [
        "Una réplica de 4 H100 atiende 2 peticiones por segundo cumpliendo los SLOs (§8.4 del enunciado, sintético).",
        "El servidor consume 6,8 kW en promedio y el PUE anual es 1,30 (ejemplo del §13).",
        "Los precios de hardware y los costos fijos de operación son los hipotéticos del §14.",
        "Los impuestos de importación en Brasil suman un 50 % sobre el equipo importado.",
        "La red agrega 0,2 s de ida y vuelta, y procesar la entrada en una réplica cargada toma 0,5 s.",
    ])


def s3(doc, wb):
    doc.add_heading("3. Demanda y objetivos de servicio", 1)
    para(doc, "La cadena es corta. 200.000 usuarios por 1,5 peticiones dan 300.000 peticiones al día. Repartidas en "
              "86.400 segundos son 3,47 por segundo, y en el pico, seis veces más: **20,83 peticiones por segundo**.")
    para(doc, "Los tokens se cuentan por separado porque no cuestan lo mismo. En el pico entran unos 83.333 tokens por "
              "segundo y salen unos 10.417. La relación es de 8 a 1: este servicio lee mucho más de lo que escribe, y "
              "eso pesa en la fase de prefill.")
    ws = wb["Demanda"]
    rows = []
    for r in ws.iter_rows(min_row=5, max_row=20, values_only=True):
        if r[0] and r[2] is not None and not isinstance(r[2], str):
            dec = 2 if r[0] in ("RPS_PROM", "RPS_PICO") else (1 if r[0] == "RATIO_ENT_SAL" else 0)
            rows.append((r[1], fmt_num(r[2], dec), fmt_num(r[3], dec), r[4]))
    table(doc, ["Concepto", "Base", "Demanda ×2", "Unidad"], rows, widths=[7.5, 3, 3, 3])
    para(doc, "Para la concurrencia se usa la ley de Little con promedios compatibles: si una petición pasa en "
              "promedio 20 segundos en el sistema, en el pico hay unas 417 a la vez. Esos 20 segundos son un "
              "supuesto, y es coherente con el punto de operación de la sección 5. No hay que confundir este número "
              "con usuarios conectados ni con conversaciones abiertas: alguien puede tener el chat abierto sin "
              "ninguna petición en curso.")
    para(doc, "Las tablas de sensibilidad del Excel muestran algo incómodo. El §14.4 pide probar una entrada "
              "promedio del doble, es decir, 8.000 tokens. Pero 8.000 más 500 de salida suman 8.500, y eso **no cabe** "
              "en la ventana de 8.192. Ese escenario no es solo más caro: obliga a ampliar la ventana o a recortar "
              "contexto.")
    para(doc, "Todos los tiempos se miden desde el cliente: red, espera en la fila y procesamiento. Los rechazos, "
              "cancelaciones y fallas se reportan, no se descartan. Y vale recordar que disponibilidad y latencia son "
              "cosas distintas: un servidor puede estar encendido y aun así responder tarde.")


def s4(doc, wb):
    doc.add_heading("4. Modelo y memoria", 1)
    para(doc, "Se compararon dos versiones del mismo modelo: Llama 3.1 70B Instruct en BF16, que es la referencia del "
              "enunciado, y su versión cuantizada a FP8 publicada por RedHat. Ser de la misma familia hace que la "
              "comparación sea limpia: lo único que cambia es cómo se guardan los números.")
    table(doc, ["Criterio", "Llama 3.1 70B BF16", "Llama 3.1 70B FP8"], [
        ("Identificador", "meta-llama/Llama-3.1-70B-Instruct", "RedHatAI/Meta-Llama-3.1-70B-Instruct-FP8"),
        ("Parámetros", "70.550 millones, densos", "Los mismos"),
        ("Licencia", "Llama 3.1 Community License", "La misma"),
        ("Español", "Soportado; MMLU multilingüe en español 80,05", "Sin evaluar: la ficha deja fuera otros idiomas"),
        ("Calidad publicada", "MGSM 86,9; IFEval 87,5", "99,88 % del original, solo en inglés"),
        ("Pesos / KV cache", "BF16 / BF16", "FP8 en capas lineales / KV en 16 bits"),
        ("Hardware por réplica", "4 × H100 80 GB", "2 o 4 × H100 80 GB"),
    ], widths=[3.6, 6.2, 6.6])
    para(doc, "Los parámetros se calcularon desde el config.json del modelo: 80 capas, 64 cabezas de atención, 8 "
              "cabezas KV y dimensión oculta de 8.192. El total da 70.550 millones, lo mismo que muestra Hugging "
              "Face. Con eso:")
    bullets(doc, [
        "Los pesos ocupan **141,1 GB en BF16** y 72,7 GB en FP8 (embeddings y lm_head siguen en 16 bits).",
        "El KV cache ocupa 0,33 MB por token: **1,47 GB** por petición promedio de 4.500 tokens y 2,68 GB con la "
        "ventana completa.",
    ])
    ws = wb["Modelo y memoria"]
    want = {"MEM_PESOS": 1, "KV_DISP": 1, "PET_MAX_PROM": 0, "PET_MAX_PEOR": 0, "EJ_TOTAL_PEOR": 1}
    rows = []
    for r in ws.iter_rows(min_row=1, values_only=True):
        if r[0] in want:
            d = want[r[0]]
            rows.append((r[1], fmt_num(r[2], d), fmt_num(r[3], d), fmt_num(r[4], d), r[5]))
    table(doc, ["Por réplica (80 % de la memoria utilizable)", "BF16 · 4 GPUs", "FP8 · 4 GPUs", "FP8 · 2 GPUs",
                "Unidad"], rows, widths=[6.5, 2.5, 2.5, 2.5, 2.4])
    para(doc, "Que el modelo quepa no dice nada de si alcanza. En BF16 caben 50 peticiones promedio por réplica, pero "
              "solo 27 si todas usan la ventana completa. El ejemplo del enunciado, con 40 peticiones de 8.192 "
              "tokens, se pasa de la memoria utilizable. Por eso el control de admisión limita **tokens reservados**, "
              "no solo número de peticiones.")
    para(doc, "La versión FP8 ahorraría GPUs, pero su ficha deja fuera de alcance los idiomas distintos del inglés. "
              "No hay forma de afirmar que mantenga el 90 % de calidad en español. El diseño se hace en BF16 y FP8 "
              "queda como ahorro posible a futuro, si se mide su velocidad y su calidad en español.")
    doc.add_heading("Plan de evaluación de calidad", 2)
    para(doc, "No se entregó el conjunto de 100 tareas y no hubo dónde correr el modelo, así que no hay puntajes. "
              "Lo que sigue es el plan:")
    bullets(doc, [
        "Congelar la revisión exacta del modelo de referencia y del candidato antes de evaluar.",
        "Armar 100 tareas en español: 40 de comprensión y respuesta sobre un texto dado, 20 de resumen, 20 de seguir "
        "instrucciones, 10 de contexto largo (cerca de 8.000 tokens) y 10 de rechazo o de instrucciones maliciosas "
        "escondidas en documentos.",
        "Calificar cada tarea con una rúbrica común de 0 a 2 en corrección, seguimiento de instrucciones y apego "
        "al texto. La calificación la hacen personas; un LLM juez puede apoyar, pero no decidir solo.",
        "Aceptar el candidato si saca al menos el 90 % del puntaje de la referencia y no tiene ninguna falla crítica "
        "de privacidad o de mezcla de información entre usuarios.",
    ])


def s5(doc, wb):
    doc.add_heading("5. Evidencia de capacidad", 1)
    para(doc, "No se midió nada. El dato de capacidad es el sintético del §8.4: una réplica de 4 H100 atiende **2 "
              "peticiones por segundo** de 4.000 tokens de entrada y 500 de salida, cumpliendo todos los SLOs a la "
              "vez. Eso equivale a 8.000 tokens de entrada y 1.000 de salida por segundo en el mismo punto de "
              "operación, no a dos máximos separados.")
    para(doc, "Lo que sí se puede hacer es revisar que el número no rompa la física de la H100:")
    ws = wb["Rendimiento"]
    vals = {r[0]: r[2] for r in ws.iter_rows(min_row=5, values_only=True) if r[0]}
    table(doc, ["Prueba", "Resultado", "Lectura"], [
        ("Cómputo que pide el punto de operación", f"{fmt_num(vals['FLOPS_USADOS'])} de {fmt_num(vals['FLOPS_PICO'])} "
         f"TFLOPS ({fmt_pct(vals['USO_COMPUTO'])})", "Posible, ni imposible ni sospechoso"),
        ("Peticiones generando a la vez (40 ms por token)", f"{fmt_num(vals['SEC_DECODE'])} de 50 que caben",
         "Cabe en la memoria"),
        ("Tiempo mínimo de un paso de generación", f"{fmt_num(vals['T_PASO_MIN'], 1)} ms", "Por debajo de 40 ms"),
    ], widths=[6.5, 5, 4.9])
    para(doc, "Las dos pruebas dicen que 2 peticiones por segundo es físicamente posible. No dicen que se alcance. "
              "El cómputo denso de la H100 se obtuvo dividiendo entre dos la cifra que publica NVIDIA, que es con "
              "dispersión; eso también es un supuesto.")
    para(doc, "Como referencia externa, MLPerf Inference v5.0 tiene una prueba interactiva de Llama 2 70B con 40 ms "
              "por token en p99. Muestra que la meta es realista en la industria, pero usa otro modelo y textos mucho "
              "más cortos, así que sus cifras no se trasladan a este caso.")
    para(doc, "**Cómo se validaría:** fijar modelo, revisión, motor (por ejemplo vLLM) y su versión; usar 4 H100 con "
              "paralelismo tensorial de 4 y el perfil 4.000/500; calentar; subir la tasa de llegada en lazo abierto "
              "hasta que se rompa algún SLO; repetir tres veces y reportar TTFT, tiempo por token y total en p95, "
              "peticiones completadas, rechazadas y fallidas, y consumo eléctrico. Después, repetir con la mezcla de "
              "peticiones cortas y largas, con arranque en frío y con la pérdida de un nodo.")


def s6(doc, wb):
    doc.add_heading("6. Réplicas y nodos", 1)
    para(doc, "Una réplica es una copia completa del modelo: 4 GPUs. Un nodo es un servidor físico DGX H100 con 8 GPUs, "
              "así que en cada nodo caben 2 réplicas.")
    para(doc, "Cada réplica se carga al 70 % de su capacidad, 1,4 peticiones por segundo, para absorber la "
              "variación dentro de la hora pico. Con eso hacen falta 15 réplicas, es decir 8 nodos. Se suma un nodo "
              "para poder perder uno: **9 nodos, 72 GPUs y 18 réplicas**. En operación normal el pico usa el 57,9 % "
              "de la capacidad instalada, y el 65,1 % si falta un nodo.")
    ws = wb["Inventario"]
    rows = []
    in_fail = False
    for r in ws.iter_rows(min_row=1, values_only=True):
        if r[1] == "Falla":
            in_fail = True
            continue
        if in_fail and r[1] and r[2] is not None and not str(r[1]).startswith('"'):
            rows.append((r[1], fmt_num(r[2]), fmt_num(r[3], 1), r[4], fmt_num(r[5], 1), r[6]))
        if in_fail and r[1] is None and r[0] is None and rows:
            break
    table(doc, ["Falla", "Réplicas", "Capacidad con holgura (pet/s)", "¿Cubre?", "Al 100 % (pet/s)", "¿Cubre?"],
          rows[:4], widths=[4.6, 1.8, 3, 1.6, 2.8, 1.6])
    para(doc, "El diseño cubre la pérdida de un nodo, que es lo que pide el enunciado. No cubre un rack completo: si "
              "caen sus 3 nodos, quedan 12 réplicas y 16,8 peticiones por segundo con margen, menos que el pico. Al "
              "100 % alcanzaría, pero sin ningún margen. Prometer tolerancia a un rack exige 11 nodos.")
    para(doc, "Que la capacidad alcance no prueba que se cumpla la latencia durante la falla. Falta contar la "
              "detección, las peticiones que estaban en curso y los reintentos; eso se ve en la sección 11.")


def s7(doc, wb):
    doc.add_heading("7. Arquitectura y despliegue", 1)
    para(doc, "Los componentes se describen por lo que hacen, con un producto de ejemplo entre paréntesis.")
    figure(doc, DIAG / "arquitectura_logica.png", "Figura 1. Camino lógico de una petición.")
    doc.add_heading("Una petición normal", 2)
    para(doc, "El usuario se conecta por HTTPS al balanceador, que termina TLS. El gateway revisa quién es, si le "
              "queda cuota y que la entrada no pase de 8.192 tokens. El router busca una réplica que tenga espacio "
              "en su presupuesto de tokens y le pasa la petición. El motor de inferencia (por ejemplo vLLM) la suma a "
              "su lote continuo, procesa la entrada y devuelve la respuesta palabra por palabra. Las métricas "
              "registran tiempos, tokens y errores, nunca el texto de la conversación.")
    doc.add_heading("Cuando hay sobrecarga", 2)
    para(doc, "Si ninguna réplica tiene espacio, la petición espera en una fila corta. La espera máxima es de "
              "**1 segundo**, y no es un número al azar: de los 2 segundos del TTFT hay que descontar 0,2 de red y "
              "0,5 para procesar la entrada, lo que deja 1,3. Pasado ese segundo, el servicio responde \"intenta en "
              "un momento\" en lugar de dejar al usuario esperando algo que igual llegaría tarde.")
    bullets(doc, [
        "La fila tiene como máximo unas 21 peticiones en todo el servicio.",
        "Cada usuario puede tener 3 peticiones a la vez y 20 por minuto: nadie puede ocupar más del 1,6 % del pico.",
        "Si el usuario cierra el chat, su generación se cancela y la memoria se libera.",
        "Hay un solo reintento automático, y solo si todavía no se le envió ninguna palabra.",
        "Al admitir una petición se reserva memoria para su entrada y su salida máxima. Así caben unas 45 por "
        "réplica, más que las 26 que habría si falla un nodo.",
    ])
    doc.add_heading("Cuando falla un nodo", 2)
    para(doc, "El router deja de recibir los latidos de sus dos réplicas y las saca del reparto. Las respuestas que "
              "iban a medias se pierden: el usuario ve un error y vuelve a preguntar. No se promete continuar "
              "respuestas parciales, porque eso exigiría guardar y reconstruir el estado, y no está diseñado. Quedan "
              "16 réplicas con 22,4 peticiones por segundo de capacidad, por encima del pico.")
    figure(doc, DIAG / "despliegue.png", "Figura 2. Despliegue físico: 3 racks de cómputo y uno de servicios.",
           width=13.5)
    doc.add_heading("Paralelismo, lotes y caché", 2)
    para(doc, "Hay paralelismo de datos (18 réplicas independientes) y paralelismo tensorial de 4 GPUs dentro de cada "
              "servidor, unidas por NVLink. No se usa paralelismo de tubería ni de expertos: el modelo es denso y "
              "cabe en 4 GPUs. Los lotes continuos aprovechan la GPU sin hacer esperar a nadie. El caché de prefijos "
              "solo se comparte dentro del mismo cliente, por ejemplo para sus instrucciones de sistema, nunca entre "
              "usuarios, y no se supone que acierte siempre. Tampoco se guardan respuestas completas para "
              "reutilizarlas.")
    doc.add_heading("Escalado", 2)
    para(doc, "El límite es físico: 18 réplicas, siempre encendidas y cargadas, porque recargar 141 GB de pesos no es "
              "instantáneo. Escalar contenedores no crea GPUs, ni electricidad, ni enfriamiento. Las señales para "
              "comprar más nodos son un uso sostenido por encima del 70 %, rechazos por fila llena o un TTFT p95 que "
              "se acerca a 2 segundos.")
    doc.add_heading("Operación y seguridad", 2)
    para(doc, "Los despliegues empiezan por una réplica de prueba y, si va bien, siguen con el resto; si no, se vuelve "
              "a la versión anterior. Los modelos se registran con versión y huella. Los secretos viven en un gestor "
              "y no en el código. Todo va cifrado en tránsito y en reposo. Los clientes quedan aislados en cuotas, "
              "caché y registros, y los registros no guardan el contenido de las conversaciones. Las instrucciones "
              "escondidas en un documento se tratan como texto del usuario, no como órdenes, y el modelo no ejecuta "
              "herramientas.")
    doc.add_heading("Red y almacenamiento", 2)
    para(doc, "Hay tres tráficos muy distintos, y conviene no mezclarlos:")
    figure(doc, DIAG / "mapa_trafico.png", "Figura 3. Mapa de tráfico.")
    bullets(doc, [
        "El de los usuarios es pequeño: unos 15 Mb/s en el pico. Pesan más las respuestas que las preguntas, porque "
        "cada palabra viaja en su propio mensaje.",
        "Entre las 4 GPUs de una réplica circulan unos 24 GB/s. Es el 2,6 % de NVLink, pero 7,5 veces lo que cabe "
        "en un enlace de 25 Gb/s. Esa es la razón, en números, para no repartir una réplica entre servidores.",
        "Cargar el modelo pesa 141 GB por réplica: 45 segundos ideales por un enlace de 25 Gb/s, y unos 4 minutos "
        "para arrancar las 18 a la vez desde el almacenamiento central. Por eso cada servidor guarda una copia de "
        "los pesos en su propio disco.",
    ])
    para(doc, "El almacenamiento central necesita unos 640 GB útiles (tres versiones del modelo, imágenes, datos de la "
              "aplicación y 30 días de registros), que con tres copias y holgura son 3,9 TB. Fuera del sitio se "
              "guardan unos 0,6 TB; restaurar una versión del modelo desde ahí tarda unos 19 minutos.")


def s8(doc, wb):
    doc.add_heading("8. Ubicación e infraestructura", 1)
    para(doc, "Se eligió **colocation**: el proveedor pone edificio, electricidad, enfriamiento, UPS, seguridad y "
              "conectividad, y nosotros compramos y operamos los servidores. Se compararon tres mercados de la región.")
    ws = wb["Ubicación"]
    keep = {"Precio de la electricidad": (3, "USD/kWh"), "Emisiones de la red eléctrica": (3, "kg CO₂e/kWh"),
            "Renta de espacio (punto medio del rango)": (0, "USD/kW/mes"),
            "Capacidad disponible en el mercado": (1, "MW"), "Riesgo sísmico": (0, ""),
            "Peor tiempo de red para un usuario": (0, "ms"), "¿Cabe en la red supuesta para el TTFT?": (0, ""),
            "Costo de electricidad": (0, "USD/año"), "Renta de espacio": (0, "USD/año"),
            "Emisiones de operación": (0, "t CO₂e/año")}
    rows = []
    for r in ws.iter_rows(min_row=5, values_only=True):
        if r[1] in keep:
            d, unit = keep[r[1]]
            rows.append((r[1], fmt_num(r[2], d), fmt_num(r[3], d), fmt_num(r[4], d), unit))
    table(doc, ["", "Querétaro", "São Paulo", "Santiago", "Unidad"], rows, widths=[6, 2.6, 2.6, 2.6, 2.6])
    para(doc, "São Paulo sale mejor en casi todo, pero el dato decisivo es la red. Sumando 30 ms de la conexión del "
              "propio usuario, el peor caso desde São Paulo es Ciudad de México con 188 ms, dentro de los 0,2 "
              "segundos del presupuesto del TTFT. Desde Querétaro, Bogotá queda a 223 ms porque el tráfico pasa por "
              "Miami, y desde Santiago, Ciudad de México queda a 208 ms. Ninguna de las dos cumple.")
    para(doc, "Hay dos advertencias. La renta de CBRE es para contratos de 250 a 500 kW, y el nuestro ronda los 104, "
              "así que el precio por kW puede salir más alto. Y CBRE menciona problemas de confiabilidad eléctrica en "
              "algunas zonas de São Paulo: el proveedor elegido debe tener UPS y generadores propios.")
    doc.add_heading("Diseño físico", 2)
    para(doc, "Son 4 racks: 3 de cómputo con 3 DGX cada uno, y uno de red, almacenamiento y servicios. Cada rack de "
              "cómputo ocupa 24 de sus 42 unidades, pesa unos 391 kg solo en servidores, mueve 3.315 CFM de aire y "
              "consume **30,6 kW**. El límite no es el espacio: es la electricidad y el calor. Los DGX H100 se "
              "enfrían por aire, así que con esa densidad hace falta pasillo caliente contenido o intercambiadores "
              "en la puerta trasera. Si el proveedor solo soporta unos 25 kW por rack, se pasa a 2 servidores por "
              "rack y 6 racks en total.")
    doc.add_heading("Opciones descartadas", 2)
    para(doc, "Descartar con números también es una conclusión:")
    bullets(doc, [
        "**Marte.** La señal tarda entre 3 y 22 minutos solo de ida. Un TTFT de 2 segundos es imposible para "
        "usuarios en la Tierra.",
        "**Órbita.** Solo para radiar 103,8 kW de calor harían falta al menos 251 m² de radiadores, en el caso "
        "ideal. A eso se suman radiación, eclipses, enlaces limitados y lanzamientos que no caben en USD 8 millones.",
        "**Submarino.** El mar recibe el calor, pero cada reparación de una GPU obliga a sacar la cápsula. Project "
        "Natick fue un experimento, no una garantía.",
    ])


def s9(doc, wb):
    doc.add_heading("9. Electricidad y enfriamiento", 1)
    para(doc, "Una aclaración antes de los números: kW es potencia, lo que se consume en un instante; kWh es energía, "
              "potencia por tiempo.")
    ws = wb["Energía"]
    keep = {"E_P_TI": 1, "E_P_SITIO": 1, "E_HOLG_ACOM": "%", "E_FRIO_DISP": 0, "E_FRIO_USO": "%",
            "E_CAUDAL_LMIN": 0}
    rows = []
    for r in ws.iter_rows(min_row=5, max_row=22, values_only=True):
        if r[0] in keep:
            d = keep[r[0]]
            f = (lambda v: fmt_pct(v)) if d == "%" else (lambda v, d=d: fmt_num(v, d))
            rows.append((r[1], f(r[2]), f(r[3]), f(r[4]), r[5]))
    table(doc, ["", "Normal", "Máxima", "Degradada", "Unidad"], rows, widths=[6.4, 2.4, 2.4, 2.6, 2.6])
    para(doc, "La potencia de TI máxima es 9 × 10,2 kW más 12 kW de red y servicios: 103,8 kW. La potencia del "
              "servidor ya incluye sus GPUs, así que no se suman otra vez. Con 35 kW de cargas que no son de TI, el "
              "sitio pide 138,8 kW, y los 160 kW contratados dejan un 13 % de margen.")
    para(doc, "En promedio, la TI usa 70,2 kW. Con un PUE de 1,30 eso da **799 MWh al año**, unos USD 117.000 en "
              "São Paulo. El enfriamiento tiene tres módulos de 60 kW: si uno falla quedan 120 kW, más que el calor "
              "máximo. El circuito de agua necesita unos 149 litros por minuto con 10 °C de subida.")
    para(doc, "La UPS protege 120 kW durante 10 minutos, mientras arrancan los generadores. Son 20 kWh útiles, unos "
              "35 kWh nominales de baterías. Cubre toda la TI, pero no el sitio completo; hay que confirmar con el "
              "proveedor qué parte del enfriamiento queda protegida.")
    para(doc, "Hay dos hallazgos que conviene mostrar:")
    bullets(doc, [
        "NVIDIA declara 38.557 BTU/h de calor por DGX, que son 11,3 kW, no 10,2. Con ese valor la TI llegaría a "
        "unos 114 kW, y el enfriamiento con un módulo fuera igual alcanza. El diseño aguanta la duda, pero hay que "
        "aclararla.",
        "Con demanda doble, 16 nodos, el sitio pediría unos 234 kW, mucho más que los 160 contratados. Crecer no es "
        "solo comprar servidores: hay que negociar potencia y enfriamiento con meses de anticipación.",
    ])
    para(doc, "En lo ambiental, la red de Brasil deja unas **88 toneladas de CO₂e al año**; en Querétaro serían 379. "
              "El agua directa se supuso en cero por ser un circuito cerrado, pero si el proveedor usa torres "
              "evaporativas serían entre 615 y 1.230 m³ al año, y ese dato hay que pedirlo. Cada unos cinco años "
              "habrá que reciclar 1,2 toneladas de servidores. Un PUE bajo no dice nada sobre emisiones ni sobre "
              "agua.")


def s10(doc, wb):
    doc.add_heading("10. Presupuesto", 1)
    para(doc, "Todo está en dólares, con precios a noviembre de 2026. Salvo donde se indica, son los valores "
              "hipotéticos del enunciado, no cotizaciones. NVIDIA no publica un precio oficial del DGX H100; los "
              "USD 350.000 usados quedan dentro del rango público de 300.000 a 400.000.")
    ws = wb["Presupuesto"]
    capex, opex, ind = [], [], []
    section = None
    for r in ws.iter_rows(min_row=4, values_only=True):
        if r[0] and str(r[0]).startswith("CAPEX ("):
            section = "c"
            continue
        if r[0] and str(r[0]).startswith("OPEX ("):
            section = "o"
            continue
        if r[0] and str(r[0]).startswith("TCO e"):
            section = "t"
            continue
        if r[0] and str(r[0]).startswith("Límites"):
            section = None
        if not r[0] or r[4] is None or section is None:
            continue
        if r[0] in ("IMPUESTO_IMPORT", "CONTINGENCIA", "CX_IMPORTADO", "OX_FIJO", "FRAC_BUENAS"):
            continue
        if section == "c":
            capex.append((r[1], fmt_num(r[2]) if r[2] is not None else "", fmt_num(r[4])))
        elif section == "o":
            opex.append((r[1], fmt_num(r[4])))
        else:
            dec = 4 if "COSTO_PET" in r[0] else (2 if "MTOK" in r[0] else 0)
            ind.append((r[1], fmt_num(r[4], dec)))
    table(doc, ["CAPEX", "Cantidad", "USD"], capex, widths=[10, 2.4, 4])
    para(doc, "Al ser colocation, se quitaron del ejemplo la obra eléctrica, la UPS, el enfriamiento y el edificio: "
              "USD 1,2 millones que el proveedor ya cobra en la renta. A cambio, se agregaron USD 100.000 para "
              "adecuar el espacio. Los impuestos de importación de Brasil se estimaron, con prudencia, en un 50 % "
              "sobre el equipo importado; son la segunda partida más grande.")
    table(doc, ["OPEX anual", "USD"], opex, widths=[12.4, 4])
    para(doc, "El personal se paga para tener guardia las 24 horas con turnos y soporte remoto; unas pocas personas "
              "no equivalen a presencia continua en el sitio. El soporte que ya trae el hardware no se cobra dos "
              "veces.")
    table(doc, ["Indicador", "Valor (USD)"], ind, widths=[12.4, 4])
    para(doc, "El TCO usa demanda base constante, sin descuento financiero ni valor residual. El costo por millón de "
              "tokens cuenta entrada y salida en el denominador; no se puede comparar con una tarifa de API que cobra "
              "solo una parte. El costo por petición correcta y a tiempo es una cota, porque la calidad de las "
              "respuestas no está medida.")
    para(doc, "Los tres límites se cumplen: CAPEX al 95 %, OPEX al 91 % y TCO al 93 %. Es poco margen, y los "
              "escenarios muestran por qué importa:")
    ws = wb["Escenarios"]
    hdr = [c for c in next(ws.iter_rows(min_row=4, max_row=4, values_only=True))]
    vals = {r[0]: r for r in ws.iter_rows(min_row=5, values_only=True) if r[0]}
    rows = []
    for i in range(2, len(hdr)):
        if hdr[i] is None:
            continue
        rows.append((hdr[i], fmt_num(vals["S_NODOS"][i]), fmt_num(vals["S_CAPEX"][i] / 1e6, 2),
                     fmt_num(vals["S_OPEX"][i] / 1e6, 2), fmt_num(vals["S_TCO"][i] / 1e6, 2),
                     fmt_num(vals["S_COSTO_PET"][i], 4), vals["S_OK_TCO"][i]))
    table(doc, ["Escenario", "Nodos", "CAPEX (M)", "OPEX (M)", "TCO (M)", "USD/petición", "¿TCO cabe?"], rows,
          widths=[4.4, 1.4, 2, 2, 2, 2.4, 2])
    para(doc, "La variable que más mueve el costo es la demanda, aunque con el doble de demanda cada petición sale "
              "más barata. Las que más mueven el **riesgo** son el rendimiento real por réplica y el precio del "
              "hardware: con un 30 % de desvío en cualquiera de las dos, el CAPEX se pasa del límite. La energía casi "
              "no cambia nada. Si el proyecto califica al régimen ReData, los impuestos bajan y el TCO cae a unos "
              "USD 10,2 millones.")


def s11(doc, wb):
    doc.add_heading("11. Resiliencia y crecimiento", 1)
    para(doc, "Un 99,9 % mensual deja **43,2 minutos** de caída en un mes de 30 días. Una sola intervención mal hecha "
              "se los puede gastar. También hay un presupuesto de errores: el 0,1 % de 9 millones de peticiones, "
              "9.000 al mes.")
    ws = wb["Resiliencia"]
    rows, on = [], False
    for r in ws.iter_rows(min_row=1, values_only=True):
        if r[1] == "Falla":
            on = True
            continue
        if on:
            if not r[1]:
                break
            rows.append((r[1], r[2], r[3], fmt_num(r[4], 1), r[5], r[6], r[7], r[8]))
    landscape_section(doc)
    para(doc, "Matriz de fallas", size=12.5, color=ACCENT).runs[0].bold = True
    table(doc, ["Falla", "Detección", "Impacto", "Capacidad (pet/s)", "¿Cubre?", "Acción", "Recuperación",
                "Riesgo que queda"], rows, widths=[2.6, 3.4, 3.4, 1.8, 1.4, 3.8, 3, 4.4], size=8)
    portrait_section(doc)
    para(doc, "**Ejercicio: se cae un nodo en plena hora pico.** Durante los 30 segundos que tarda en detectarse, se "
              "pierden unas 46 peticiones que estaban en curso y unas 69 que el router alcanza a mandarle. Son 116 "
              "errores, el 1,3 % del presupuesto del mes. Después quedan 22,4 peticiones por segundo de capacidad, "
              "por encima del pico. Si cae un segundo nodo antes de reparar el primero, unas 48 horas, quedan 19,6 "
              "con margen: se atiende, pero trabajando al límite.")
    para(doc, "Lo que no se cubre se dice. Un rack completo deja el servicio por debajo del pico con margen, y "
              "cubrirlo cuesta unos USD 1,5 millones más. Perder el sitio entero deja el servicio caído durante "
              "semanas, porque habría que comprar e importar equipo; no hay segundo sitio en el presupuesto. Los dos "
              "riesgos se aceptan y quedan documentados.")
    table(doc, ["Qué", "RTO", "RPO", "Cómo se prueba"], [
        ("Falla de GPU, nodo o enlace", "Segundos", "Nada (no se guardan conversaciones)",
         "Apagar un nodo a propósito cada trimestre"),
        ("Versión defectuosa", "~15 min", "Nada", "Cada despliegue empieza con una réplica de prueba"),
        ("Pesos del modelo", "~19 min desde fuera", "Nada (copias versionadas)",
         "Restaurar una versión desde fuera cada semestre"),
        ("Datos de la aplicación", "~1 h", "1 h", "Restaurar en un entorno de prueba cada mes"),
        ("Pérdida del sitio", "Semanas", "1 h de datos", "Ejercicio de mesa anual"),
    ], widths=[4.2, 2.8, 4, 5.4])
    para(doc, "Dos componentes no son alta disponibilidad si comparten alimentación, red o configuración. Y tener una "
              "copia no prueba que se pueda restaurar a tiempo: hay que ensayarlo.")
    para(doc, "**Crecer.** Con el doble de demanda hacen falta 16 nodos y unos 234 kW, más de lo contratado. El TCO "
              "sube a USD 16,8 millones, que se presentan aparte del caso base. La compra y la negociación de "
              "potencia toman meses, así que el aviso tiene que llegar antes de que la demanda lo exija.")


def s12(doc):
    doc.add_heading("12. Recomendación y condiciones que podrían cambiarla", 1)
    para(doc, "Se recomienda instalar 9 DGX H100 en colocation en São Paulo con Llama 3.1 70B BF16, bajo cuatro "
              "condiciones:")
    bullets(doc, [
        "**Medir antes de comprar.** Hay que confirmar con una prueba de carga que una réplica de 4 H100 atiende "
        "cerca de 2 peticiones por segundo dentro de los SLOs. Si rinde un 30 % menos, el diseño necesita 12 nodos "
        "y se sale del presupuesto.",
        "**Cotizar.** El precio del DGX, la renta para un contrato de unos 104 kW y la carga real de impuestos en "
        "Brasil (con o sin ReData) pueden mover el CAPEX más que cualquier otra cosa.",
        "**Confirmar con el proveedor** que soporta ~31 kW por rack con aire, qué enfriamiento queda en UPS y cuánta "
        "agua consume.",
        "**Evaluar la calidad** con las 100 tareas en español antes de cambiar a FP8 o a cualquier otro modelo.",
    ])
    para(doc, "Algunas cosas cambiarían la recomendación. Si el rendimiento medido es mucho mejor, sobran nodos y "
              "conviene bajar a 8 o menos. Si FP8 mantiene la calidad en español, se podría servir con menos GPUs. "
              "Si la demanda se concentra en México y no en toda la región, Querétaro merecería otra mirada pese a "
              "su costo. Y si el negocio no acepta semanas de caída ante la pérdida del sitio, habrá que presupuestar "
              "un segundo sitio, que hoy no cabe.")


def annex(doc, wb):
    landscape_section(doc)
    doc.add_heading("Anexo A. Registro de decisiones y datos", 1)
    ws = wb["Registro"]
    rows = [(r[0], r[1], r[2], r[3], r[4], r[6]) for r in ws.iter_rows(min_row=5, values_only=True) if r[0]]
    table(doc, ["#", "Decisión o dato", "Valor y unidad", "Fuente o cálculo", "Incertidumbre", "Clasificación"],
          rows, widths=[0.8, 5, 5.4, 6, 5.4, 2.4], size=8)
    para(doc, "Responsable de verificar todos los datos: Juliana Briceño.", italic=True, size=9, color=MUTED)

    doc.add_heading("Anexo B. Uso de IA", 1)
    ws = wb["Uso de IA"]
    rows, errs = [], []
    for row in ws.iter_rows(min_row=5):
        r = [c.value for c in row]
        if isinstance(r[0], int):
            rows.append((r[0], r[2], r[3], r[4], r[5]))
            fill = row[0].fill.fgColor.rgb if row[0].fill and row[0].fill.fgColor else ""
            if str(fill).upper().endswith("FBE9DD"):
                errs.append(str(r[0]))
    para(doc, f"Herramienta en todas las filas: Claude Code (Claude Opus 5.5), con búsqueda y lectura web. De las "
              f"{len(rows)} consultas, {len(errs)} tuvieron errores, omisiones o simplificaciones (filas "
              f"{', '.join(errs[:-1])} y {errs[-1]}).", size=10)
    table(doc, ["#", "Pregunta relevante", "Resultado usado", "Verificación humana", "Corrección realizada"],
          rows, widths=[0.8, 5, 5.6, 7, 6.6], size=8)
    portrait_section(doc)
    doc.add_heading("Anexo C. Fuentes", 1)
    bullets(doc, [
        "NVIDIA. DGX H100/H200 User Guide; página de producto H100 y H200. Consultadas el 13-nov-2026.",
        "Meta. Llama 3.1 Model Card; config.json de Meta-Llama-3.1-70B-Instruct (espejo NousResearch).",
        "RedHatAI. Meta-Llama-3.1-70B-Instruct-FP8, ficha del modelo.",
        "MLCommons. MLPerf Inference v5.0 Advances Language Model Capabilities for GenAI.",
        "CBRE. Global Data Center Trends 2026 (datos de Q1 2026).",
        "GlobalPetrolPrices. Precios de electricidad para empresas en México, Brasil y Chile (marzo 2026).",
        "Ember 2025, vía Carbon Recycling y GreenCalculus. Factores de emisión de la red eléctrica.",
        "WonderNetwork. Pings entre Ciudad de México, São Paulo y Santiago (13-nov-2026).",
        "PwC Brasil, Brazil Bottom Up, BNamericas y Global Data Center Hub. Impuestos y régimen ReData.",
        "NASA. Speaking Percy's Language; Small Spacecraft Technology: Thermal Control.",
        "Microsoft Research. Project Natick. The Green Grid. PUE and Efficiency Metrics.",
    ])
    para(doc, "El modelo de cálculo completo está en modelo_calculo.xlsx; los números de este documento salen de él.",
         italic=True, size=9, color=MUTED)


def main():
    wb = recalculated_workbook()
    doc = Document()
    setup(doc)
    cover(doc)
    s1(doc)
    s2(doc)
    s3(doc, wb)
    s4(doc, wb)
    s5(doc, wb)
    s6(doc, wb)
    s7(doc, wb)
    s8(doc, wb)
    s9(doc, wb)
    s10(doc, wb)
    s11(doc, wb)
    s12(doc)
    annex(doc, wb)
    OUT.parent.mkdir(parents=True, exist_ok=True)
    doc.save(OUT)
    print(f"Saved {OUT.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
