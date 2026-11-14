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
    sec.left_margin = sec.right_margin = Cm(2.2)
    sec.top_margin = sec.bottom_margin = Cm(2.0)


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
    gap = doc.add_paragraph()
    gap.paragraph_format.space_after = Pt(0)
    gap.paragraph_format.line_spacing = 0.6
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
# Body (sections 1-12) must fit in 12 pages; long tables and plans live in the appendices.
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
    para(doc, "São Paulo ganó a Querétaro y a Santiago por casi todo lo que se midió. Pero el argumento que más pesa es "
              "otro: es la única de las tres desde donde el usuario más lejano queda dentro del presupuesto de red "
              "que exige el tiempo de primera palabra.")
    para(doc, "El riesgo está en un lugar concreto. La capacidad por réplica es un dato **sintético** del enunciado, no "
              "una medición. Si una réplica rinde un 30 % menos, o si el hardware cuesta un 30 % más, el proyecto se "
              "sale del presupuesto. Por eso la recomendación pide medir y cotizar antes de comprar.")


def s2(doc):
    doc.add_heading("2. Requisitos y supuestos", 1)
    para(doc, "Es un asistente de texto en español para Latinoamérica: responde preguntas, resume documentos y ayuda a "
              "escribir. Solo hay inferencia; no se entrena nada y no se incluye RAG, que queda como extensión.")
    table(doc, ["Requisito", "Valor"], [
        ("Demanda", "200.000 usuarios/día × 1,5 peticiones; factor de pico 6"),
        ("Perfil controlado", "4.000 tokens de entrada y 500 de salida; ventana 8.192; salida máx. 1.000"),
        ("Latencia (p95)", "TTFT ≤ 2 s; ≤ 40 ms por token; ≤ 30 s para 500 tokens"),
        ("Disponibilidad y errores", "99,9 % mensual; ≤ 0,1 % de errores; pico cubierto tras perder un nodo"),
        ("Calidad", "≥ 90 % de Llama 3.1 70B Instruct BF16, sin fallas críticas de privacidad"),
        ("Presupuesto", "CAPEX ≤ USD 8 M; OPEX ≤ 1,5 M/año; TCO a 3 años ≤ 12,5 M"),
    ], widths=[4, 12.6])
    para(doc, "Cada dato del Excel está clasificado como documentado, medido, calculado o supuesto, con su fuente "
              "(anexo A). No hay datos medidos: no hubo acceso a GPUs. Los supuestos que más mueven el resultado son "
              "la capacidad sintética de 2 peticiones/s por réplica, el consumo promedio de 6,8 kW por servidor con "
              "PUE de 1,30, los precios hipotéticos del §14, un 50 % de impuestos de importación en Brasil y 0,2 s de "
              "red más 0,5 s para procesar la entrada.")


def s3(doc, wb):
    doc.add_heading("3. Demanda y objetivos de servicio", 1)
    para(doc, "200.000 usuarios por 1,5 peticiones dan 300.000 al día: 3,47 por segundo en promedio y, con factor 6, "
              "**20,83 por segundo en el pico**. Los tokens se cuentan aparte porque no cuestan lo mismo: en el pico "
              "entran unos 83.333 por segundo y salen unos 10.417. El servicio lee ocho veces más de lo que escribe, y "
              "eso pesa en el prefill. Con demanda doble, todo se duplica: 41,67 peticiones por segundo.")
    para(doc, "Con la ley de Little y un tiempo promedio de 20 s en el sistema (supuesto), hay unas **417 peticiones a "
              "la vez** en el pico. No son usuarios conectados ni conversaciones abiertas; incluyen las que esperan y "
              "las que se procesan.")
    para(doc, "La sensibilidad deja una sorpresa. El §14.4 pide probar una entrada promedio del doble, 8.000 tokens. "
              "Con 500 de salida suman 8.500, y eso **no cabe** en la ventana de 8.192: ese escenario obliga a ampliar "
              "la ventana o a recortar contexto.")
    para(doc, "Los tiempos se miden desde el cliente (red, fila y proceso), y los rechazos y fallas se reportan, no se "
              "descartan. Disponibilidad y latencia son indicadores distintos: un servidor encendido puede responder "
              "tarde.")


def s4(doc, wb):
    doc.add_heading("4. Modelo y memoria", 1)
    para(doc, "Se compararon dos versiones del mismo modelo, Llama 3.1 70B Instruct en BF16 (la referencia) y en FP8 "
              "(versión de RedHat). Al ser la misma familia, lo único que cambia es cómo se guardan los números.")
    table(doc, ["Criterio", "BF16", "FP8"], [
        ("Parámetros y licencia", "70.550 millones, densos; Llama 3.1 Community", "Los mismos"),
        ("Español", "Soportado; MMLU en español 80,05", "Sin evaluar: fuera de alcance según su ficha"),
        ("Calidad publicada", "MGSM 86,9; IFEval 87,5", "99,88 % del original, solo en inglés"),
        ("Pesos / KV cache", "BF16 / BF16", "FP8 en capas lineales / KV en 16 bits"),
    ], widths=[3.6, 6.5, 6.5])
    para(doc, "Desde el config.json (80 capas, 8 cabezas KV, dimensión 128) salen los datos clave: **141,1 GB de pesos "
              "en BF16** (72,7 GB en FP8) y 0,33 MB de KV cache por token, es decir **1,47 GB por petición** de 4.500 "
              "tokens y 2,68 GB con la ventana llena.")
    ws = wb["Modelo y memoria"]
    want = {"MEM_PESOS": 1, "KV_DISP": 1, "PET_MAX_PROM": 0, "PET_MAX_PEOR": 0}
    rows = []
    for r in ws.iter_rows(min_row=1, values_only=True):
        if r[0] in want:
            d = want[r[0]]
            rows.append((r[1], fmt_num(r[2], d), fmt_num(r[3], d), fmt_num(r[4], d)))
    table(doc, ["Por réplica, con 80 % de memoria utilizable", "BF16 · 4 GPUs", "FP8 · 4 GPUs", "FP8 · 2 GPUs"],
          rows, widths=[7.6, 3, 3, 3])
    para(doc, "Que el modelo quepa no dice si alcanza. En BF16 caben 50 peticiones promedio por réplica, pero solo 27 "
              "si todas usan la ventana completa, y el ejemplo del enunciado con 40 de 8.192 tokens se pasa. Por eso "
              "la admisión reserva **tokens**, no solo cupos. FP8 ahorraría GPUs, pero no hay forma de afirmar que "
              "mantenga el 90 % de calidad en español: el diseño va en BF16 y FP8 queda como ahorro a futuro. La "
              "calidad no se pudo medir; el plan de evaluación con las 100 tareas está en el anexo D.")


def s5(doc, wb):
    doc.add_heading("5. Evidencia de capacidad", 1)
    para(doc, "No se midió nada. Se usa el dato sintético del §8.4: una réplica de 4 H100 atiende **2 peticiones por "
              "segundo** del perfil 4.000/500 cumpliendo todos los SLOs. Son 8.000 tokens de entrada y 1.000 de salida "
              "por segundo en el mismo punto, no dos máximos separados. Lo que sí se puede revisar es que no rompa la "
              "física de la H100:")
    ws = wb["Rendimiento"]
    vals = {r[0]: r[2] for r in ws.iter_rows(min_row=5, values_only=True) if r[0]}
    table(doc, ["Prueba", "Resultado", "Lectura"], [
        ("Cómputo que pide el punto de operación", f"{fmt_num(vals['FLOPS_USADOS'])} de {fmt_num(vals['FLOPS_PICO'])} "
         f"TFLOPS ({fmt_pct(vals['USO_COMPUTO'])})", "Posible, ni imposible ni sospechoso"),
        ("Peticiones generando a la vez", f"{fmt_num(vals['SEC_DECODE'])} de 50 que caben", "Cabe en memoria"),
        ("Paso mínimo de generación", f"{fmt_num(vals['T_PASO_MIN'], 1)} ms", "Por debajo de 40 ms"),
    ], widths=[6, 5.4, 5.2])
    para(doc, "Las pruebas dicen que el número es posible, no que se alcance. El cómputo denso de la H100 sale de "
              "dividir entre dos la cifra con dispersión que publica NVIDIA, y eso también es un supuesto. MLPerf "
              "v5.0 tiene una prueba interactiva de Llama 2 70B con 40 ms por token: confirma que la meta es "
              "realista, pero con otro modelo y textos más cortos. El procedimiento para medirlo de verdad está en el "
              "anexo E.")


def s6(doc, wb):
    doc.add_heading("6. Réplicas y nodos", 1)
    para(doc, "Una réplica es una copia del modelo en 4 GPUs; un nodo es un DGX H100 con 8 GPUs, o sea 2 réplicas. Cada "
              "réplica se carga al 70 % (1,4 peticiones/s) para absorber la variación del pico. Hacen falta 15 "
              "réplicas, 8 nodos, más uno para poder perderlo: **9 nodos, 72 GPUs y 18 réplicas**. El pico usa el "
              "57,9 % de lo instalado, y el 65,1 % si falta un nodo.")
    ws = wb["Inventario"]
    rows, on = [], False
    for r in ws.iter_rows(min_row=1, values_only=True):
        if r[1] == "Falla":
            on = True
            continue
        if on and r[1] and r[2] is not None and not str(r[1]).startswith('"'):
            rows.append((r[1], fmt_num(r[2]), fmt_num(r[3], 1), r[4], fmt_num(r[5], 1)))
        if on and len(rows) == 4:
            break
    table(doc, ["Falla", "Réplicas", "Con holgura (pet/s)", "¿Cubre el pico?", "Al 100 % (pet/s)"], rows,
          widths=[5, 2, 3.4, 3, 3.2])
    para(doc, "Se cubre la pérdida de un nodo, que es lo pedido. Un rack completo no: quedan 16,8 peticiones/s con "
              "margen, y solo al 100 % alcanzaría. Prometerlo exige 11 nodos. Además, que la capacidad alcance no "
              "prueba la latencia durante la falla; eso se ve en la sección 11.")


def s7(doc, wb):
    doc.add_heading("7. Arquitectura y despliegue", 1)
    figure(doc, DIAG / "arquitectura_logica.png", "Figura 1. Camino lógico de una petición.", width=16.5)
    para(doc, "**Petición normal.** El balanceador termina TLS; el gateway revisa identidad, cuota y que la entrada no "
              "pase de 8.192 tokens; el router elige una réplica con espacio en su presupuesto de tokens; el motor "
              "(por ejemplo vLLM) la suma a su lote continuo y devuelve la respuesta palabra por palabra. Las métricas "
              "guardan tiempos y errores, nunca el texto.")
    para(doc, "**Sobrecarga.** Sin espacio, la petición espera en una fila de **1 segundo** como máximo: de los 2 s "
              "del TTFT se descuentan 0,2 de red y 0,5 de procesar la entrada, y quedan 1,3. Después se responde "
              "\"intenta en un momento\". La fila no pasa de unas 21 peticiones; cada usuario tiene 3 a la vez y 20 "
              "por minuto (nunca más del 1,6 % del pico); si cierra el chat, su generación se cancela; y hay un solo "
              "reintento, solo si aún no recibió nada.")
    para(doc, "**Falla de un nodo.** El router deja de recibir sus latidos y saca sus 2 réplicas. Las respuestas a "
              "medias se pierden y el usuario reintenta: no se promete continuarlas. Quedan 22,4 peticiones/s, por "
              "encima del pico.")
    figure(doc, DIAG / "despliegue.png", "Figura 2. Despliegue: 3 racks de cómputo y uno de servicios.", width=12.5)
    para(doc, "**Paralelismo y caché.** Hay paralelismo de datos (18 réplicas) y tensorial de 4 GPUs dentro de cada "
              "servidor por NVLink; no hace falta de tubería ni de expertos. El caché de prefijos se comparte solo "
              "dentro de un mismo cliente, y no se reutilizan respuestas completas.")
    para(doc, "**Escalado.** Las 18 réplicas quedan siempre cargadas, porque recargar 141 GB no es instantáneo. Escalar "
              "contenedores no crea GPUs ni electricidad: crecer es comprar nodos cuando el uso pase del 70 % de forma "
              "sostenida, aparezcan rechazos o el TTFT p95 se acerque a 2 s.")
    para(doc, "**Operación y seguridad.** Despliegues con una réplica de prueba y vuelta atrás; modelos con versión y "
              "huella; secretos en un gestor; cifrado en tránsito y en reposo; aislamiento entre clientes; registros "
              "sin conversaciones. Las instrucciones escondidas en documentos se tratan como texto, y el modelo no "
              "ejecuta herramientas.")
    figure(doc, DIAG / "mapa_trafico.png", "Figura 3. Los tres tipos de tráfico.", width=16.5)
    para(doc, "**Red y almacenamiento.** El tráfico de usuarios es pequeño, unos 15 Mb/s en el pico. Entre las 4 GPUs "
              "de una réplica circulan unos 24 GB/s: el 2,6 % de NVLink, pero 7,5 veces lo que cabe en un enlace de "
              "25 Gb/s, y por eso una réplica no se reparte entre servidores. Cargar el modelo toma 45 s por réplica "
              "con 25 Gb/s y unos 4 minutos para las 18 a la vez, así que cada servidor guarda los pesos en su disco. "
              "El almacenamiento central necesita 3,9 TB brutos, y restaurar una versión desde la copia externa "
              "tarda unos 19 minutos.")


def s8(doc, wb):
    doc.add_heading("8. Ubicación e infraestructura", 1)
    para(doc, "Se eligió **colocation**: el proveedor pone edificio, electricidad, enfriamiento, UPS y conectividad, y "
              "nosotros los servidores. Se compararon tres mercados:")
    ws = wb["Ubicación"]
    keep = {"Precio de la electricidad": (3, "USD/kWh"), "Emisiones de la red eléctrica": (3, "kg CO₂e/kWh"),
            "Renta de espacio (punto medio del rango)": (0, "USD/kW/mes"),
            "Capacidad disponible en el mercado": (1, "MW"), "Riesgo sísmico": (0, ""),
            "Peor tiempo de red para un usuario": (0, "ms"), "¿Cabe en la red supuesta para el TTFT?": (0, ""),
            "Electricidad + renta": (0, "USD/año"), "Emisiones de operación": (0, "t CO₂e/año")}
    rows = []
    for r in ws.iter_rows(min_row=5, values_only=True):
        if r[1] in keep:
            d, unit = keep[r[1]]
            rows.append((r[1], fmt_num(r[2], d), fmt_num(r[3], d), fmt_num(r[4], d), unit))
    table(doc, ["", "Querétaro", "São Paulo", "Santiago", "Unidad"], rows, widths=[6, 2.6, 2.6, 2.6, 2.8])
    para(doc, "São Paulo sale mejor en casi todo, pero lo decisivo es la red: con 30 ms de la conexión del usuario, "
              "su peor caso es Ciudad de México con 188 ms, dentro de los 0,2 s del TTFT. Desde Querétaro, Bogotá "
              "queda a 223 ms (la ruta pasa por Miami), y desde Santiago, Ciudad de México a 208. Hay dos "
              "advertencias: la renta de CBRE es para contratos de 250 a 500 kW y el nuestro ronda los 104, y CBRE "
              "menciona fallas eléctricas en zonas de São Paulo, así que el proveedor debe tener UPS y generadores.")
    para(doc, "**Diseño físico.** 4 racks: 3 de cómputo con 3 DGX y uno de red y servicios. Cada rack de cómputo "
              "ocupa 24 de 42 U, pesa unos 391 kg en servidores y consume **30,6 kW**. El límite no es el espacio sino "
              "el calor: los DGX se enfrían por aire, así que hace falta pasillo contenido o intercambiadores en la "
              "puerta trasera. Si el proveedor solo da ~25 kW por rack, se pasa a 2 servidores por rack y 6 racks.")
    para(doc, "**Descartadas con números.** Marte: la señal tarda de 3 a 22 minutos solo de ida, imposible para un "
              "TTFT de 2 s. Órbita: solo para radiar 103,8 kW harían falta al menos 251 m² de radiadores ideales, "
              "más radiación, eclipses y lanzamientos. Submarino: cada reparación de una GPU obliga a sacar la "
              "cápsula, y Natick fue un experimento.")


def s9(doc, wb):
    doc.add_heading("9. Electricidad y enfriamiento", 1)
    ws = wb["Energía"]
    keep = {"E_P_TI": 1, "E_P_SITIO": 1, "E_HOLG_ACOM": "%", "E_FRIO_DISP": 0, "E_FRIO_USO": "%",
            "E_CAUDAL_LMIN": 0}
    rows = []
    for r in ws.iter_rows(min_row=5, max_row=22, values_only=True):
        if r[0] in keep:
            d = keep[r[0]]
            f = (lambda v: fmt_pct(v)) if d == "%" else (lambda v, d=d: fmt_num(v, d))
            rows.append((r[1], f(r[2]), f(r[3]), f(r[4]), r[5]))
    table(doc, ["kW = potencia; kWh = energía", "Normal", "Máxima", "Degradada", "Unidad"], rows,
          widths=[6.4, 2.4, 2.4, 2.6, 2.8])
    para(doc, "La TI máxima es 9 × 10,2 kW más 12 kW: 103,8 kW. La potencia del servidor ya incluye sus GPUs y no se "
              "suman otra vez. Con 35 kW que no son de TI, el sitio pide 138,8 kW y los 160 contratados dejan 13 % de "
              "margen. En promedio la TI usa 70,2 kW; con PUE de 1,30 son **799 MWh al año**, unos USD 117.000. Tres "
              "módulos de 60 kW dejan 120 kW si uno falla, y el agua del circuito necesita unos 149 L/min. La UPS "
              "protege 120 kW durante 10 minutos (unos 35 kWh nominales): cubre la TI, no todo el sitio.")
    para(doc, "Dos hallazgos. NVIDIA declara 38.557 BTU/h por DGX, que son 11,3 kW y no 10,2; aun así el "
              "enfriamiento con un módulo fuera alcanza. Y con demanda doble el sitio pediría unos 234 kW, mucho más "
              "que los 160 contratados: crecer exige negociar potencia, no solo comprar servidores.")
    para(doc, "En lo ambiental, la red de Brasil deja unas **88 t CO₂e al año** (en Querétaro serían 379). El agua "
              "directa se supuso en cero por circuito cerrado; con torres evaporativas serían 615 a 1.230 m³ al año, "
              "y el dato hay que pedirlo. Cada cinco años habrá que reciclar 1,2 t de servidores. Un PUE bajo no dice "
              "nada sobre emisiones ni agua.")


def s10(doc, wb):
    doc.add_heading("10. Presupuesto", 1)
    para(doc, "En dólares, a noviembre de 2026. Salvo donde se indica son los valores hipotéticos del enunciado. NVIDIA "
              "no publica precio oficial del DGX H100; los USD 350.000 usados quedan en el rango público de 300.000 a "
              "400.000. El detalle con fuentes está en el Excel.")
    ws = wb["Presupuesto"]
    capex, vals, section = [], {}, None
    for r in ws.iter_rows(min_row=4, values_only=True):
        if r[0] and str(r[0]).startswith("CAPEX ("):
            section = "c"
            continue
        if r[0] and str(r[0]).startswith("OPEX ("):
            section = None
        if r[0] and r[4] is not None:
            vals[r[0]] = r[4]
        if section == "c" and r[0] and r[4] is not None and r[0] not in (
                "IMPUESTO_IMPORT", "CONTINGENCIA", "CX_IMPORTADO"):
            capex.append((r[1], fmt_num(r[4])))
    table(doc, ["CAPEX", "USD"], capex, widths=[12.6, 4])
    para(doc, "Al ser colocation se quitaron la obra eléctrica, la UPS, el enfriamiento y el edificio del ejemplo (USD "
              "1,2 millones que el proveedor ya cobra en la renta) y se sumaron USD 100.000 para adecuar el espacio. "
              f"El OPEX anual es de USD {fmt_num(vals['OPEX'])}: electricidad {fmt_num(vals['OX_ELEC'])}, renta "
              f"{fmt_num(vals['OX_RENTA'])}, personal con guardia 24/7 {fmt_num(vals['OX_PERSONAL'])} y el resto en "
              "conectividad, mantenimiento, software, seguros, respaldo y contingencia.")
    para(doc, f"El TCO a tres años es de **USD {fmt_num(vals['TCO'])}**, sin descuento financiero ni valor residual. "
              f"Cada petición cuesta USD {fmt_num(vals['COSTO_PET'], 4)}, y USD {fmt_num(vals['COSTO_PET_BUENA'], 4)} "
              "si solo se cuentan las que llegan a tiempo y sin error (una cota, porque la calidad no está medida). "
              f"Por millón de tokens de entrada y salida son USD {fmt_num(vals['COSTO_MTOK'], 2)}; no se compara con "
              "tarifas de API que cobran solo una parte. Se cumplen los tres límites: CAPEX al 95 %, OPEX al 91 % y "
              "TCO al 93 %. Es poco margen:")
    ws = wb["Escenarios"]
    hdr = [c for c in next(ws.iter_rows(min_row=4, max_row=4, values_only=True))]
    sv = {r[0]: r for r in ws.iter_rows(min_row=5, values_only=True) if r[0]}
    rows = []
    for i in range(2, len(hdr)):
        if hdr[i] is None:
            continue
        rows.append((hdr[i], fmt_num(sv["S_NODOS"][i]), fmt_num(sv["S_CAPEX"][i] / 1e6, 2),
                     fmt_num(sv["S_TCO"][i] / 1e6, 2), fmt_num(sv["S_COSTO_PET"][i], 4), sv["S_OK_TCO"][i]))
    table(doc, ["Escenario", "Nodos", "CAPEX (M)", "TCO (M)", "USD/petición", "¿TCO cabe?"], rows,
          widths=[5, 1.6, 2.4, 2.4, 2.6, 2.6])
    para(doc, "La demanda es lo que más mueve el costo, aunque con el doble cada petición sale más barata. Lo que más "
              "mueve el **riesgo** es el rendimiento real y el precio del hardware: un 30 % de desvío en cualquiera "
              "rompe el CAPEX. La energía casi no cambia nada, y calificar al régimen ReData bajaría el TCO a unos "
              "USD 10,2 millones.")


def s11(doc, wb):
    doc.add_heading("11. Resiliencia y crecimiento", 1)
    para(doc, "Un 99,9 % mensual deja **43,2 minutos** de caída en 30 días, y una sola intervención mal hecha se los "
              "puede gastar. El presupuesto de errores es el 0,1 % de 9 millones: 9.000 al mes. La matriz completa de "
              "diez fallas y los objetivos de recuperación están en el anexo C. En resumen:")
    bullets(doc, [
        "**Cubiertas:** GPU, nodo, switch o enlace, una fuente o un corte eléctrico, un módulo de enfriamiento, el "
        "almacenamiento central, una actualización defectuosa (vuelta atrás en ~15 min) y la saturación.",
        "**No cubiertas, y aceptadas:** un rack completo (queda sin margen; cubrirlo cuesta ~USD 1,5 millones más) y "
        "la pérdida del sitio (semanas para comprar e importar equipo; no hay segundo sitio en el presupuesto).",
        "**RTO/RPO:** segundos ante fallas de hardware, sin pérdida de datos porque no se guardan conversaciones; "
        "~19 min para traer los pesos desde fuera; 1 h de RPO para los datos de la aplicación.",
    ])
    para(doc, "**Ejercicio: cae un nodo en plena hora pico.** En los 30 s de detección se pierden unas 46 peticiones en "
              "curso y unas 69 que el router alcanza a mandarle: 116 errores, el 1,3 % del presupuesto del mes. "
              "Quedan 22,4 peticiones/s, por encima del pico. Si cae otro nodo antes de las ~48 h de reparación, "
              "quedan 19,6 con margen: se atiende, pero al límite. Las pruebas se ensayan: apagar un nodo cada "
              "trimestre y restaurar copias cada mes y cada semestre, porque tener la copia no prueba que se pueda "
              "restaurar a tiempo.")
    para(doc, "**Crecer.** Con demanda doble hacen falta 16 nodos, unos 234 kW y un TCO de USD 16,8 millones, que se "
              "presenta aparte del caso base. Comprar y negociar potencia toma meses: el aviso tiene que llegar antes.")


def s12(doc):
    doc.add_heading("12. Recomendación y condiciones que podrían cambiarla", 1)
    para(doc, "Se recomienda instalar 9 DGX H100 en colocation en São Paulo con Llama 3.1 70B BF16, con cuatro "
              "condiciones:")
    bullets(doc, [
        "**Medir antes de comprar** que una réplica de 4 H100 atiende cerca de 2 peticiones/s dentro de los SLOs. "
        "Con un 30 % menos se necesitan 12 nodos y el diseño se sale del presupuesto.",
        "**Cotizar** el DGX, la renta para unos 104 kW y los impuestos reales en Brasil, con y sin ReData.",
        "**Confirmar con el proveedor** los ~31 kW por rack con aire, qué enfriamiento queda en UPS y cuánta agua "
        "consume.",
        "**Evaluar la calidad** con las 100 tareas en español antes de pasar a FP8 o a otro modelo.",
    ])
    para(doc, "La recomendación cambiaría si el rendimiento medido es mucho mejor (sobran nodos), si FP8 mantiene la "
              "calidad en español (menos GPUs), si la demanda se concentra en México (Querétaro merecería otra "
              "mirada) o si el negocio no acepta semanas de caída ante la pérdida del sitio (habría que presupuestar "
              "un segundo sitio, que hoy no cabe).")


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

    doc.add_heading("Anexo C. Matriz de fallas y recuperación", 1)
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
    table(doc, ["Falla", "Detección", "Impacto", "Capacidad con holgura (pet/s)", "¿Cubre?", "Acción",
                "Recuperación", "Riesgo que queda"], rows, widths=[2.6, 3.4, 3.4, 1.8, 1.4, 3.8, 3, 4.4], size=8)
    table(doc, ["Qué", "RTO", "RPO", "Dónde está la copia", "Cómo se prueba"], [
        ("Falla de GPU, nodo o enlace", "Segundos", "Nada (no se guardan conversaciones)", "No aplica",
         "Apagar un nodo a propósito cada trimestre"),
        ("Versión defectuosa", "~15 min", "Nada", "Versión anterior en disco local",
         "Cada despliegue empieza con una réplica de prueba"),
        ("Pesos del modelo", "~19 min desde fuera", "Nada (copias versionadas)",
         "Disco local, almacenamiento central y copia externa", "Restaurar desde fuera cada semestre"),
        ("Datos de la aplicación", "~1 h", "1 h", "Almacenamiento central y copia externa",
         "Restaurar en un entorno de prueba cada mes"),
        ("Registros y métricas", "No urgente", "Hasta 24 h", "Solo en el sitio", "No se respaldan fuera"),
        ("Pérdida del sitio", "Semanas", "1 h de datos", "Copia externa", "Ejercicio de mesa anual"),
    ], widths=[4.4, 3, 4.6, 6, 6], size=8)
    para(doc, "Dos componentes no son alta disponibilidad si comparten alimentación, red, configuración o una causa "
              "común de falla.", italic=True, size=9, color=MUTED)

    portrait_section(doc)
    doc.add_heading("Anexo D. Plan de evaluación de calidad", 1)
    para(doc, "No se entregó el conjunto de 100 tareas y no hubo dónde correr el modelo, así que no hay puntajes. "
              "El plan es este:")
    bullets(doc, [
        "Congelar la revisión exacta del modelo de referencia y del candidato antes de evaluar.",
        "Armar 100 tareas en español: 40 de comprensión y respuesta sobre un texto dado, 20 de resumen, 20 de seguir "
        "instrucciones, 10 de contexto largo (cerca de 8.000 tokens) y 10 de rechazo o de instrucciones maliciosas "
        "escondidas en documentos.",
        "Calificar cada tarea de 0 a 2 en corrección, seguimiento de instrucciones y apego al texto. Califican "
        "personas; un LLM juez puede apoyar, pero no decidir solo.",
        "Aceptar el candidato si saca al menos el 90 % del puntaje de la referencia y no tiene ninguna falla crítica "
        "de privacidad o de mezcla de información entre usuarios.",
    ])
    doc.add_heading("Anexo E. Plan de validación del rendimiento", 1)
    bullets(doc, [
        "Fijar modelo y revisión, motor de inferencia y su versión, 4 × H100 con paralelismo tensorial de 4, "
        "política de lotes y estado del caché de prefijos.",
        "Calentar el sistema y subir la tasa de llegada en lazo abierto, desde una tasa baja, hasta que se rompa "
        "algún SLO. Una prueba de concurrencia fija escondería la sobrecarga.",
        "Repetir al menos tres veces y reportar TTFT, tiempo por token y tiempo total en p95, peticiones "
        "completadas, rechazadas, canceladas y fallidas, y consumo eléctrico medido.",
        "Repetir con la mezcla de peticiones cortas y largas que conserva los promedios, con arranque en frío y con "
        "la pérdida de un nodo.",
        "No reutilizar prefijos guardados entre pruebas sin decirlo.",
    ])
    doc.add_heading("Anexo F. Fuentes", 1)
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
    for section in (s1, s2):
        section(doc)
    for section in (s3, s4, s5, s6, s7, s8, s9, s10, s11):
        section(doc, wb)
    s12(doc)
    annex(doc, wb)
    OUT.parent.mkdir(parents=True, exist_ok=True)
    doc.save(OUT)
    print(f"Saved {OUT.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
