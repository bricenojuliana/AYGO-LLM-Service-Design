# Notas para la propuesta (borrador interno)

Lo que debe ir en cada sección del Word. Se completa paso a paso; el texto final se redacta al final.

## 3. Demanda y SLOs

- Cadena completa: 200.000 usuarios × 1,5 = 300.000 peticiones/día → 3,47/s en promedio → 20,83/s en el pico (factor 6).
- Tokens en el pico, por separado: ~83.333/s de entrada y ~10.417/s de salida. La relación es 8 a 1, así que el prefill pesa mucho.
- Concurrencia con Little: ~417 peticiones en el pico, usando 20 s promedio en el sistema. Ese tiempo es un supuesto que se ajusta en el paso 3.
- Aclarar que la concurrencia no son usuarios conectados ni sesiones.
- Escenario ×2: 41,67/s en el pico, ~833 peticiones simultáneas.
- Hallazgo para destacar: duplicar la entrada promedio (8.000) no cabe con 500 de salida en la ventana de 8.192. Ese escenario del §14.4 obliga a ampliar la ventana o a recortar contexto, y hay que decirlo.
- SLOs: copiar la tabla del §4.3 y explicar que se miden desde el cliente (red + cola + proceso), y que disponibilidad y latencia son indicadores distintos.

## 4. Modelo y memoria

- Se comparan dos alternativas del mismo modelo: Llama 3.1 70B Instruct en BF16 (referencia) y en FP8 (RedHat). GPU: H100 de 80 GB.
- Parámetros calculados desde config.json: 70.550 millones. Coincide con los "71B" de Hugging Face.
- Pesos: 141,1 GB en BF16 y 72,7 GB en FP8. En FP8 los embeddings y lm_head siguen en 16 bits.
- KV cache: 0,33 MB por token, 1,47 GB por petición promedio y 2,68 GB con la ventana completa.
- BF16 en 4×H100: caben 50 peticiones promedio o 27 en el peor caso. El ejemplo del enunciado con 40 peticiones de 8.192 tokens no cabe.
- FP8 en 4×H100: caben 97 promedio o 53 en el peor caso. FP8 en 2×H100 deja solo ~15 GB para KV (10 peticiones), muy justo.
- Mensaje clave: caber en memoria no es atender a tiempo. Solo por memoria harían falta 9 réplicas BF16 para el KV del pico; el número real sale de rendimiento (pasos 3 y 4).
- Riesgo grande de FP8: su ficha deja fuera de alcance los idiomas distintos al inglés. La calidad en español queda pendiente de evaluar y no se puede afirmar que llegue al 90 %.
- La política de admisión debe limitar tokens en KV, no solo número de peticiones.

## 5. Evidencia de capacidad

- No hay medición propia. Se usa el dato sintético del §8.4: 2 peticiones/s por réplica de 4×H100 (8.000 tokens/s de entrada y 1.000 de salida a la vez). Siempre decir "supuesto", nunca "medimos".
- Prueba de cómputo: ese punto pide ~1.270 TFLOPS de ~3.960 densos, un 32 %. Es físicamente posible.
- Prueba de memoria: con 40 ms por token hay ~40 peticiones generando a la vez, y caben 50 en el KV. El paso mínimo de generación toma ~14,7 ms, por debajo de 40 ms.
- Coherencia: 2 peticiones/s × 20 s = 40 peticiones en la réplica, el mismo tiempo supuesto en la hoja Demanda.
- MLPerf v5.0 tiene una prueba interactiva de Llama 2 70B con 40 ms por token: sirve para decir que la meta es realista, no para sacar cifras.
- Sensibilidad: con 30 % menos de rendimiento, 1,4 peticiones/s por réplica.
- FP8 queda como variante a futuro: podría ahorrar GPUs, pero falta medir velocidad y calidad en español.
- Incluir el plan de validación (procedimiento del §8.2) tal como está en la hoja Rendimiento.

## 6. Réplicas y nodos

- Diferenciar réplica (copia del modelo, 4 GPUs) y nodo (servidor DGX H100, 8 GPUs, 2 réplicas).
- Con 70 % de utilización: 15 réplicas mínimas → 8 nodos sin fallas → 9 nodos para tolerar la pérdida de uno. Total: 72 GPUs, 18 réplicas, 3 racks de cómputo.
- El pico usa el 57,9 % de la capacidad instalada, y el 65,1 % tras perder un nodo.
- Por réplica: ~23 peticiones en curso en el pico y ~26 tras la falla. Caben en memoria (50).
- Fallas: GPU y nodo cubiertos con holgura. Un rack completo NO: quedan 12 réplicas y 16,8 peticiones/s con holgura. Al 100 % sí alcanzaría (24/s), pero sin margen. Para prometer tolerancia a rack hacen falta 11 nodos.
- Escenarios: demanda ×2 → 16 nodos; rendimiento −30 % → 12 nodos; ambos → 23 nodos.
- Recordar que la capacidad no prueba la latencia durante la falla (detección, reintentos, calentamiento).

## 7. Arquitectura y despliegue

Diagramas: `diagramas/arquitectura_logica.png` y `diagramas/despliegue.png`. Componentes por función, con un producto de ejemplo entre paréntesis. Sin RAG (posible extensión).

**Petición normal.** El usuario abre HTTPS contra el balanceador, que termina TLS. El gateway revisa identidad, cuotas y que la entrada no pase de 8.192 tokens. El router busca una réplica con espacio en su presupuesto de tokens y la manda allí. El motor la mete en el lote continuo, procesa la entrada y devuelve la respuesta palabra a palabra. Las métricas guardan tiempos y errores, no el texto.

**Sobrecarga.** Si ninguna réplica tiene espacio, la petición espera en una fila corta (máximo 1 s; caben 1,3 s en el presupuesto del TTFT). Pasado ese tiempo se responde "intenta en un momento" (429 o 503 con Retry-After). Fila máxima: ~21 peticiones en todo el servicio. Límites por usuario: 3 a la vez y 20 por minuto, así que un usuario solo puede ocupar el 1,6 % del pico. Si el usuario cierra la conexión, se cancela su generación para liberar memoria. Un solo reintento, y solo si no se envió nada. No hay tráfico por lotes en el escenario base; si lo hubiera, iría con menor prioridad.

**Falla de un nodo.** El router deja de recibir latidos de sus 2 réplicas y las saca del reparto. Las peticiones que estaban a medias se pierden: el usuario ve un error y puede reintentar (no prometemos continuidad de respuestas parciales). Quedan 16 réplicas con 22,4 peticiones/s de capacidad admitida, por encima del pico de 20,83.

**Paralelismo.** Datos: 18 réplicas independientes. Tensorial: 4 GPUs dentro del mismo servidor, por NVLink. No se usa paralelismo de tubería ni de expertos (el modelo es denso y cabe en 4 GPUs).

**Lotes y caché.** Lotes continuos para aprovechar la GPU sin hacer esperar a nadie. Caché de prefijos solo dentro del mismo cliente (por ejemplo, las instrucciones de sistema), nunca entre usuarios, y sin suponer aciertos del 100 %. El KV cache de una conversación no es un caché de respuestas: no se reutilizan respuestas completas.

**Admisión por memoria.** Al admitir se reservan entrada + salida máxima (5.000 tokens en el perfil promedio). Así caben ~45 peticiones por réplica, más que las ~26 que habría tras perder un nodo.

**Escalado.** El límite es físico: 18 réplicas. Todas quedan cargadas y calientes; no se apagan en horas valle porque cargar 141 GB de pesos toma tiempo (se calcula en el paso 6). Para crecer hay que comprar nodos, y con ellos energía y refrigeración: escalar contenedores no crea GPUs. Indicadores para ampliar: uso sostenido por encima del 70 %, rechazos por fila llena o TTFT p95 acercándose a 2 s.

**Operación y seguridad.** Despliegues graduales por réplica (una primero, luego el resto) con vuelta atrás. Registro de modelos con versión y huella de los pesos. Secretos en un gestor, nunca en el código. Cifrado en tránsito y en reposo. Aislamiento entre clientes en cuotas, caché y registros. Registros sin contenido de conversaciones por defecto, con retención corta y borrado a pedido. Las instrucciones maliciosas dentro de documentos se tratan como texto del usuario, no como órdenes; el modelo no ejecuta herramientas en el escenario base.

## 7 (cont.). Red y almacenamiento

Diagrama: `diagramas/mapa_trafico.png`.

- Tres tráficos muy distintos. El de los usuarios es pequeño: ~15 Mb/s en el pico (~30 Mb/s con demanda ×2). Pesan más las respuestas que las preguntas, porque cada palabra viaja en su propio mensaje.
- Entre las 4 GPUs de una réplica: ~24 GB/s, el 2,6 % de NVLink. Por un enlace de 25 Gb/s sería 7,5 veces más de lo que cabe: esa es la razón numérica para no repartir una réplica entre servidores.
- Cargar el modelo: 45 s ideales por réplica con 25 Gb/s (el enunciado da 44,8 s). Arrancar las 18 a la vez desde el almacenamiento central: ~254 s, limitado por la lectura (10 GB/s supuesta). Por eso cada DGX guarda copia local de los pesos en su disco de 30,7 TB.
- Almacenamiento central: ~640 GB útiles (3 versiones del modelo, imágenes, datos, registros de 30 días) → ~3,9 TB brutos con 3 copias y holgura ×2.
- Copia fuera del sitio: ~0,6 TB. Restaurar una versión del modelo por 1 Gb/s tarda ~19 min; esto entra en el RTO del paso 10. Tener la copia no prueba que se pueda restaurar a tiempo: hay que ensayarlo.
- No se guardan conversaciones, así que no hay que respaldarlas.

## 8. Ubicación e infraestructura

- Colocation: el proveedor pone edificio, electricidad, enfriamiento, seguridad y conectividad; nosotros compramos y operamos los servidores. Así se aprovecha el ejemplo del enunciado de "adaptar una instalación existente".
- Se compararon Querétaro, São Paulo y Santiago (hoja Ubicación). Gana São Paulo: electricidad más barata (0,146 USD/kWh) y más limpia (0,110 kg CO₂e/kWh), espacio más barato (USD 130–190/kW/mes), más capacidad libre (51,5 MW) y riesgo sísmico bajo.
- Hallazgo que pesa mucho: con 0,2 s de red en el presupuesto del TTFT, solo São Paulo cumple para el peor usuario (188 ms con última milla, hacia Ciudad de México). Querétaro (223 ms hacia Bogotá, porque la ruta pasa por Miami) y Santiago (208 ms) no cumplen.
- Al año en São Paulo: ~USD 117.000 de electricidad y ~USD 199.000 de renta; ~88 t CO₂e. En Querétaro serían ~USD 483.000 y ~379 t.
- Riesgos: el precio de renta de CBRE es para 250–500 kW (el nuestro es ~104 kW, puede salir más caro por kW); CBRE menciona problemas de confiabilidad de la red eléctrica en algunas zonas de São Paulo, así que el proveedor debe tener UPS y generadores. Brasil habla portugués, pero eso no importa: los usuarios están en el resto de la región.
- Diseño físico: 4 racks (3 de cómputo con 3 DGX cada uno + 1 de red y almacenamiento). Cada rack de cómputo: 24 U de 42, 30,6 kW, ~391 kg de servidores, 3.315 CFM de aire. El límite no es el espacio: es la energía y el calor. Si el proveedor solo soporta ~25 kW por rack, se pasa a 2 DGX por rack (6 racks).
- Los DGX H100 se enfrían por aire: hace falta pasillo contenido o intercambiadores en la puerta trasera para ~31 kW por rack. No presentar enfriamiento líquido directo, porque este equipo no lo usa.
- Opciones descartadas con números: Marte (la señal tarda 6 a 44 min ida y vuelta contra un TTFT de 2 s), órbita (mínimo ~251 m² de radiadores ideales, más radiación, eclipses y lanzamientos), submarino (cada reparación exige sacar la cápsula; Natick es un experimento).

## 9. Electricidad y enfriamiento

Hoja Energía, con tres casos lado a lado: normal (promedio), máxima y degradada (un nodo y un módulo de enfriamiento fuera).

- TI máxima: 9 × 10,2 kW + 12 kW = 103,8 kW. La potencia del DGX ya incluye sus GPUs; no se suman otra vez.
- Sitio: 138,8 kW en el punto de diseño (35 kW no TI del ejemplo). Con 160 kW contratados queda 13 % de holgura. En promedio el sitio usa ~91 kW.
- Energía: 799 MWh al año (TI promedio 70,2 kW × PUE 1,30 × 8.760 h). En São Paulo cuesta ~USD 117.000.
- Enfriamiento: 3 módulos de 60 kW. Con uno fuera quedan 120 kW, más que los 103,8 kW de calor. El agua del circuito necesita ~2,5 kg/s (~149 L/min) con 10 °C de subida.
- Hallazgo 1: NVIDIA declara 38.557 BTU/h por DGX, que son 11,3 kW, más que los 10,2 kW de entrada. Con ese valor la TI llegaría a ~114 kW y el enfriamiento N+1 igual alcanza (120 kW). Hay que aclararlo con NVIDIA o con una medición.
- Hallazgo 2: con demanda ×2 (16 nodos) el sitio pediría ~234 kW, por encima de los 160 kW contratados. Crecer exige negociar más potencia y enfriamiento con el proveedor, no solo comprar servidores.
- UPS: 120 kW durante 10 min = 20 kWh útiles → ~35 kWh nominales de baterías. Cubre la TI máxima, pero no todo el sitio (138,8 kW): hay que confirmar con el proveedor qué parte del enfriamiento queda protegida mientras arrancan los generadores. Las baterías también deben poder entregar 120 kW, no solo la energía.
- Ambiente: ~88 t CO₂e al año con la red de Brasil. Agua directa supuesta en 0 (circuito cerrado); con 1 o 2 L/kWh serían ~615 o ~1.230 m³ al año, así que hay que pedirle el WUE al proveedor. Residuos: ~1,2 t de servidores cada ~5 años, para reciclar.
- Recordar: kW es potencia y kWh es energía. Un PUE bajo no implica bajas emisiones ni poca agua.

## 10. Presupuesto

Hojas Presupuesto y Escenarios. Todo en USD, precios a nov-2026; los del enunciado son hipotéticos, no cotizaciones.

- CAPEX: USD 7,56 M (94,5 % del límite). Servidores 9 × 350.000 = 3,15 M; red 0,35 M; almacenamiento 0,25 M; racks 0,15 M; adecuación en colocation 0,10 M; integración 0,20 M; repuestos 0,15 M; impuestos de importación 1,95 M (50 % sobre 3,9 M importados); contingencia 20 %.
- Por ser colocation se quitaron obra eléctrica, UPS, enfriamiento y edificio (1,2 M del ejemplo): el proveedor los cobra en la renta. Si no, se pagarían dos veces.
- OPEX: USD 1,36 M al año (91 % del límite). Electricidad 0,117 M, renta 0,199 M, personal 0,6 M y el resto del ejemplo.
- TCO a 3 años: USD 11,64 M (93 % del límite). Cumple los tres límites.
- Costo por petición: USD 0,0354; contando solo las que llegan a tiempo y sin error, USD 0,0374. La calidad de la respuesta no está medida, así que este número es una cota.
- Por millón de tokens: USD 7,87 si se cuentan entrada y salida; USD 70,9 si todo se carga a la salida. No compararlo con una tarifa de API que cobra solo una parte.
- Escenarios: demanda ×2 sube el TCO a 16,8 M (pero baja el costo por petición a 0,026); rendimiento −30 % → 13,9 M; hardware +30 % → 13,8 M; energía +50 % → solo +0,18 M; tolerar un rack → 13,1 M; con ReData → 10,2 M.
- Conclusión: la variable que domina el costo es la demanda; las que dominan el riesgo son el rendimiento real por réplica y el precio del hardware, que rompen el CAPEX con un 30 % de desvío. Antes de comprar hay que medir el rendimiento y conseguir cotizaciones. La energía casi no mueve el resultado.
- Los impuestos de Brasil pesan 1,95 M: calificar a ReData ahorraría ~1,4 M en TCO.

## 11. Resiliencia y crecimiento

Hoja Resiliencia: matriz de 10 fallas, objetivos de recuperación y un ejercicio con números.

- Presupuesto del mes: 43,2 min de caída (99,9 % en 30 días) y 9.000 errores (0,1 % de 9 millones de peticiones). Una sola intervención mal hecha puede gastar todo el tiempo de caída.
- Cubierto con holgura: GPU, nodo, switch o enlace (red duplicada), una fuente o un corte de red eléctrica (UPS + generadores del proveedor), un módulo de enfriamiento, almacenamiento central (pesos en disco local), actualización defectuosa (réplica de prueba y vuelta atrás en ~15 min) y saturación (rechazos rápidos).
- No cubierto: un rack completo (quedan 16,8 peticiones/s con holgura; al 100 % alcanzaría, sin margen) y la pérdida del sitio (semanas para comprar e importar equipo). Se aceptan y se documentan. Cubrir el rack cuesta ~USD 1,5 M más de TCO.
- Ejercicio, nodo que cae en hora pico: se pierden ~46 peticiones en curso y ~69 que llegan a ese nodo durante los 30 s de detección. Total ~116 errores, el 1,3 % del presupuesto de errores del mes. Después quedan 22,4 peticiones/s, por encima del pico. Si cae un segundo nodo antes de los ~48 h de reparación, quedan 19,6 con holgura: se atiende, pero sin margen.
- RTO/RPO: fallas de hardware → segundos, sin pérdida de datos (no se guardan conversaciones); versión mala → ~15 min; pesos desde fuera → ~19 min; datos de la aplicación → RPO de 1 h con copias cada hora; registros → hasta 24 h, sin copia externa.
- Ensayos: apagar un nodo a propósito cada trimestre, restaurar una versión desde fuera cada semestre, restaurar datos cada mes en un entorno de prueba y un ejercicio de mesa anual de pérdida del sitio. Tener la copia no prueba que se pueda restaurar.
- Crecer: demanda ×2 pide 16 nodos y ~234 kW, más que los 160 kW contratados. Hay que negociar potencia y enfriamiento con el proveedor con meses de anticipación.
