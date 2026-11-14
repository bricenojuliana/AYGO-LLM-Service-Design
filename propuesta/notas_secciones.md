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
