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
