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
