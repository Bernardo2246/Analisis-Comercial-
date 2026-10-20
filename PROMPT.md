Actualiza el proyecto "Análisis Comercial Integra Cimentaciones" (repo GitHub Pages
bernardo2246.github.io/Analisis-Comercial-) con datos hasta [MES/AÑO] — trimestre [Qn AAAA].

FORMATO (desde Q3 2026)
1. Presentación para inversionistas = index.html (URL pública). Ejecutiva, cuenta una historia
   de resultados, sin planeación interna. 14 slides en un lienzo fijo 1920×1080 escalado con rem
   (misma retícula en 1920×1080, 1440×900 y 1280×720), navegación con flechas e índice (Esc).
2. Reporte PDF extenso para Bernardo = reporte/Reporte_Comercial_[Qn-AAAA].pdf, generado de
   reporte/reporte.html (carta, encabezado/pie con logo y "Página X de Y") con Chrome headless:
       bash reporte/build_pdf.sh
   Contiene todo lo de la presentación + trayectoria, recomendaciones, cierre, pipeline GHL,
   actividades, cobranza, explorador como tablas y anexos de calidad de datos y coordenadas.
   reporte.html, data_interno.js y el PDF NO se publican (están en .gitignore).

REGLA DE TRABAJO: por fases, con "aprobado" explícito de Bernardo entre fases; nada de commit
ni push hasta la fase de publicación. Si un dato no cuadra, se pregunta; no se "arregla".
  Fase 0 auditoría (solo lectura) → 1 GHL → 2 propuesta por slide → 3 muestras de diseño con
  capturas → 4 construcción y verificación → 5 publicación.

FUENTES
- Excel de ventas (único archivo que se comparte cada trimestre): copiarlo a
  data/ANALISIS_DE_DATOS_2.0.xlsx (gitignored: trae cobranza por cliente). Hojas por año con
  bloques mensuales y encabezados repetidos; 2026+ con columnas Abonado/Restante/Liquidado/Crédito.
- Go High Level: módulo ghl/ (ver ghl/README.md). Token Private Integration de solo lectura en
  ghl/.env. Si no hay token, fuente provisional = transcripción del tablero Master Metrics de SRM
  en ghl/raw/ + data/ghl_mastermetrics_[corte].json (ambos gitignored), marcada PRELIMINAR.
- Baselines congelados (no se vuelven a pedir): data/baseline_clientify_2026-07.json (funnel y
  actividades Clientify), data/forecast_publicado_[mes].json, data/corte_anterior_[mes].json.

QUÉ HACER CADA TRIMESTRE
1. python3 data/build.py — capa central ÚNICA (presentación y PDF leen lo mismo; no duplicar
   cálculos en el HTML). Genera data/data.js (público), reporte/data_interno.js (interno) y
   snapshots/pipeline_AAAA-MM.json. La consola compara cada mes contra los subtotales del Excel.
   - Ubicaciones nuevas: agregarlas a UBIC en build.py y sus coordenadas (cabecera municipal) a
     data/geo.json; mostrar la tabla municipio→coordenadas a Bernardo antes de usarla.
   - Valores nuevos de servicio/canal/tipo: agregarlos a los mapas de NORMALIZACIÓN.
2. Recalcula todo con el rango nuevo; el trimestre cerrado pasa de proyección a real.
3. Proyección vs real contra el forecast publicado el trimestre anterior; nuevo forecast del
   siguiente trimestre con la TENDENCIA DEL AÑO EN CURSO (regresión lineal; escenarios ± 1σ de
   residuales); guardar el forecast nuevo en data/forecast_publicado_[mes].json.
4. GHL: python3 ghl/extract.py → python3 ghl/summarize.py (validar conteos con Bernardo) →
   --write para el snapshot. Recompras y recomendados entran fuera del CRM (esperado).
5. Capturas de las 14 slides en las 3 resoluciones + PDF completo; verificar que cada cifra de
   las slides coincida con DATA.metrics.

DECISIONES VIGENTES (aprobadas por Bernardo)
- Tema visual ORIGINAL: navy #0A1929 / dorado #F5A623, gráficas familia azul/dorado
  (#1F3A68, #2E5C9A, #C49B3A…). NO reaplicar la paleta del sitio web. Logos de Integra en
  portada y esquina de cada slide.
- Cifras: titulares $4.80M / tablas $4,799,584. Una idea por slide; periodo visible en cada cifra.
- Meta del trimestre = escenario OPTIMISTA del forecast publicado.
- Ventas con pérdida (2025) se cuentan y se marcan; no se excluyen.
- Recompra + recomendación se mide por "Canal de llegada" (Recurrente + Recomendado).
- INBOX = WhatsApp. Servicios: solo se corrigen errores de escritura; servicios distintos NO se
  agrupan (TRE = Tomografía de resistividad; GPR = Georradar).
- Una venta con varios sitios cuenta como 1 venta; en el mapa el monto se reparte entre sitios.
- Nombres de clientes SÍ van en la presentación pública. El Explorador se queda en la presentación.
- Hallazgos internos (etapas del pipeline, motivos de pérdida, cobranza) solo en el PDF.
- Snapshots de pipeline: snapshots/pipeline_AAAA-MM.json por corte; comparar existencias solo
  si el alcance es el mismo (p. ej. el de 2026-10 es solo lo creado en Q3).

PUBLICACIÓN (solo con aprobado final)
Commit + push a main; confirmar la URL; verificar que ghl/.env, ghl/raw/, data/*.xlsx,
data/ghl_mastermetrics_*.json, reporte/reporte.html, reporte/data_interno.js y el PDF NO estén
en el repo ni respondan en el sitio publicado.
