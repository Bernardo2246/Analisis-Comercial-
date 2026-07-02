Actualiza el proyecto "Análisis Comercial Integra Cimentaciones" (repo GitHub Pages existente
en bernardo2246.github.io/Analisis-Comercial-) con datos hasta [MES/AÑO].

FUENTE PRINCIPAL (única que se comparte en cada actualización trimestral):
Un .xlsx con ventas jun 2025–[mes actual]. Columnas: ubicación, fecha solicitud, fecha
confirmación, días transcurridos, tipo de servicio, canal, tipo de cliente, costo, cliente,
abono, resta, %avance.

BASELINE DE FUNNEL (ya integrado al proyecto, snapshot congelado con corte al 1-jul-2026,
NO se vuelve a pedir cada trimestre a menos que se indique):
- 1,464 oportunidades Clientify (78.2% perdidas, 9.5% ganadas, 155 abiertas, 25 vencidas)
- Motivo de pérdida dominante: "Pérdida de contacto" (74% de las pérdidas)
- Cuello de botella real: oportunidades mueren en etapa "Oportunidad congelada" y "Lead In",
  no en cotización enviada
- Canal de origen: usar columna "tags" de contactos como fuente primaria (más completa que
  "origen del contacto", que solo está llena en 18% de los registros)

QUÉ HACER:
1. Recalcula TODAS las métricas del Excel de ventas (revenue, ticket promedio, ciclo de venta,
   geografía, servicios, canales) incluyendo el rango nuevo.
2. Mantén la estructura actual, EXCEPTO: omite diapositivas 12 (Cotización SRM) y 13
   (Plan de marketing).
3. Conserva la interactividad: hover en gráficas, tablas ordenables, Explorador con filtros,
   selector de escenarios en Forecast.
4. Cruza revenue cerrado (del Excel) contra funnel completo (del baseline CRM) para mostrar
   tasa de conversión real lead→cierre, no solo revenue de lo ganado.
5. Actualiza el forecast con los datos reales del trimestre que ya cerró (deja de ser
   proyección, es real).
6. Paleta de color (Integra Cimentaciones):
   - Fondo oscuro: #2A3540
   - Verde lima: #A8CB62
   - Verde acento: #96BC4B
   - Ámbar/CTA: #F5BD41
   - Texto sobre oscuro: #FFFFFF
7. Al terminar: commit + push al repo, confírmame la URL actualizada.

REUTILIZABLE: cada trimestre solo cambio el rango de fechas del punto 1 y adjunto el .xlsx
de ventas actualizado — el baseline de funnel se refresca solo si mando un nuevo export
de Clientify explícitamente.
