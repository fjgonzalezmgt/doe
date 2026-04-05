library(openxlsx)

plan_df <- data.frame(
  Temperatura = c(180, 220, 180, 220, 200),
  Presion = c(45, 45, 65, 65, 55),
  Catalizador = c("A", "A", "B", "B", "A"),
  std_order = 1:5,
  run_order = c(3, 1, 5, 2, 4),
  coded_Temperatura = c(-1, 1, -1, 1, 0),
  coded_Presion = c(-1, -1, 1, 1, 0),
  coded_Catalizador = c(-1, -1, 1, 1, -1),
  run_status = c("Completada", "Completada", "Completada", "Completada", "Completada"),
  batch_id = c("B2401", "B2401", "B2402", "B2402", "B2403"),
  operator_id = c("OP01", "OP02", "OP01", "OP03", "OP02"),
  timestamp = c(
    "2026-04-05 08:00:00",
    "2026-04-05 09:00:00",
    "2026-04-05 10:00:00",
    "2026-04-05 11:00:00",
    "2026-04-05 12:00:00"
  ),
  comments = c(
    "Corrida estable",
    "Ligera espuma al arranque",
    "Cambio de lote de materia prima",
    "Sin incidencias",
    "Punto centro para confirmar tendencia"
  ),
  Yield = c(81.2, 89.7, 84.5, 94.1, 88.0),
  DefectRate = c(5.6, 3.2, 4.8, 2.1, 3.7),
  stringsAsFactors = FALSE,
  check.names = FALSE
)

factors_df <- data.frame(
  Factor = c("Temperatura", "Presion", "Catalizador"),
  Column = c("Temperatura", "Presion", "Catalizador"),
  Type = c("Numerico", "Numerico", "Categorico"),
  Low = c("180", "45", "A"),
  High = c("220", "65", "B"),
  stringsAsFactors = FALSE,
  check.names = FALSE
)

summary_df <- data.frame(
  Metrica = c(
    "Familia",
    "Corridas",
    "Factores",
    "Factores_numericos",
    "Factores_categoricos",
    "Respuestas",
    "Aleatorizado",
    "Semilla",
    "Motor",
    "Nota"
  ),
  Valor = c(
    "Factorial completo 2 niveles",
    "5",
    "3",
    "2",
    "1",
    "Yield, DefectRate",
    "Si",
    "123",
    "Base R",
    "Ejemplo DOE para validar el flujo completo de la aplicacion."
  ),
  stringsAsFactors = FALSE,
  check.names = FALSE
)

instructions_df <- data.frame(
  Paso = c(
    "1",
    "2",
    "3",
    "4"
  ),
  Instruccion = c(
    "Carga la hoja 'Plan_corridas' como archivo de ejecucion en la app.",
    "Analiza primero la respuesta 'Yield'.",
    "Analiza despues la respuesta 'DefectRate'.",
    "Usa este archivo para validar generacion de tablas, grafico y exportacion."
  ),
  stringsAsFactors = FALSE,
  check.names = FALSE
)

wb <- createWorkbook()

addWorksheet(wb, "Resumen_plan")
writeData(wb, "Resumen_plan", summary_df)

addWorksheet(wb, "Factores")
writeData(wb, "Factores", factors_df)

addWorksheet(wb, "Plan_corridas")
writeData(wb, "Plan_corridas", plan_df)

addWorksheet(wb, "Instrucciones")
writeData(wb, "Instrucciones", instructions_df)

setColWidths(wb, "Resumen_plan", cols = 1:2, widths = "auto")
setColWidths(wb, "Factores", cols = 1:ncol(factors_df), widths = "auto")
setColWidths(wb, "Plan_corridas", cols = 1:ncol(plan_df), widths = "auto")
setColWidths(wb, "Instrucciones", cols = 1:2, widths = c(10, 100))

header_style <- createStyle(textDecoration = "bold", fgFill = "#DCE6F1")
for (sheet in c("Resumen_plan", "Factores", "Plan_corridas", "Instrucciones")) {
  addStyle(wb, sheet, header_style, rows = 1, cols = 1:50, gridExpand = TRUE)
}

saveWorkbook(wb, "doe_example_workbook.xlsx", overwrite = TRUE)
