library(openxlsx)

clamp_values <- function(x, lower, upper) {
  pmax(lower, pmin(upper, x))
}

build_workbook <- function(path, summary_df, factors_df, plan_df, instructions_df) {
  wb <- createWorkbook()

  addWorksheet(wb, "Resumen_plan")
  writeData(wb, "Resumen_plan", summary_df)

  addWorksheet(wb, "Factores")
  writeData(wb, "Factores", factors_df)

  addWorksheet(wb, "Plan_corridas")
  writeData(wb, "Plan_corridas", plan_df)

  addWorksheet(wb, "Instrucciones")
  writeData(wb, "Instrucciones", instructions_df)

  header_style <- createStyle(textDecoration = "bold", fgFill = "#DCE6F1")
  for (sheet in c("Resumen_plan", "Factores", "Plan_corridas", "Instrucciones")) {
    addStyle(wb, sheet, header_style, rows = 1, cols = 1:50, gridExpand = TRUE)
  }

  setColWidths(wb, "Resumen_plan", cols = 1:2, widths = "auto")
  setColWidths(wb, "Factores", cols = 1:ncol(factors_df), widths = "auto")
  setColWidths(wb, "Plan_corridas", cols = 1:ncol(plan_df), widths = "auto")
  setColWidths(wb, "Instrucciones", cols = 1:2, widths = c(10, 110))

  saveWorkbook(wb, path, overwrite = TRUE)
}

build_summary <- function(family, runs, factors_n, numeric_n, categorical_n, responses, motor, note) {
  data.frame(
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
      family,
      as.character(runs),
      as.character(factors_n),
      as.character(numeric_n),
      as.character(categorical_n),
      responses,
      "Si",
      "123",
      motor,
      note
    ),
    stringsAsFactors = FALSE,
    check.names = FALSE
  )
}

build_instructions <- function(family) {
  data.frame(
    Paso = c("1", "2", "3", "4"),
    Instruccion = c(
      sprintf("Carga este workbook en la app para probar el flujo %s.", family),
      "Selecciona la hoja Plan_corridas en la seccion de Ejecucion.",
      "Analiza primero la respuesta Yield.",
      "Analiza despues la respuesta secundaria indicada en el workbook."
    ),
    stringsAsFactors = FALSE,
    check.names = FALSE
  )
}

stamp_runs <- function(plan_df, prefix, start_date, comments) {
  n <- nrow(plan_df)
  plan_df$run_status <- rep("Completada", n)
  plan_df$batch_id <- sprintf("%s%04d", prefix, seq_len(n) + 2400)
  plan_df$operator_id <- rep(c("OP01", "OP02", "OP03", "OP04"), length.out = n)
  plan_df$timestamp <- sprintf("%s %02d:00:00", start_date, seq(8, length.out = n))
  plan_df$comments <- comments
  plan_df
}

simulate_full_factorial <- function(plan_df) {
  t <- plan_df$coded_Temperatura
  p <- plan_df$coded_Presion
  c <- plan_df$coded_Catalizador
  noise_y <- c(-0.4, 0.8, -0.6, 0.5, -0.9, 0.7, -0.3, 0.6, 0.1, -0.2)
  noise_d <- c(0.20, -0.15, 0.18, -0.10, 0.24, -0.18, 0.11, -0.09, 0.05, -0.04)

  plan_df$Yield <- round(clamp_values(
    86.5 + 4.6 * t + 2.5 * p + 3.1 * c + 1.2 * t * p - 0.8 * p * c + noise_y,
    72,
    98
  ), 1)
  plan_df$DefectRate <- round(clamp_values(
    4.6 - 0.9 * t - 0.7 * p - 0.8 * c - 0.3 * t * p + 0.2 * p * c + noise_d,
    1.2,
    7.5
  ), 2)
  plan_df
}

simulate_fractional <- function(plan_df) {
  t <- plan_df$coded_Temperatura
  p <- plan_df$coded_Presion
  ti <- plan_df$coded_Tiempo
  c <- plan_df$coded_Catalizador
  noise_y <- c(-0.6, 0.4, -0.3, 0.5, -0.5, 0.6, -0.2, 0.4)
  noise_d <- c(0.22, -0.16, 0.12, -0.14, 0.18, -0.20, 0.10, -0.12)

  plan_df$Yield <- round(clamp_values(
    84.0 + 3.4 * t + 1.4 * p + 2.8 * ti + 2.1 * c + 0.9 * t * c + noise_y,
    72,
    96
  ), 1)
  plan_df$DefectRate <- round(clamp_values(
    4.8 - 0.7 * t - 0.4 * p - 0.8 * ti - 0.6 * c - 0.2 * t * c + noise_d,
    1.5,
    7.0
  ), 2)
  plan_df
}

simulate_plackett_burman <- function(plan_df) {
  t <- plan_df$coded_Temperatura
  p <- plan_df$coded_Presion
  ti <- plan_df$coded_Tiempo
  a <- plan_df$coded_Agitacion
  ph <- plan_df$coded_pH
  noise_y <- c(-0.5, 0.7, -0.4, 0.3, -0.2, 0.4, -0.6, 0.5)
  noise_d <- c(0.18, -0.12, 0.15, -0.08, 0.10, -0.14, 0.16, -0.11)

  plan_df$Yield <- round(clamp_values(
    82.5 + 2.9 * t + 1.1 * p + 1.7 * ti + 3.3 * a + 2.5 * ph + noise_y,
    70,
    95
  ), 1)
  plan_df$DefectRate <- round(clamp_values(
    5.2 - 0.6 * t - 0.3 * p - 0.4 * ti - 0.9 * a - 0.7 * ph + noise_d,
    1.8,
    7.2
  ), 2)
  plan_df
}

simulate_ccd <- function(plan_df) {
  t <- plan_df$coded_Temperatura
  p <- plan_df$coded_Presion
  ti <- plan_df$coded_Tiempo
  noise_y <- c(-0.4, 0.6, -0.5, 0.4, -0.3, 0.5, -0.2, 0.4, -0.1, 0.2, -0.3, 0.3, 0.1, -0.2)
  noise_p <- c(-0.10, 0.12, -0.08, 0.10, -0.06, 0.08, -0.04, 0.07, -0.05, 0.06, -0.02, 0.03, 0.01, -0.01)

  plan_df$Yield <- round(clamp_values(
    88.0 + 3.2 * t + 1.8 * p + 2.1 * ti + 1.0 * t * p - 1.8 * (t^2) - 1.0 * (p^2) - 1.4 * (ti^2) + noise_y,
    76,
    97
  ), 1)
  plan_df$Purity <- round(clamp_values(
    95.2 + 0.8 * t + 0.5 * p + 0.7 * ti + 0.3 * t * ti - 0.5 * (t^2) - 0.4 * (p^2) - 0.6 * (ti^2) + noise_p,
    92.5,
    98.5
  ), 2)
  plan_df
}

simulate_bbd <- function(plan_df) {
  t <- plan_df$coded_Temperatura
  p <- plan_df$coded_Presion
  ti <- plan_df$coded_Tiempo
  noise_y <- c(-0.3, 0.5, -0.4, 0.4, -0.2, 0.3, -0.1, 0.2, -0.2, 0.3, 0.1, -0.1, 0.0)
  noise_p <- c(-0.08, 0.09, -0.06, 0.07, -0.04, 0.05, -0.03, 0.04, -0.02, 0.03, 0.01, -0.01, 0.00)

  plan_df$Yield <- round(clamp_values(
    87.4 + 2.8 * t + 1.6 * p + 1.9 * ti + 0.8 * t * ti + 0.6 * p * ti - 1.2 * (t^2) - 0.8 * (p^2) - 1.1 * (ti^2) + noise_y,
    77,
    96
  ), 1)
  plan_df$Purity <- round(clamp_values(
    95.0 + 0.7 * t + 0.4 * p + 0.6 * ti + 0.2 * t * p - 0.4 * (t^2) - 0.3 * (p^2) - 0.5 * (ti^2) + noise_p,
    93.0,
    98.0
  ), 2)
  plan_df
}

factorial_factors <- data.frame(
  Factor = c("Temperatura", "Presion", "Catalizador"),
  Column = c("Temperatura", "Presion", "Catalizador"),
  Type = c("Numerico", "Numerico", "Categorico"),
  Low = c("180", "45", "A"),
  High = c("220", "65", "B"),
  stringsAsFactors = FALSE,
  check.names = FALSE
)

factorial_signs <- expand.grid(
  coded_Temperatura = c(-1, 1),
  coded_Presion = c(-1, 1),
  coded_Catalizador = c(-1, 1),
  KEEP.OUT.ATTRS = FALSE,
  stringsAsFactors = FALSE
)
factorial_signs <- rbind(
  factorial_signs,
  factorial_signs[c(2, 7), , drop = FALSE]
)
factorial_plan <- data.frame(
  Temperatura = ifelse(factorial_signs$coded_Temperatura < 0, 180, 220),
  Presion = ifelse(factorial_signs$coded_Presion < 0, 45, 65),
  Catalizador = ifelse(factorial_signs$coded_Catalizador < 0, "A", "B"),
  std_order = seq_len(nrow(factorial_signs)),
  run_order = c(4, 9, 2, 7, 1, 10, 5, 8, 3, 6),
  coded_Temperatura = factorial_signs$coded_Temperatura,
  coded_Presion = factorial_signs$coded_Presion,
  coded_Catalizador = factorial_signs$coded_Catalizador,
  stringsAsFactors = FALSE,
  check.names = FALSE
)
factorial_plan <- stamp_runs(
  factorial_plan,
  prefix = "F",
  start_date = "2026-04-05",
  comments = c(
    "Arranque estable",
    "Replica para confirmar tendencia",
    "Ligera espuma al inicio",
    "Sin incidencias",
    "Cambio menor de lote",
    "Replica en condicion alta",
    "Ajuste fino de valvula",
    "Operacion nominal",
    "Verificacion de dosificacion",
    "Condicion estable"
  )
)
factorial_plan <- simulate_full_factorial(factorial_plan)

build_workbook(
  "doe_example_full_factorial.xlsx",
  build_summary(
    "Factorial completo 2 niveles",
    nrow(factorial_plan), 3, 2, 1,
    "Yield, DefectRate",
    "Base R",
    "Ejemplo con replicas para evitar ajuste saturado y mostrar efectos plausibles."
  ),
  factorial_factors,
  factorial_plan,
  build_instructions("Factorial completo 2 niveles")
)

frac_factors <- data.frame(
  Factor = c("Temperatura", "Presion", "Tiempo", "Catalizador"),
  Column = c("Temperatura", "Presion", "Tiempo", "Catalizador"),
  Type = c("Numerico", "Numerico", "Numerico", "Categorico"),
  Low = c("180", "45", "20", "A"),
  High = c("220", "65", "40", "B"),
  stringsAsFactors = FALSE,
  check.names = FALSE
)

frac_plan <- data.frame(
  Temperatura = c(180, 220, 180, 220, 180, 220, 180, 220),
  Presion = c(45, 45, 65, 65, 45, 45, 65, 65),
  Tiempo = c(20, 20, 20, 20, 40, 40, 40, 40),
  Catalizador = c("A", "B", "B", "A", "B", "A", "A", "B"),
  std_order = 1:8,
  run_order = c(5, 2, 7, 1, 6, 3, 8, 4),
  coded_Temperatura = c(-1, 1, -1, 1, -1, 1, -1, 1),
  coded_Presion = c(-1, -1, 1, 1, -1, -1, 1, 1),
  coded_Tiempo = c(-1, -1, -1, -1, 1, 1, 1, 1),
  coded_Catalizador = c(-1, 1, 1, -1, 1, -1, -1, 1),
  stringsAsFactors = FALSE,
  check.names = FALSE
)
frac_plan <- stamp_runs(
  frac_plan,
  prefix = "R",
  start_date = "2026-04-06",
  comments = c(
    "Arranque rapido",
    "Cambio de recipiente",
    "Sin observaciones",
    "Espuma controlada",
    "Estabilidad correcta",
    "Lote intermedio",
    "Revision visual conforme",
    "Operacion sin desvio"
  )
)
frac_plan <- simulate_fractional(frac_plan)

build_workbook(
  "doe_example_fractional_factorial.xlsx",
  build_summary(
    "Factorial fraccional 2 niveles",
    nrow(frac_plan), 4, 3, 1,
    "Yield, DefectRate",
    "FrF2",
    "Ejemplo de screening con efectos principales claros y ligera interaccion temperatura-catalizador."
  ),
  frac_factors,
  frac_plan,
  build_instructions("Factorial fraccional 2 niveles")
)

pb_factors <- data.frame(
  Factor = c("Temperatura", "Presion", "Tiempo", "Agitacion", "pH"),
  Column = c("Temperatura", "Presion", "Tiempo", "Agitacion", "pH"),
  Type = rep("Numerico", 5),
  Low = c("180", "45", "20", "200", "5.2"),
  High = c("220", "65", "40", "300", "6.0"),
  stringsAsFactors = FALSE,
  check.names = FALSE
)

pb_signs <- expand.grid(
  coded_Temperatura = c(-1, 1),
  coded_Presion = c(-1, 1),
  coded_Tiempo = c(-1, 1),
  KEEP.OUT.ATTRS = FALSE,
  stringsAsFactors = FALSE
)
pb_signs <- pb_signs[1:8, , drop = FALSE]
pb_plan <- data.frame(
  Temperatura = ifelse(pb_signs$coded_Temperatura < 0, 180, 220),
  Presion = ifelse(pb_signs$coded_Presion < 0, 45, 65),
  Tiempo = ifelse(pb_signs$coded_Tiempo < 0, 20, 40),
  Agitacion = c(200, 300, 300, 200, 300, 200, 200, 300),
  pH = c(5.2, 5.2, 6.0, 6.0, 6.0, 6.0, 5.2, 5.2),
  std_order = 1:8,
  run_order = c(4, 8, 2, 6, 1, 5, 3, 7),
  coded_Temperatura = pb_signs$coded_Temperatura,
  coded_Presion = pb_signs$coded_Presion,
  coded_Tiempo = pb_signs$coded_Tiempo,
  coded_Agitacion = c(-1, 1, 1, -1, 1, -1, -1, 1),
  coded_pH = c(-1, -1, 1, 1, 1, 1, -1, -1),
  stringsAsFactors = FALSE,
  check.names = FALSE
)
pb_plan <- stamp_runs(
  pb_plan,
  prefix = "P",
  start_date = "2026-04-07",
  comments = rep("Screening de factores", 8)
)
pb_plan <- simulate_plackett_burman(pb_plan)

build_workbook(
  "doe_example_plackett_burman.xlsx",
  build_summary(
    "Plackett-Burman",
    nrow(pb_plan), 5, 5, 0,
    "Yield, DefectRate",
    "FrF2",
    "Ejemplo de screening donde agitacion y pH son los factores mas influyentes."
  ),
  pb_factors,
  pb_plan,
  build_instructions("Plackett-Burman")
)

ccd_factors <- data.frame(
  Factor = c("Temperatura", "Presion", "Tiempo"),
  Column = c("Temperatura", "Presion", "Tiempo"),
  Type = rep("Numerico", 3),
  Low = c("180", "45", "20"),
  High = c("220", "65", "40"),
  stringsAsFactors = FALSE,
  check.names = FALSE
)

ccd_plan <- data.frame(
  Temperatura = c(180, 220, 180, 220, 180, 220, 180, 220, 200, 200, 200, 200, 200, 200),
  Presion = c(45, 45, 65, 65, 45, 45, 65, 65, 55, 55, 45, 65, 55, 55),
  Tiempo = c(20, 20, 20, 20, 40, 40, 40, 40, 20, 40, 30, 30, 30, 30),
  std_order = 1:14,
  run_order = c(9, 2, 11, 5, 13, 1, 6, 10, 3, 14, 8, 4, 12, 7),
  coded_Temperatura = c(-1, 1, -1, 1, -1, 1, -1, 1, 0, 0, 0, 0, 0, 0),
  coded_Presion = c(-1, -1, 1, 1, -1, -1, 1, 1, 0, 0, -1, 1, 0, 0),
  coded_Tiempo = c(-1, -1, -1, -1, 1, 1, 1, 1, -1, 1, 0, 0, 0, 0),
  stringsAsFactors = FALSE,
  check.names = FALSE
)
ccd_plan <- stamp_runs(
  ccd_plan,
  prefix = "C",
  start_date = "2026-04-08",
  comments = c(rep("Factorial", 8), rep("Axial o centro", 6))
)
ccd_plan <- simulate_ccd(ccd_plan)

build_workbook(
  "doe_example_ccd.xlsx",
  build_summary(
    "Central Composite Design",
    nrow(ccd_plan), 3, 3, 0,
    "Yield, Purity",
    "rsm",
    "Ejemplo CCD con curvatura moderada y maximo cercano al centro operativo."
  ),
  ccd_factors,
  ccd_plan,
  build_instructions("Central Composite Design")
)

bbd_factors <- data.frame(
  Factor = c("Temperatura", "Presion", "Tiempo"),
  Column = c("Temperatura", "Presion", "Tiempo"),
  Type = rep("Numerico", 3),
  Low = c("180", "45", "20"),
  High = c("220", "65", "40"),
  stringsAsFactors = FALSE,
  check.names = FALSE
)

bbd_plan <- data.frame(
  Temperatura = c(180, 220, 180, 220, 180, 220, 200, 200, 200, 200, 200, 200, 200),
  Presion = c(45, 45, 65, 65, 55, 55, 45, 45, 65, 65, 55, 55, 55),
  Tiempo = c(30, 30, 30, 30, 20, 20, 20, 40, 20, 40, 20, 40, 30),
  std_order = 1:13,
  run_order = c(4, 10, 1, 7, 12, 3, 8, 2, 11, 5, 13, 6, 9),
  coded_Temperatura = c(-1, 1, -1, 1, -1, 1, 0, 0, 0, 0, 0, 0, 0),
  coded_Presion = c(-1, -1, 1, 1, 0, 0, -1, -1, 1, 1, 0, 0, 0),
  coded_Tiempo = c(0, 0, 0, 0, -1, -1, -1, 1, -1, 1, -1, 1, 0),
  stringsAsFactors = FALSE,
  check.names = FALSE
)
bbd_plan <- stamp_runs(
  bbd_plan,
  prefix = "B",
  start_date = "2026-04-09",
  comments = rep("Superficie de respuesta", 13)
)
bbd_plan <- simulate_bbd(bbd_plan)

build_workbook(
  "doe_example_box_behnken.xlsx",
  build_summary(
    "Box-Behnken",
    nrow(bbd_plan), 3, 3, 0,
    "Yield, Purity",
    "rsm",
    "Ejemplo Box-Behnken con optimo suave alrededor del punto central."
  ),
  bbd_factors,
  bbd_plan,
  build_instructions("Box-Behnken")
)
