#' @title Motor DOE para la app Shiny
NULL

or_default <- function(x, default) {
  if (is.null(x) || identical(x, "")) {
    default
  } else {
    x
  }
}

is_missing_scalar <- function(x) {
  length(x) < 1 || is.null(x) || (length(x) == 1 && is.na(x))
}

package_available <- function(pkg) {
  requireNamespace(pkg, quietly = TRUE)
}

require_package <- function(pkg, reason) {
  if (!package_available(pkg)) {
    stop(
      sprintf(
        "Falta el paquete '%s' para %s. Instalar con install.packages('%s').",
        pkg,
        reason,
        pkg
      ),
      call. = FALSE
    )
  }
}

read_delimited_data <- function(path, header, sep, quote, dec) {
  utils::read.table(
    file = path,
    header = header,
    sep = sep,
    quote = quote,
    dec = dec,
    stringsAsFactors = FALSE,
    check.names = FALSE
  )
}

read_input_data <- function(path, header, sep, quote, dec, sheet = NULL) {
  ext <- tolower(tools::file_ext(path))

  if (ext %in% c("xls", "xlsx")) {
    require_package("readxl", "leer archivos Excel")
    return(as.data.frame(readxl::read_excel(path = path, sheet = sheet)))
  }

  read_delimited_data(
    path = path,
    header = header,
    sep = sep,
    quote = quote,
    dec = dec
  )
}

design_family_catalog <- list(
  full_factorial = list(
    label = "Factorial completo 2 niveles",
    planner = "Base R",
    objective = "Caracterizar efectos principales e interacciones de primer orden.",
    notes = "Apto para factores numericos o categoricos de dos niveles."
  ),
  fractional_factorial = list(
    label = "Factorial fraccional 2 niveles",
    planner = "FrF2",
    objective = "Screening industrial con menos corridas manteniendo estructura formal de alias.",
    notes = "Requiere FrF2."
  ),
  screening_pb = list(
    label = "Plackett-Burman",
    planner = "FrF2",
    objective = "Identificar rapido factores vitales con foco en efectos principales.",
    notes = "Requiere FrF2."
  ),
  ccd = list(
    label = "Central Composite Design",
    planner = "rsm",
    objective = "Construir modelos cuadraticos para optimizacion.",
    notes = "Requiere rsm y factores numericos."
  ),
  box_behnken = list(
    label = "Box-Behnken",
    planner = "rsm",
    objective = "Ajustar superficies de respuesta con menos corridas que un CCD.",
    notes = "Requiere rsm y factores numericos."
  )
)

design_family_choices <- function() {
  stats::setNames(names(design_family_catalog), vapply(
    design_family_catalog,
    function(x) sprintf("%s (%s)", x$label, x$planner),
    character(1)
  ))
}

package_status_table <- function() {
  data.frame(
    Paquete = c("shiny", "bslib", "ggplot2", "readxl", "openxlsx", "FrF2", "DoE.base", "rsm"),
    Estado = vapply(
      c("shiny", "bslib", "ggplot2", "readxl", "openxlsx", "FrF2", "DoE.base", "rsm"),
      function(pkg) if (package_available(pkg)) "Disponible" else "No instalado",
      character(1)
    ),
    Uso = c(
      "Interfaz",
      "UI moderna",
      "Graficos",
      "Lectura Excel",
      "Exportacion Excel",
      "Fraccionales y screening",
      "Base DOE",
      "Superficie de respuesta"
    ),
    row.names = NULL,
    check.names = FALSE
  )
}

sanitize_name <- function(x) {
  cleaned <- gsub("[^A-Za-z0-9_]+", "_", trimws(x))
  cleaned <- gsub("^_+|_+$", "", cleaned)
  cleaned <- ifelse(nzchar(cleaned), cleaned, "X")
  make.unique(cleaned, sep = "_")
}

parse_factor_definitions <- function(text) {
  lines <- trimws(unlist(strsplit(or_default(text, ""), "\n", fixed = TRUE)))
  lines <- lines[nzchar(lines)]

  if (length(lines) < 1) {
    stop(
      "Define al menos un factor. Usa: nombre, nivel_bajo, nivel_alto.",
      call. = FALSE
    )
  }

  parsed <- lapply(lines, function(line) {
    parts <- trimws(unlist(strsplit(line, "[,;\t]")))
    if (length(parts) < 3) {
      stop(sprintf("Linea invalida para factor: %s", line), call. = FALSE)
    }

    low_num <- suppressWarnings(as.numeric(parts[[2]]))
    high_num <- suppressWarnings(as.numeric(parts[[3]]))
    is_numeric <- !is.na(low_num) && !is.na(high_num)

    data.frame(
      Factor = parts[[1]],
      Low = if (is_numeric) as.character(low_num) else parts[[2]],
      High = if (is_numeric) as.character(high_num) else parts[[3]],
      Type = if (is_numeric) "Numerico" else "Categorico",
      stringsAsFactors = FALSE,
      check.names = FALSE
    )
  })

  factors_df <- do.call(rbind, parsed)
  factors_df$Column <- sanitize_name(factors_df$Factor)
  factors_df$Center <- ifelse(
    factors_df$Type == "Numerico",
    (as.numeric(factors_df$Low) + as.numeric(factors_df$High)) / 2,
    NA_real_
  )
  factors_df$Step <- ifelse(
    factors_df$Type == "Numerico",
    (as.numeric(factors_df$High) - as.numeric(factors_df$Low)) / 2,
    NA_real_
  )

  if (any(as.numeric(factors_df$Step[factors_df$Type == "Numerico"]) <= 0)) {
    stop("En factores numericos, el nivel alto debe ser mayor que el nivel bajo.", call. = FALSE)
  }

  factors_df
}

parse_response_names <- function(text) {
  responses <- trimws(unlist(strsplit(or_default(text, ""), "[,;\n\t]")))
  responses <- responses[nzchar(responses)]
  if (length(responses) < 1) {
    return(character())
  }
  sanitize_name(responses)
}

numeric_column_names <- function(df) {
  names(df)[vapply(df, is.numeric, logical(1))]
}

column_choice_values <- function(df) {
  cols <- names(df)
  labels <- vapply(
    cols,
    function(col) sprintf("%s (%s)", col, class(df[[col]])[[1]]),
    character(1)
  )
  stats::setNames(cols, labels)
}

data_frame_from_named_list <- function(x) {
  if (length(x) < 1) {
    return(data.frame())
  }

  data.frame(
    Metrica = names(x),
    Valor = vapply(x, function(value) {
      if (length(value) > 1) {
        paste(format(value), collapse = ", ")
      } else {
        as.character(value)
      }
    }, character(1)),
    row.names = NULL,
    check.names = FALSE
  )
}

sanitize_sheet_name <- function(x) {
  cleaned <- gsub("[\\\\/:*?\\[\\]]", "-", x)
  substr(cleaned, 1, 31)
}

non_empty_value <- function(value, missing_label = "No definido") {
  if (is.null(value) || identical(value, "") || identical(value, "_none")) {
    missing_label
  } else {
    value
  }
}

build_plot_export <- function(plot_obj, path, width = 11, height = 7, dpi = 180) {
  ggplot2::ggsave(
    filename = path,
    plot = plot_obj,
    width = width,
    height = height,
    dpi = dpi,
    units = "in"
  )
  invisible(path)
}

decode_numeric_values <- function(coded_values, factor_row) {
  factor_row$Center + coded_values * factor_row$Step
}

build_actual_from_coded <- function(coded_df, factors) {
  actual_list <- vector("list", nrow(factors))
  names(actual_list) <- factors$Column

  for (index in seq_len(nrow(factors))) {
    factor_row <- factors[index, ]
    coded_values <- coded_df[[factor_row$Column]]

    actual_list[[factor_row$Column]] <- if (identical(factor_row$Type, "Numerico")) {
      decode_numeric_values(as.numeric(coded_values), factor_row)
    } else {
      ifelse(as.numeric(coded_values) <= 0, factor_row$Low, factor_row$High)
    }
  }

  as.data.frame(actual_list, check.names = FALSE)
}

append_coded_columns <- function(plan_df, coded_df, factors) {
  for (index in seq_len(nrow(factors))) {
    factor_row <- factors[index, ]
    plan_df[[paste0("coded_", factor_row$Column)]] <- as.numeric(coded_df[[factor_row$Column]])
  }

  plan_df
}

add_execution_columns <- function(plan_df, responses) {
  plan_df$run_status <- "Pendiente"
  plan_df$batch_id <- NA_character_
  plan_df$operator_id <- NA_character_
  plan_df$timestamp <- NA_character_
  plan_df$comments <- NA_character_

  for (response in responses) {
    plan_df[[response]] <- NA_real_
  }

  plan_df
}

finalize_plan <- function(plan_df, coded_df, factors, responses, design_family, randomize, seed, notes = NULL) {
  plan_df$std_order <- seq_len(nrow(plan_df))

  if (isTRUE(randomize)) {
    if (!is_missing_scalar(seed)) {
      set.seed(as.integer(seed))
    }
    plan_df$run_order <- sample(seq_len(nrow(plan_df)))
  } else {
    plan_df$run_order <- seq_len(nrow(plan_df))
  }

  plan_df <- append_coded_columns(plan_df, coded_df, factors)
  plan_df <- add_execution_columns(plan_df, responses)
  plan_df <- plan_df[order(plan_df$run_order), , drop = FALSE]
  row.names(plan_df) <- NULL

  list(
    design_family = design_family,
    plan = plan_df,
    factors = factors,
    responses = responses,
    summary = list(
      Familia = design_family_catalog[[design_family]]$label,
      Corridas = nrow(plan_df),
      Factores = nrow(factors),
      Factores_numericos = sum(factors$Type == "Numerico"),
      Factores_categoricos = sum(factors$Type == "Categorico"),
      Respuestas = if (length(responses) > 0) paste(responses, collapse = ", ") else "No definidas",
      Aleatorizado = if (isTRUE(randomize)) "Si" else "No",
      Semilla = if (is_missing_scalar(seed)) "No definida" else as.character(seed),
      Motor = design_family_catalog[[design_family]]$planner,
      Nota = non_empty_value(notes)
    ),
    tables = list(
      Factores = factors[, c("Factor", "Column", "Type", "Low", "High"), drop = FALSE],
      `Plan de corridas` = plan_df
    )
  )
}

build_full_factorial_plan <- function(factors, responses, randomize, seed, replicates = 1, center_points = 0) {
  level_grid <- rep(list(c(-1, 1)), nrow(factors))
  names(level_grid) <- factors$Column

  coded_df <- expand.grid(level_grid, KEEP.OUT.ATTRS = FALSE, stringsAsFactors = FALSE)
  coded_df <- coded_df[rep(seq_len(nrow(coded_df)), each = max(1, replicates)), , drop = FALSE]

  if (center_points > 0) {
    if (!all(factors$Type == "Numerico")) {
      stop("Los puntos centro solo se habilitan para factores numericos.", call. = FALSE)
    }
    center_df <- as.data.frame(matrix(0, nrow = center_points, ncol = nrow(factors)))
    names(center_df) <- factors$Column
    coded_df <- rbind(coded_df, center_df)
  }

  plan_df <- build_actual_from_coded(coded_df, factors)
  finalize_plan(
    plan_df = plan_df,
    coded_df = coded_df,
    factors = factors,
    responses = responses,
    design_family = "full_factorial",
    randomize = randomize,
    seed = seed,
    notes = "Plan completo de dos niveles con opcion de replicas y puntos centro."
  )
}

extract_design_codes <- function(design_df, factor_columns) {
  coded_df <- data.frame(check.names = FALSE)

  for (column in factor_columns) {
    values <- design_df[[column]]
    if (is.factor(values)) {
      values <- as.character(values)
    }
    coded_df[[column]] <- as.numeric(values)
  }

  coded_df
}

build_fractional_factorial_plan <- function(factors, responses, randomize, seed, runs) {
  require_package("FrF2", "generar factoriales fraccionales")

  if (nrow(factors) < 3) {
    stop("El factorial fraccional tiene sentido a partir de 3 factores.", call. = FALSE)
  }

  design_obj <- FrF2::FrF2(
    nruns = as.integer(runs),
    nfactors = nrow(factors),
    factor.names = factors$Column,
    randomize = FALSE
  )

  coded_df <- extract_design_codes(as.data.frame(design_obj), factors$Column)
  plan_df <- build_actual_from_coded(coded_df, factors)

  finalize_plan(
    plan_df = plan_df,
    coded_df = coded_df,
    factors = factors,
    responses = responses,
    design_family = "fractional_factorial",
    randomize = randomize,
    seed = seed,
    notes = sprintf("Factorial fraccional con %s corridas generado por FrF2.", runs)
  )
}

build_screening_plan <- function(factors, responses, randomize, seed, runs) {
  require_package("FrF2", "generar disenos Plackett-Burman")

  design_obj <- FrF2::pb(
    nruns = as.integer(runs),
    nfactors = nrow(factors),
    factor.names = factors$Column,
    randomize = FALSE
  )

  coded_df <- extract_design_codes(as.data.frame(design_obj), factors$Column)
  plan_df <- build_actual_from_coded(coded_df, factors)

  finalize_plan(
    plan_df = plan_df,
    coded_df = coded_df,
    factors = factors,
    responses = responses,
    design_family = "screening_pb",
    randomize = randomize,
    seed = seed,
    notes = sprintf("Screening Plackett-Burman con %s corridas.", runs)
  )
}

build_ccd_plan <- function(factors, responses, randomize, seed, n0_cube = 4, n0_star = 4, alpha = "rotatable") {
  require_package("rsm", "generar un Central Composite Design")

  if (!all(factors$Type == "Numerico")) {
    stop("CCD solo admite factores numericos.", call. = FALSE)
  }

  basis_formula <- stats::reformulate(factors$Column)
  coded_design <- rsm::ccd(
    basis = basis_formula,
    n0 = c(as.integer(n0_cube), as.integer(n0_star)),
    alpha = alpha,
    randomize = FALSE,
    oneblock = TRUE
  )

  coded_df <- as.data.frame(coded_design[, factors$Column, drop = FALSE], check.names = FALSE)
  plan_df <- build_actual_from_coded(coded_df, factors)

  finalize_plan(
    plan_df = plan_df,
    coded_df = coded_df,
    factors = factors,
    responses = responses,
    design_family = "ccd",
    randomize = randomize,
    seed = seed,
    notes = sprintf("CCD con alpha = %s. Los niveles bajo/alto definen la region cubica base.", alpha)
  )
}

build_bbd_plan <- function(factors, responses, randomize, seed, center_points = 4) {
  require_package("rsm", "generar un Box-Behnken")

  if (!all(factors$Type == "Numerico")) {
    stop("Box-Behnken solo admite factores numericos.", call. = FALSE)
  }

  if (nrow(factors) < 3 || nrow(factors) > 7) {
    stop("Box-Behnken requiere entre 3 y 7 factores numericos.", call. = FALSE)
  }

  basis_formula <- stats::reformulate(factors$Column)
  coded_design <- rsm::bbd(
    k = basis_formula,
    n0 = as.integer(center_points),
    block = FALSE,
    randomize = FALSE
  )

  coded_df <- as.data.frame(coded_design[, factors$Column, drop = FALSE], check.names = FALSE)
  plan_df <- build_actual_from_coded(coded_df, factors)

  finalize_plan(
    plan_df = plan_df,
    coded_df = coded_df,
    factors = factors,
    responses = responses,
    design_family = "box_behnken",
    randomize = randomize,
    seed = seed,
    notes = "Diseno Box-Behnken para modelado cuadratico."
  )
}

generate_doe_plan <- function(design_family, factors, responses, randomize, seed, control_values) {
  switch(
    design_family,
    full_factorial = build_full_factorial_plan(
      factors = factors,
      responses = responses,
      randomize = randomize,
      seed = seed,
      replicates = control_values$replicates,
      center_points = control_values$center_points
    ),
    fractional_factorial = build_fractional_factorial_plan(
      factors = factors,
      responses = responses,
      randomize = randomize,
      seed = seed,
      runs = control_values$runs
    ),
    screening_pb = build_screening_plan(
      factors = factors,
      responses = responses,
      randomize = randomize,
      seed = seed,
      runs = control_values$runs
    ),
    ccd = build_ccd_plan(
      factors = factors,
      responses = responses,
      randomize = randomize,
      seed = seed,
      n0_cube = control_values$n0_cube,
      n0_star = control_values$n0_star,
      alpha = control_values$alpha
    ),
    box_behnken = build_bbd_plan(
      factors = factors,
      responses = responses,
      randomize = randomize,
      seed = seed,
      center_points = control_values$center_points
    ),
    stop(sprintf("Familia DOE no soportada: %s", design_family), call. = FALSE)
  )
}

coerce_analysis_dataset <- function(plan_result, uploaded_data = NULL) {
  if (!is.null(uploaded_data)) {
    return(as.data.frame(uploaded_data, check.names = FALSE))
  }

  if (is.null(plan_result)) {
    return(NULL)
  }

  as.data.frame(plan_result$plan, check.names = FALSE)
}

recommended_response_columns <- function(df, factors) {
  if (is.null(df)) {
    return(character())
  }

  excluded <- c(
    factors$Column,
    paste0("coded_", factors$Column),
    "std_order",
    "run_order"
  )

  numeric_cols <- numeric_column_names(df)
  setdiff(numeric_cols, excluded)
}

ensure_analysis_columns <- function(df, factors, response_col) {
  required <- c(factors$Column, response_col)
  missing_cols <- setdiff(required, names(df))

  if (length(missing_cols) > 0) {
    stop(
      sprintf("Faltan columnas en los datos de ejecucion: %s", paste(missing_cols, collapse = ", ")),
      call. = FALSE
    )
  }

  if (!is.numeric(df[[response_col]])) {
    stop("La respuesta seleccionada debe ser numerica.", call. = FALSE)
  }

  visible_df <- df[!is.na(df[[response_col]]), , drop = FALSE]
  if (nrow(visible_df) < max(4, nrow(factors) + 1)) {
    stop("No hay suficientes corridas con respuesta medida para ajustar el modelo.", call. = FALSE)
  }

  visible_df
}

ensure_coded_columns <- function(df, factors) {
  for (index in seq_len(nrow(factors))) {
    factor_row <- factors[index, ]
    coded_name <- paste0("coded_", factor_row$Column)

    if (coded_name %in% names(df)) {
      next
    }

    if (identical(factor_row$Type, "Numerico")) {
      df[[coded_name]] <- (as.numeric(df[[factor_row$Column]]) - factor_row$Center) / factor_row$Step
    } else {
      df[[coded_name]] <- ifelse(as.character(df[[factor_row$Column]]) == factor_row$Low, -1, 1)
    }
  }

  df
}

build_effect_formula <- function(response_col, predictors, include_interactions = TRUE) {
  base_terms <- paste(predictors, collapse = " + ")
  if (length(predictors) > 1 && isTRUE(include_interactions)) {
    stats::as.formula(sprintf("%s ~ (%s)^2", response_col, base_terms))
  } else {
    stats::as.formula(sprintf("%s ~ %s", response_col, base_terms))
  }
}

build_rsm_formula <- function(response_col, predictors) {
  if (length(predictors) == 1) {
    return(stats::as.formula(sprintf("%s ~ %s + I(%s^2)", response_col, predictors, predictors)))
  }

  main_and_twi <- sprintf("(%s)^2", paste(predictors, collapse = " + "))
  quadratic <- paste(sprintf("I(%s^2)", predictors), collapse = " + ")
  stats::as.formula(sprintf("%s ~ %s + %s", response_col, main_and_twi, quadratic))
}

build_coefficient_plot <- function(coef_df, title) {
  plot_df <- coef_df[coef_df$Termino != "(Intercept)", , drop = FALSE]

  if (nrow(plot_df) < 1) {
    plot_df <- data.frame(Termino = "Sin terminos", Estimate = 0, check.names = FALSE)
  }

  ggplot2::ggplot(
    plot_df,
    ggplot2::aes(x = stats::reorder(Termino, abs(Estimate)), y = Estimate, fill = Estimate > 0)
  ) +
    ggplot2::geom_col(show.legend = FALSE) +
    ggplot2::coord_flip() +
    ggplot2::scale_fill_manual(values = c("#b91c1c", "#0f766e")) +
    ggplot2::labs(title = title, x = "Termino", y = "Coeficiente") +
    ggplot2::theme_minimal(base_size = 12)
}

run_doe_analysis <- function(plan_result, execution_df, response_col) {
  factors <- plan_result$factors
  working_df <- ensure_analysis_columns(execution_df, factors, response_col)
  working_df <- ensure_coded_columns(working_df, factors)
  predictor_cols <- paste0("coded_", factors$Column)

  model_formula <- if (plan_result$design_family %in% c("ccd", "box_behnken")) {
    build_rsm_formula(response_col, predictor_cols)
  } else {
    build_effect_formula(
      response_col = response_col,
      predictors = predictor_cols,
      include_interactions = nrow(factors) <= 6
    )
  }

  fit <- stats::lm(model_formula, data = working_df)
  fit_summary <- summary(fit)

  anova_df <- as.data.frame(stats::anova(fit), check.names = FALSE)
  anova_df$Termino <- row.names(anova_df)
  row.names(anova_df) <- NULL
  anova_df <- anova_df[, c("Termino", setdiff(names(anova_df), "Termino")), drop = FALSE]

  coef_df <- as.data.frame(fit_summary$coefficients, check.names = FALSE)
  coef_df$Termino <- row.names(coef_df)
  row.names(coef_df) <- NULL
  coef_df <- coef_df[, c("Termino", setdiff(names(coef_df), "Termino")), drop = FALSE]

  plot_obj <- build_coefficient_plot(coef_df, sprintf("Efectos estimados: %s", response_col))

  list(
    analysis_id = "doe_analysis",
    title = if (plan_result$design_family %in% c("ccd", "box_behnken")) {
      "Modelo DOE de segundo orden"
    } else {
      "Modelo DOE de efectos"
    },
    subtitle = sprintf("%s sobre %s", design_family_catalog[[plan_result$design_family]]$label, response_col),
    summary = list(
      Familia = design_family_catalog[[plan_result$design_family]]$label,
      Respuesta = response_col,
      Corridas_analizadas = nrow(working_df),
      Factores = nrow(factors),
      Formula = paste(deparse(model_formula), collapse = " "),
      R2 = round(or_default(fit_summary$r.squared, NA_real_), 4),
      R2_ajustado = round(or_default(fit_summary$adj.r.squared, NA_real_), 4),
      RMSE = round(sqrt(mean(stats::residuals(fit)^2, na.rm = TRUE)), 4)
    ),
    tables = list(
      ANOVA = anova_df,
      Coeficientes = coef_df,
      `Corridas analizadas` = working_df
    ),
    plot_obj = plot_obj,
    build_plot = function(path) {
      build_plot_export(plot_obj, path = path)
    },
    log = paste(capture.output(print(fit_summary)), collapse = "\n")
  )
}

write_doe_export_workbook <- function(path, plan_result, analysis_result = NULL, interpretation_text = NULL) {
  require_package("openxlsx", "exportar resultados a Excel")

  wb <- openxlsx::createWorkbook()
  title_style <- openxlsx::createStyle(textDecoration = "bold", fgFill = "#DCE6F1")
  body_style <- openxlsx::createStyle(valign = "top", wrapText = TRUE)

  openxlsx::addWorksheet(wb, "Resumen plan")
  plan_summary_df <- data_frame_from_named_list(plan_result$summary)
  openxlsx::writeData(wb, "Resumen plan", x = plan_summary_df, rowNames = FALSE)
  openxlsx::addStyle(wb, "Resumen plan", title_style, rows = 1, cols = 1:2, gridExpand = TRUE)
  openxlsx::setColWidths(wb, "Resumen plan", cols = 1:2, widths = "auto")

  openxlsx::addWorksheet(wb, "Factores")
  openxlsx::writeData(wb, "Factores", x = plan_result$tables$Factores, rowNames = FALSE)
  openxlsx::addStyle(wb, "Factores", title_style, rows = 1, cols = seq_len(ncol(plan_result$tables$Factores)), gridExpand = TRUE)
  openxlsx::setColWidths(wb, "Factores", cols = seq_len(ncol(plan_result$tables$Factores)), widths = "auto")

  openxlsx::addWorksheet(wb, "Plan corridas")
  openxlsx::writeData(wb, "Plan corridas", x = plan_result$plan, rowNames = FALSE)
  openxlsx::addStyle(wb, "Plan corridas", title_style, rows = 1, cols = seq_len(ncol(plan_result$plan)), gridExpand = TRUE)
  openxlsx::setColWidths(wb, "Plan corridas", cols = seq_len(ncol(plan_result$plan)), widths = "auto")

  if (!is.null(analysis_result)) {
    openxlsx::addWorksheet(wb, "Resumen analisis")
    analysis_summary_df <- data_frame_from_named_list(analysis_result$summary)
    openxlsx::writeData(wb, "Resumen analisis", x = analysis_summary_df, rowNames = FALSE)
    openxlsx::addStyle(wb, "Resumen analisis", title_style, rows = 1, cols = 1:2, gridExpand = TRUE)
    openxlsx::setColWidths(wb, "Resumen analisis", cols = 1:2, widths = "auto")

    for (sheet_name in names(analysis_result$tables)) {
      target_sheet <- sanitize_sheet_name(paste("Analisis", sheet_name))
      table_df <- analysis_result$tables[[sheet_name]]
      openxlsx::addWorksheet(wb, target_sheet)
      openxlsx::writeData(wb, target_sheet, x = table_df, rowNames = FALSE)
      openxlsx::addStyle(wb, target_sheet, title_style, rows = 1, cols = seq_len(max(1, ncol(table_df))), gridExpand = TRUE)
      openxlsx::setColWidths(wb, target_sheet, cols = seq_len(max(1, ncol(table_df))), widths = "auto")
    }

    openxlsx::addWorksheet(wb, "Grafico")
    plot_path <- tempfile(fileext = ".png")
    analysis_result$build_plot(plot_path)
    openxlsx::insertImage(wb, "Grafico", plot_path, startRow = 2, startCol = 2, width = 10, height = 7, units = "in")
  }

  openxlsx::addWorksheet(wb, "Interpretacion")
  interpretation_value <- if (is.null(interpretation_text) || !nzchar(trimws(interpretation_text))) {
    "No se genero interpretacion."
  } else {
    interpretation_text
  }
  openxlsx::writeData(wb, "Interpretacion", x = data.frame(Interpretacion = interpretation_value, check.names = FALSE), rowNames = FALSE)
  openxlsx::addStyle(wb, "Interpretacion", body_style, rows = 2, cols = 1, gridExpand = TRUE, stack = TRUE)
  openxlsx::setColWidths(wb, "Interpretacion", cols = 1, widths = 120)
  openxlsx::setRowHeights(wb, "Interpretacion", rows = 2, heights = 100)

  openxlsx::saveWorkbook(wb, path, overwrite = TRUE)
  invisible(path)
}
