#' @title Motor DOE para la app Shiny
#' @description Funciones auxiliares para definir factores, generar planes DOE,
#'   analizar respuestas y exportar workbooks de resultados.
#' @keywords internal
NULL

#' @title Devolver valor por defecto
#' @description Retorna un valor alterno cuando la entrada es `NULL` o cadena vacia.
#' @param x Valor evaluado.
#' @param default Valor a retornar si `x` no contiene un dato util.
#' @return El valor original o el valor por defecto.
#' @keywords internal
or_default <- function(x, default) {
  if (is.null(x) || identical(x, "")) {
    default
  } else {
    x
  }
}

#' @title Validar escalar faltante
#' @description Indica si un objeto no contiene un valor escalar util.
#' @param x Valor a validar.
#' @return `TRUE` cuando el valor esta ausente o es `NA`.
#' @keywords internal
is_missing_scalar <- function(x) {
  length(x) < 1 || is.null(x) || (length(x) == 1 && is.na(x))
}

#' @title Verificar disponibilidad de paquete
#' @description Comprueba si un paquete de R puede cargarse mediante `requireNamespace()`.
#' @param pkg Nombre del paquete.
#' @return `TRUE` si el paquete esta instalado.
#' @keywords internal
package_available <- function(pkg) {
  requireNamespace(pkg, quietly = TRUE)
}

#' @title Exigir paquete instalado
#' @description Detiene la ejecucion si falta un paquete necesario para una tarea.
#' @param pkg Nombre del paquete requerido.
#' @param reason Motivo por el que el paquete es necesario.
#' @return Invisiblemente `NULL`; produce error si el paquete no esta disponible.
#' @keywords internal
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

#' @title Leer datos delimitados
#' @description Lee archivos tabulares delimitados conservando nombres de columna originales.
#' @param path Ruta del archivo.
#' @param header Indica si la primera fila contiene encabezados.
#' @param sep Separador de campos.
#' @param quote Caracter de comillas.
#' @param dec Separador decimal.
#' @return `data.frame` con los datos importados.
#' @keywords internal
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

#' @title Leer datos de entrada
#' @description Lee archivos CSV, TXT o Excel usados por la app para ejecucion y analisis.
#' @param path Ruta del archivo.
#' @param header Indica si la primera fila contiene encabezados.
#' @param sep Separador de campos para archivos delimitados.
#' @param quote Caracter de comillas para archivos delimitados.
#' @param dec Separador decimal para archivos delimitados.
#' @param sheet Hoja de Excel a leer cuando aplica.
#' @return `data.frame` con los datos cargados.
#' @keywords internal
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

#' @title Opciones de familias DOE
#' @description Construye el vector nombrado usado por la UI para seleccionar la familia de diseno.
#' @return Vector nombrado con ids internos y etiquetas visibles.
#' @keywords internal
design_family_choices <- function() {
  stats::setNames(names(design_family_catalog), vapply(
    design_family_catalog,
    function(x) sprintf("%s (%s)", x$label, x$planner),
    character(1)
  ))
}

#' @title Estado de paquetes recomendados
#' @description Genera una tabla con disponibilidad y uso de los paquetes de la app.
#' @return `data.frame` con nombre del paquete, estado y uso.
#' @keywords internal
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

#' @title Sanitizar nombres
#' @description Convierte etiquetas arbitrarias en nombres seguros y unicos para columnas.
#' @param x Vector de nombres originales.
#' @return Vector de nombres saneados.
#' @keywords internal
sanitize_name <- function(x) {
  cleaned <- gsub("[^A-Za-z0-9_]+", "_", trimws(x))
  cleaned <- gsub("^_+|_+$", "", cleaned)
  cleaned <- ifelse(nzchar(cleaned), cleaned, "X")
  make.unique(cleaned, sep = "_")
}

#' @title Parsear definiciones de factores
#' @description Convierte el texto ingresado por el usuario en una tabla de factores DOE.
#' @param text Texto multilinea con formato `nombre, nivel_bajo, nivel_alto`.
#' @return `data.frame` con metadatos de factores, niveles y codificacion base.
#' @keywords internal
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

#' @title Parsear nombres de respuestas
#' @description Extrae y sanea los nombres de respuestas esperadas a partir de texto libre.
#' @param text Texto con respuestas separadas por coma, salto de linea, tab o punto y coma.
#' @return Vector de nombres de respuesta.
#' @keywords internal
parse_response_names <- function(text) {
  responses <- trimws(unlist(strsplit(or_default(text, ""), "[,;\n\t]")))
  responses <- responses[nzchar(responses)]
  if (length(responses) < 1) {
    return(character())
  }
  sanitize_name(responses)
}

#' @title Columnas numericas
#' @description Obtiene los nombres de las columnas numericas de un `data.frame`.
#' @param df Tabla de datos.
#' @return Vector de nombres de columna.
#' @keywords internal
numeric_column_names <- function(df) {
  names(df)[vapply(df, is.numeric, logical(1))]
}

#' @title Valores para selector de columnas
#' @description Genera etiquetas con nombre y clase para usarlas en controles `selectInput`.
#' @param df Tabla de datos.
#' @return Vector nombrado apto para opciones de seleccion.
#' @keywords internal
column_choice_values <- function(df) {
  cols <- names(df)
  labels <- vapply(
    cols,
    function(col) sprintf("%s (%s)", col, class(df[[col]])[[1]]),
    character(1)
  )
  stats::setNames(cols, labels)
}

#' @title Convertir lista nombrada a tabla
#' @description Transforma una lista simple en una tabla de metricas y valores.
#' @param x Lista nombrada.
#' @return `data.frame` con columnas `Metrica` y `Valor`.
#' @keywords internal
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

#' @title Sanitizar nombre de hoja
#' @description Elimina caracteres invalidos y limita la longitud para nombres de hojas Excel.
#' @param x Nombre propuesto de hoja.
#' @return Cadena compatible con Excel.
#' @keywords internal
sanitize_sheet_name <- function(x) {
  cleaned <- gsub("[\\\\/:*?\\[\\]]", "-", x)
  substr(cleaned, 1, 31)
}

#' @title Rellenar valor vacio
#' @description Sustituye valores vacios por una etiqueta de ausencia.
#' @param value Valor a validar.
#' @param missing_label Texto a usar cuando el valor esta vacio.
#' @return Cadena con el valor original o la etiqueta de ausencia.
#' @keywords internal
non_empty_value <- function(value, missing_label = "No definido") {
  if (is.null(value) || identical(value, "") || identical(value, "_none")) {
    missing_label
  } else {
    value
  }
}

#' @title Exportar grafico a archivo
#' @description Guarda un objeto `ggplot` en disco con dimensiones predefinidas.
#' @param plot_obj Objeto grafico.
#' @param path Ruta destino.
#' @param width Ancho en pulgadas.
#' @param height Alto en pulgadas.
#' @param dpi Resolucion de exportacion.
#' @return Invisiblemente la ruta generada.
#' @keywords internal
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

#' @title Construir instrucciones de ejecucion
#' @description Genera la tabla de pasos operativos incluida en el workbook de ejecucion.
#' @param plan_result Lista con el plan DOE generado.
#' @return `data.frame` con pasos e instrucciones.
#' @keywords internal
build_execution_instruction_table <- function(plan_result) {
  response_text <- if (length(plan_result$responses) > 0) {
    paste(plan_result$responses, collapse = ", ")
  } else {
    "No definidas"
  }

  data.frame(
    Paso = as.character(seq_len(5)),
    Instruccion = c(
      "Completa la hoja Plan_corridas con los resultados reales de cada corrida.",
      "Mantiene intactas las columnas del plan y llena run_status, batch_id, operator_id, timestamp y comments.",
      sprintf("Registra las respuestas en las columnas esperadas: %s.", response_text),
      "Guarda el archivo y cargalo luego en la seccion Ejecucion de la app.",
      "Selecciona la hoja Plan_corridas y pulsa Analizar DOE."
    ),
    stringsAsFactors = FALSE,
    check.names = FALSE
  )
}

#' @title Decodificar valores numericos
#' @description Convierte niveles codificados a valores reales usando centro y paso del factor.
#' @param coded_values Vector de valores codificados.
#' @param factor_row Fila de metadatos del factor.
#' @return Vector numerico en unidades reales.
#' @keywords internal
decode_numeric_values <- function(coded_values, factor_row) {
  factor_row$Center + coded_values * factor_row$Step
}

#' @title Reconstruir niveles reales
#' @description Convierte una matriz codificada en un plan con niveles reales por factor.
#' @param coded_df Tabla de valores codificados.
#' @param factors Tabla de factores.
#' @return `data.frame` con valores reales de corrida.
#' @keywords internal
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

#' @title Agregar columnas codificadas
#' @description Anexa al plan las columnas `coded_*` usadas para analisis posterior.
#' @param plan_df Plan en unidades reales.
#' @param coded_df Plan codificado.
#' @param factors Tabla de factores.
#' @return `data.frame` actualizado.
#' @keywords internal
append_coded_columns <- function(plan_df, coded_df, factors) {
  for (index in seq_len(nrow(factors))) {
    factor_row <- factors[index, ]
    plan_df[[paste0("coded_", factor_row$Column)]] <- as.numeric(coded_df[[factor_row$Column]])
  }

  plan_df
}

#' @title Agregar columnas de ejecucion
#' @description Incorpora campos operativos y columnas de respuesta vacias al plan.
#' @param plan_df Plan de corridas.
#' @param responses Vector de respuestas esperadas.
#' @return `data.frame` listo para captura de ejecucion.
#' @keywords internal
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

#' @title Finalizar plan DOE
#' @description Completa ordenes de corrida, columnas operativas y resumen del plan.
#' @param plan_df Plan en unidades reales.
#' @param coded_df Plan codificado.
#' @param factors Tabla de factores.
#' @param responses Vector de respuestas esperadas.
#' @param design_family Identificador de la familia DOE.
#' @param randomize Indica si se aleatoriza el orden de corrida.
#' @param seed Semilla opcional para aleatorizacion.
#' @param notes Nota descriptiva del plan.
#' @return Lista estructurada con plan, factores, resumen y tablas.
#' @keywords internal
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

#' @title Generar factorial completo
#' @description Construye un plan factorial completo de dos niveles, con replicas y puntos centro opcionales.
#' @param factors Tabla de factores.
#' @param responses Vector de respuestas esperadas.
#' @param randomize Indica si se aleatoriza el orden.
#' @param seed Semilla opcional.
#' @param replicates Numero de replicas por combinacion.
#' @param center_points Numero de puntos centro.
#' @return Lista con el plan DOE finalizado.
#' @keywords internal
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

#' @title Extraer codigos de diseno
#' @description Convierte la salida de paquetes DOE en una tabla numerica de columnas codificadas.
#' @param design_df Tabla devuelta por el generador de diseno.
#' @param factor_columns Nombres de columnas correspondientes a factores.
#' @return `data.frame` con valores codificados numericos.
#' @keywords internal
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

#' @title Generar factorial fraccional
#' @description Construye un plan fraccional de dos niveles mediante `FrF2`.
#' @param factors Tabla de factores.
#' @param responses Vector de respuestas esperadas.
#' @param randomize Indica si se aleatoriza el orden.
#' @param seed Semilla opcional.
#' @param runs Numero de corridas.
#' @return Lista con el plan DOE finalizado.
#' @keywords internal
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

#' @title Generar diseno Plackett-Burman
#' @description Construye un plan de screening usando `FrF2::pb()`.
#' @param factors Tabla de factores.
#' @param responses Vector de respuestas esperadas.
#' @param randomize Indica si se aleatoriza el orden.
#' @param seed Semilla opcional.
#' @param runs Numero de corridas.
#' @return Lista con el plan DOE finalizado.
#' @keywords internal
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

#' @title Generar CCD
#' @description Construye un Central Composite Design para factores numericos.
#' @param factors Tabla de factores.
#' @param responses Vector de respuestas esperadas.
#' @param randomize Indica si se aleatoriza el orden.
#' @param seed Semilla opcional.
#' @param n0_cube Numero de puntos centro en la porcion factorial.
#' @param n0_star Numero de puntos centro en la porcion axial.
#' @param alpha Configuracion de alpha para `rsm::ccd()`.
#' @return Lista con el plan DOE finalizado.
#' @keywords internal
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

#' @title Generar Box-Behnken
#' @description Construye un diseno Box-Behnken para modelado cuadratico.
#' @param factors Tabla de factores.
#' @param responses Vector de respuestas esperadas.
#' @param randomize Indica si se aleatoriza el orden.
#' @param seed Semilla opcional.
#' @param center_points Numero de puntos centro.
#' @return Lista con el plan DOE finalizado.
#' @keywords internal
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

#' @title Generar plan DOE
#' @description Despacha la generacion del plan a la familia DOE seleccionada.
#' @param design_family Identificador de la familia DOE.
#' @param factors Tabla de factores.
#' @param responses Vector de respuestas esperadas.
#' @param randomize Indica si se aleatoriza el orden.
#' @param seed Semilla opcional.
#' @param control_values Lista de parametros especificos de la familia.
#' @return Lista con el plan DOE finalizado.
#' @keywords internal
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

#' @title Preparar dataset para analisis
#' @description Devuelve los datos subidos por el usuario o, si no existen, el plan actual.
#' @param plan_result Lista con el plan DOE.
#' @param uploaded_data Datos de ejecucion cargados opcionalmente.
#' @return `data.frame` o `NULL`.
#' @keywords internal
coerce_analysis_dataset <- function(plan_result, uploaded_data = NULL) {
  if (!is.null(uploaded_data)) {
    return(as.data.frame(uploaded_data, check.names = FALSE))
  }

  if (is.null(plan_result)) {
    return(NULL)
  }

  as.data.frame(plan_result$plan, check.names = FALSE)
}

#' @title Recomendar columnas de respuesta
#' @description Identifica columnas numericas candidatas a respuesta excluyendo factores y metadatos.
#' @param df Tabla de datos.
#' @param factors Tabla de factores.
#' @return Vector de nombres de columnas recomendadas.
#' @keywords internal
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

#' @title Validar columnas para analisis
#' @description Comprueba que existan factores y respuesta, y que haya suficientes corridas observadas.
#' @param df Tabla de datos de ejecucion.
#' @param factors Tabla de factores.
#' @param response_col Nombre de la respuesta a modelar.
#' @return `data.frame` filtrado a corridas con respuesta observada.
#' @keywords internal
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

#' @title Asegurar columnas codificadas
#' @description Crea columnas `coded_*` faltantes a partir de niveles reales.
#' @param df Tabla de datos.
#' @param factors Tabla de factores.
#' @return `data.frame` con columnas codificadas disponibles.
#' @keywords internal
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

#' @title Construir formula de efectos
#' @description Genera una formula lineal con efectos principales e interacciones de segundo orden opcionales.
#' @param response_col Nombre de la respuesta.
#' @param predictors Vector de predictores.
#' @param include_interactions Indica si deben incluirse interacciones de primer orden.
#' @return Objeto `formula`.
#' @keywords internal
build_effect_formula <- function(response_col, predictors, include_interactions = TRUE) {
  base_terms <- paste(predictors, collapse = " + ")
  if (length(predictors) > 1 && isTRUE(include_interactions)) {
    stats::as.formula(sprintf("%s ~ (%s)^2", response_col, base_terms))
  } else {
    stats::as.formula(sprintf("%s ~ %s", response_col, base_terms))
  }
}

#' @title Construir formula RSM
#' @description Genera una formula cuadratica para superficies de respuesta.
#' @param response_col Nombre de la respuesta.
#' @param predictors Vector de predictores codificados.
#' @return Objeto `formula`.
#' @keywords internal
build_rsm_formula <- function(response_col, predictors) {
  if (length(predictors) == 1) {
    return(stats::as.formula(sprintf("%s ~ %s + I(%s^2)", response_col, predictors, predictors)))
  }

  main_and_twi <- sprintf("(%s)^2", paste(predictors, collapse = " + "))
  quadratic <- paste(sprintf("I(%s^2)", predictors), collapse = " + ")
  stats::as.formula(sprintf("%s ~ %s + %s", response_col, main_and_twi, quadratic))
}

#' @title Grafico de coeficientes
#' @description Construye un grafico de barras con los coeficientes estimados del modelo.
#' @param coef_df Tabla de coeficientes.
#' @param title Titulo del grafico.
#' @return Objeto `ggplot`.
#' @keywords internal
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

#' @title Grafico observado vs ajustado
#' @description Dibuja la comparacion entre respuesta observada y valor ajustado.
#' @param fit_df Tabla con observados, ajustados y residuales.
#' @param response_col Nombre de la respuesta.
#' @return Objeto `ggplot`.
#' @keywords internal
build_observed_vs_fitted_plot <- function(fit_df, response_col) {
  ggplot2::ggplot(
    fit_df,
    ggplot2::aes(x = Ajustado, y = Observado)
  ) +
    ggplot2::geom_point(size = 2.8, color = "#1d4ed8", alpha = 0.85) +
    ggplot2::geom_abline(intercept = 0, slope = 1, linetype = "dashed", color = "#6b7280") +
    ggplot2::labs(
      title = sprintf("Observado vs ajustado: %s", response_col),
      x = "Valor ajustado",
      y = "Valor observado"
    ) +
    ggplot2::theme_minimal(base_size = 12)
}

#' @title Grafico residuales vs ajustados
#' @description Dibuja residuales frente a valores ajustados para revisar patrones.
#' @param fit_df Tabla con observados, ajustados y residuales.
#' @param response_col Nombre de la respuesta.
#' @return Objeto `ggplot`.
#' @keywords internal
build_residuals_vs_fitted_plot <- function(fit_df, response_col) {
  ggplot2::ggplot(
    fit_df,
    ggplot2::aes(x = Ajustado, y = Residual)
  ) +
    ggplot2::geom_hline(yintercept = 0, linetype = "dashed", color = "#6b7280") +
    ggplot2::geom_point(size = 2.8, color = "#b45309", alpha = 0.85) +
    ggplot2::labs(
      title = sprintf("Residuales vs ajustados: %s", response_col),
      x = "Valor ajustado",
      y = "Residual"
    ) +
    ggplot2::theme_minimal(base_size = 12)
}

#' @title Grafico residuales por orden
#' @description Dibuja residuales contra el orden de corrida para detectar deriva temporal.
#' @param fit_df Tabla con observados, ajustados y residuales.
#' @param response_col Nombre de la respuesta.
#' @return Objeto `ggplot`.
#' @keywords internal
build_residuals_run_order_plot <- function(fit_df, response_col) {
  ggplot2::ggplot(
    fit_df,
    ggplot2::aes(x = run_order, y = Residual)
  ) +
    ggplot2::geom_hline(yintercept = 0, linetype = "dashed", color = "#6b7280") +
    ggplot2::geom_line(color = "#0f766e", linewidth = 0.7) +
    ggplot2::geom_point(size = 2.5, color = "#0f766e") +
    ggplot2::labs(
      title = sprintf("Residuales por orden de corrida: %s", response_col),
      x = "Orden de corrida",
      y = "Residual"
    ) +
    ggplot2::theme_minimal(base_size = 12)
}

#' @title QQ plot de residuales
#' @description Dibuja un grafico QQ para revisar normalidad aproximada de los residuales.
#' @param fit_df Tabla con observados, ajustados y residuales.
#' @param response_col Nombre de la respuesta.
#' @return Objeto `ggplot`.
#' @keywords internal
build_qq_residuals_plot <- function(fit_df, response_col) {
  ggplot2::ggplot(
    fit_df,
    ggplot2::aes(sample = Residual)
  ) +
    ggplot2::stat_qq(color = "#7c3aed", size = 2.2, alpha = 0.85) +
    ggplot2::stat_qq_line(color = "#111827", linetype = "dashed") +
    ggplot2::labs(
      title = sprintf("QQ plot de residuales: %s", response_col),
      x = "Cuantiles teoricos",
      y = "Cuantiles muestrales"
    ) +
    ggplot2::theme_minimal(base_size = 12)
}

#' @title Ejecutar analisis DOE
#' @description Ajusta un modelo sobre los datos ejecutados y genera tablas y graficos diagnosticos.
#' @param plan_result Lista con el plan DOE.
#' @param execution_df Tabla con datos de ejecucion.
#' @param response_col Nombre de la respuesta a analizar.
#' @return Lista estructurada con resumen, tablas, graficos y log del modelo.
#' @keywords internal
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

  fit_df <- data.frame(
    run_order = if ("run_order" %in% names(working_df)) working_df$run_order else seq_len(nrow(working_df)),
    Observado = working_df[[response_col]],
    Ajustado = stats::fitted(fit),
    Residual = stats::residuals(fit),
    check.names = FALSE
  )

  plots <- list(
    `Efectos estimados` = build_coefficient_plot(coef_df, sprintf("Efectos estimados: %s", response_col)),
    `Observado vs ajustado` = build_observed_vs_fitted_plot(fit_df, response_col),
    `Residuales vs ajustados` = build_residuals_vs_fitted_plot(fit_df, response_col),
    `Residuales por corrida` = build_residuals_run_order_plot(fit_df, response_col),
    `QQ plot residuales` = build_qq_residuals_plot(fit_df, response_col)
  )
  plot_obj <- plots[[1]]

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
      `Corridas analizadas` = working_df,
      `Ajustados y residuales` = fit_df
    ),
    plots = plots,
    plot_obj = plot_obj,
    build_plot = function(path) {
      build_plot_export(plot_obj, path = path)
    },
    log = paste(capture.output(print(fit_summary)), collapse = "\n")
  )
}

#' @title Exportar workbook DOE
#' @description Escribe un workbook Excel con plan, analisis, graficos e interpretacion.
#' @param path Ruta del archivo a crear.
#' @param plan_result Lista con el plan DOE.
#' @param analysis_result Resultado del analisis, opcional.
#' @param interpretation_text Texto de interpretacion asistida, opcional.
#' @return Invisiblemente la ruta exportada.
#' @keywords internal
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

    openxlsx::addWorksheet(wb, "Graficos")
    plot_list <- analysis_result$plots
    if (is.null(plot_list) || length(plot_list) < 1) {
      plot_list <- list(Principal = analysis_result$plot_obj)
    }

    current_row <- 2
    for (plot_name in names(plot_list)) {
      plot_path <- tempfile(fileext = ".png")
      build_plot_export(plot_list[[plot_name]], plot_path, width = 10, height = 6)
      openxlsx::writeData(wb, "Graficos", x = plot_name, startRow = current_row, startCol = 2)
      openxlsx::addStyle(wb, "Graficos", title_style, rows = current_row, cols = 2, gridExpand = FALSE, stack = TRUE)
      openxlsx::insertImage(
        wb,
        "Graficos",
        plot_path,
        startRow = current_row + 1,
        startCol = 2,
        width = 10,
        height = 6,
        units = "in"
      )
      current_row <- current_row + 32
    }
    openxlsx::setColWidths(wb, "Graficos", cols = 1:3, widths = c(3, 18, 3))
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

#' @title Exportar workbook de ejecucion
#' @description Escribe una plantilla Excel para capturar resultados de corrida en planta o laboratorio.
#' @param path Ruta del archivo a crear.
#' @param plan_result Lista con el plan DOE.
#' @return Invisiblemente la ruta exportada.
#' @keywords internal
write_doe_execution_workbook <- function(path, plan_result) {
  require_package("openxlsx", "exportar resultados a Excel")

  wb <- openxlsx::createWorkbook()
  title_style <- openxlsx::createStyle(textDecoration = "bold", fgFill = "#DCE6F1")

  summary_df <- data_frame_from_named_list(plan_result$summary)
  factors_df <- plan_result$tables$Factores
  plan_df <- plan_result$plan
  instructions_df <- build_execution_instruction_table(plan_result)

  openxlsx::addWorksheet(wb, "Resumen_plan")
  openxlsx::writeData(wb, "Resumen_plan", x = summary_df, rowNames = FALSE)

  openxlsx::addWorksheet(wb, "Factores")
  openxlsx::writeData(wb, "Factores", x = factors_df, rowNames = FALSE)

  openxlsx::addWorksheet(wb, "Plan_corridas")
  openxlsx::writeData(wb, "Plan_corridas", x = plan_df, rowNames = FALSE)

  openxlsx::addWorksheet(wb, "Instrucciones")
  openxlsx::writeData(wb, "Instrucciones", x = instructions_df, rowNames = FALSE)

  for (sheet_name in c("Resumen_plan", "Factores", "Plan_corridas", "Instrucciones")) {
    sheet_df <- switch(
      sheet_name,
      Resumen_plan = summary_df,
      Factores = factors_df,
      Plan_corridas = plan_df,
      Instrucciones = instructions_df
    )
    openxlsx::addStyle(wb, sheet_name, title_style, rows = 1, cols = seq_len(max(1, ncol(sheet_df))), gridExpand = TRUE)
  }

  openxlsx::setColWidths(wb, "Resumen_plan", cols = 1:2, widths = "auto")
  openxlsx::setColWidths(wb, "Factores", cols = seq_len(ncol(factors_df)), widths = "auto")
  openxlsx::setColWidths(wb, "Plan_corridas", cols = seq_len(ncol(plan_df)), widths = "auto")
  openxlsx::setColWidths(wb, "Instrucciones", cols = 1:2, widths = c(10, 110))

  openxlsx::saveWorkbook(wb, path, overwrite = TRUE)
  invisible(path)
}
