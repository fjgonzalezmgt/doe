#' @title Helpers OpenAI para DOE
NULL

sanitize_analysis_result <- function(analysis_result) {
  list(
    analysis_id = analysis_result$analysis_id,
    title = analysis_result$title,
    subtitle = analysis_result$subtitle,
    summary = analysis_result$summary,
    tables = lapply(analysis_result$tables, function(tbl) {
      if (is.null(tbl)) {
        return(NULL)
      }
      as.data.frame(tbl)
    })
  )
}

encode_image_data_url <- function(image_path) {
  require_package("base64enc", "codificar imagenes para OpenAI")

  ext <- tolower(tools::file_ext(image_path))
  mime_type <- switch(
    ext,
    png = "image/png",
    jpg = "image/jpeg",
    jpeg = "image/jpeg",
    webp = "image/webp",
    stop("Formato de imagen no soportado para OpenAI.", call. = FALSE)
  )

  paste0("data:", mime_type, ";base64,", base64enc::base64encode(image_path))
}

extract_response_text <- function(body) {
  if (!is.null(body$output_text) && is.character(body$output_text) && nzchar(body$output_text)) {
    return(body$output_text)
  }

  if (!is.null(body$output) && length(body$output) > 0) {
    text_chunks <- unlist(
      lapply(body$output, function(item) {
        if (!is.list(item) || !identical(item$type, "message") || is.null(item$content)) {
          return(character())
        }

        unlist(
          lapply(item$content, function(content_item) {
            if (
              is.list(content_item) &&
              identical(content_item$type, "output_text") &&
              !is.null(content_item$text)
            ) {
              content_item$text
            } else {
              character()
            }
          }),
          use.names = FALSE
        )
      }),
      use.names = FALSE
    )

    text_chunks <- text_chunks[nzchar(text_chunks)]
    if (length(text_chunks) > 0) {
      return(paste(text_chunks, collapse = "\n\n"))
    }
  }

  NULL
}

openai_interpret_analysis <- function(analysis_result, plot_path = NULL, language = "es", extra_instructions = "") {
  require_package("httr2", "consultar la API de OpenAI")
  require_package("jsonlite", "serializar resultados")

  api_key <- Sys.getenv("OPENAI_API_KEY", unset = "")
  model <- Sys.getenv("OPENAI_MODEL", unset = "gpt-5-mini")

  if (!nzchar(api_key)) {
    stop("OPENAI_API_KEY no esta configurada.", call. = FALSE)
  }

  prompt_text <- paste(
    "Interpreta el siguiente analisis DOE en", language, ".",
    "Devuelve solo este formato:",
    "1. Resumen ejecutivo",
    "2. Factores y efectos criticos",
    "3. Riesgos de escalamiento o de ejecucion",
    "4. Siguiente iteracion recomendada",
    "Se breve, claro y orientado a toma de decisiones industriales.",
    "Usa las tablas y el grafico, sin inventar datos faltantes.",
    if (nzchar(extra_instructions)) extra_instructions else "",
    "\n\nResultado:\n",
    jsonlite::toJSON(sanitize_analysis_result(analysis_result), auto_unbox = TRUE, pretty = TRUE, null = "null")
  )

  message_content <- list(list(type = "input_text", text = prompt_text))
  if (!is.null(plot_path) && file.exists(plot_path)) {
    message_content[[length(message_content) + 1]] <- list(
      type = "input_image",
      image_url = encode_image_data_url(plot_path),
      detail = "high"
    )
  }

  payload <- list(
    model = model,
    instructions = paste(
      "Actua como especialista senior en diseno de experimentos, escalamiento y mejora industrial.",
      "Evalua robustez del diseno, factores significativos y proximos pasos experimentales."
    ),
    input = list(
      list(
        role = "user",
        content = message_content
      )
    ),
    text = list(verbosity = "low")
  )

  response <- httr2::request("https://api.openai.com/v1/responses") |>
    httr2::req_headers(
      Authorization = paste("Bearer", api_key),
      "Content-Type" = "application/json"
    ) |>
    httr2::req_body_json(payload, auto_unbox = TRUE) |>
    httr2::req_perform()

  body <- httr2::resp_body_json(response, simplifyVector = FALSE)
  text <- extract_response_text(body)

  if (!is.null(text) && nzchar(text)) {
    return(text)
  }

  jsonlite::toJSON(body, auto_unbox = TRUE, pretty = TRUE, null = "null")
}
