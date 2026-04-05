#' @title Server DOE workbench
NULL

server <- function(input, output, session) {
  analysis_state <- reactiveVal(NULL)

  workbook_plan_bundle <- reactive({
    req(input$execution_file)

    ext <- tolower(tools::file_ext(input$execution_file$name))
    if (!ext %in% c("xls", "xlsx")) {
      return(NULL)
    }

    if (!package_available("readxl")) {
      return(NULL)
    }

    sheets <- readxl::excel_sheets(input$execution_file$datapath)
    if (!all(c("Factores", "Plan_corridas") %in% sheets)) {
      return(NULL)
    }

    factors_df <- as.data.frame(readxl::read_excel(input$execution_file$datapath, sheet = "Factores"))
    plan_df <- as.data.frame(readxl::read_excel(input$execution_file$datapath, sheet = "Plan_corridas"))
    summary_df <- if ("Resumen_plan" %in% sheets) {
      as.data.frame(readxl::read_excel(input$execution_file$datapath, sheet = "Resumen_plan"))
    } else {
      NULL
    }

    family_label <- NULL
    if (!is.null(summary_df) && all(c("Metrica", "Valor") %in% names(summary_df))) {
      family_row <- summary_df[summary_df$Metrica == "Familia", , drop = FALSE]
      if (nrow(family_row) > 0) {
        family_label <- trimws(as.character(family_row$Valor[[1]]))
      }
    }

    design_family <- or_default(input$execution_design_family, "full_factorial")
    if (!is.null(family_label)) {
      matched_family <- names(Filter(function(meta) identical(trimws(meta$label), family_label), design_family_catalog))
      if (length(matched_family) > 0 && (is.null(input$execution_design_family) || !nzchar(input$execution_design_family))) {
        design_family <- matched_family[[1]]
      }
    }

    summary_list <- if (!is.null(summary_df) && all(c("Metrica", "Valor") %in% names(summary_df))) {
      as.list(stats::setNames(as.character(summary_df$Valor), summary_df$Metrica))
    } else {
      list(
        Familia = design_family_catalog[[design_family]]$label,
        Corridas = nrow(plan_df),
        Factores = nrow(factors_df)
      )
    }

    response_cols <- setdiff(
      names(plan_df)[vapply(plan_df, is.numeric, logical(1))],
      c(
        factors_df$Column,
        paste0("coded_", factors_df$Column),
        "std_order",
        "run_order"
      )
    )

    list(
      design_family = design_family,
      plan = plan_df,
      factors = factors_df,
      responses = response_cols,
      summary = summary_list,
      tables = list(
        Factores = factors_df,
        `Plan de corridas` = plan_df
      )
    )
  })

  output$design_family_help <- renderUI({
    meta <- design_family_catalog[[input$design_family]]
    req(meta)
    tagList(
      p(tags$strong(meta$label)),
      p(style = "color: #4b5563; margin-bottom: 0.25rem;", meta$objective),
      p(style = "font-size: 0.9rem; color: #6b7280;", meta$notes)
    )
  })

  output$execution_family_help <- renderUI({
    meta <- design_family_catalog[[input$execution_design_family]]
    req(meta)
    p(style = "font-size: 0.9rem; color: #6b7280; margin-bottom: 0.75rem;", meta$notes)
  })

  output$design_controls <- renderUI({
    switch(
      input$design_family,
      full_factorial = tagList(
        numericInput("full_replicates", "Replicas", value = 1, min = 1, step = 1),
        numericInput("full_center_points", "Puntos centro", value = 0, min = 0, step = 1)
      ),
      fractional_factorial = numericInput("frac_runs", "Corridas", value = 8, min = 4, step = 2),
      screening_pb = numericInput("pb_runs", "Corridas", value = 12, min = 8, step = 4),
      ccd = tagList(
        numericInput("ccd_n0_cube", "Centros en cubo", value = 4, min = 1, step = 1),
        numericInput("ccd_n0_star", "Centros axiales", value = 4, min = 1, step = 1),
        selectInput("ccd_alpha", "Alpha", choices = c("rotatable", "orthogonal", "faced"), selected = "rotatable")
      ),
      box_behnken = numericInput("bbd_center_points", "Puntos centro", value = 4, min = 1, step = 1)
    )
  })

  output$execution_sheet_control <- renderUI({
    req(input$execution_file)
    if (!tolower(tools::file_ext(input$execution_file$name)) %in% c("xls", "xlsx")) {
      return(NULL)
    }

    if (!package_available("readxl")) {
      return(p("Para leer Excel instala el paquete ", code("readxl"), "."))
    }

    sheets <- readxl::excel_sheets(input$execution_file$datapath)
    default_sheet <- if ("Plan_corridas" %in% sheets) "Plan_corridas" else sheets[[1]]
    selectInput("execution_sheet", "Hoja", choices = sheets, selected = default_sheet)
  })

  plan_attempt <- eventReactive(input$generate_plan, {
    tryCatch(
      {
        factors <- parse_factor_definitions(input$factor_definitions)
        responses <- parse_response_names(input$response_names)

        control_values <- switch(
          input$design_family,
          full_factorial = list(replicates = input$full_replicates, center_points = input$full_center_points),
          fractional_factorial = list(runs = input$frac_runs),
          screening_pb = list(runs = input$pb_runs),
          ccd = list(n0_cube = input$ccd_n0_cube, n0_star = input$ccd_n0_star, alpha = input$ccd_alpha),
          box_behnken = list(center_points = input$bbd_center_points)
        )

        list(
          ok = TRUE,
          plan = generate_doe_plan(
            design_family = input$design_family,
            factors = factors,
            responses = responses,
            randomize = isTRUE(input$randomize_runs),
            seed = input$random_seed,
            control_values = control_values
          )
        )
      },
      error = function(e) {
        list(ok = FALSE, message = conditionMessage(e))
      }
    )
  })

  current_plan <- reactive({
    if (isTruthy(input$generate_plan) && input$generate_plan > 0) {
      generated_plan <- plan_attempt()
      if (!is.null(generated_plan) && isTRUE(generated_plan$ok)) {
        return(generated_plan$plan)
      }
    }

    workbook_plan_bundle()
  })

  output$plan_status <- renderUI({
    if (isTruthy(input$generate_plan) && input$generate_plan > 0) {
      generated_plan <- plan_attempt()
      if (!is.null(generated_plan) && isTRUE(generated_plan$ok)) {
        return(p(style = "color: #0f766e;", "Plan DOE generado."))
      }
      if (!is.null(generated_plan) && !isTRUE(generated_plan$ok)) {
        return(p(style = "color: #b91c1c;", generated_plan$message))
      }
    }

    if (!is.null(workbook_plan_bundle())) {
      return(p(style = "color: #0f766e;", "Plan DOE detectado desde el workbook cargado."))
    }

    p("Define factores y genera el plan DOE, o carga un workbook DOE existente.")
  })

  output$plan_summary <- renderTable({
    req(current_plan())
    data_frame_from_named_list(current_plan()$summary)
  }, rownames = FALSE)

  output$factor_table <- renderTable({
    req(current_plan())
    current_plan()$tables$Factores
  }, rownames = FALSE)

  output$plan_table <- renderTable({
    req(current_plan())
    current_plan()$plan
  }, rownames = FALSE)

  execution_attempt <- reactive({
    if (is.null(input$execution_file)) {
      return(NULL)
    }

    tryCatch(
      {
        ext <- tolower(tools::file_ext(input$execution_file$name))
        selected_sheet <- if (ext %in% c("xls", "xlsx")) {
          available_sheets <- readxl::excel_sheets(input$execution_file$datapath)
          if (!is.null(input$execution_sheet) && nzchar(input$execution_sheet) && input$execution_sheet %in% available_sheets) {
            input$execution_sheet
          } else if ("Plan_corridas" %in% available_sheets) {
            "Plan_corridas"
          } else {
            available_sheets[[1]]
          }
        } else {
          NULL
        }

        list(
          ok = TRUE,
          data = read_input_data(
            path = input$execution_file$datapath,
            header = input$execution_header,
            sep = or_default(input$execution_sep, ","),
            quote = or_default(input$execution_quote, '"'),
            dec = or_default(input$execution_dec, "."),
            sheet = selected_sheet
          )
        )
      },
      error = function(e) {
        list(ok = FALSE, message = conditionMessage(e))
      }
    )
  })

  execution_data <- reactive({
    if (is.null(execution_attempt()) || !isTRUE(execution_attempt()$ok)) {
      return(NULL)
    }
    execution_attempt()$data
  })

  analysis_dataset <- reactive({
    if (is.null(current_plan())) {
      return(NULL)
    }
    coerce_analysis_dataset(current_plan(), execution_data())
  })

  output$execution_status <- renderUI({
    if (is.null(current_plan())) {
      return(p("Genera un plan DOE o carga un workbook DOE que contenga las hojas Factores y Plan_corridas."))
    }

    if (is.null(execution_attempt())) {
      return(p("Puedes descargar el workbook DOE, capturar resultados y luego cargar el archivo completado."))
    }

    if (!isTRUE(execution_attempt()$ok)) {
      return(p(style = "color: #b91c1c;", execution_attempt()$message))
    }

    p(sprintf("Archivo cargado con %s filas y %s columnas.", nrow(execution_data()), ncol(execution_data())))
  })

  output$execution_preview <- renderTable({
    if (is.null(execution_attempt()) || !isTRUE(execution_attempt()$ok) || is.null(execution_data())) {
      return(NULL)
    }
    utils::head(execution_data(), 15)
  }, rownames = FALSE)

  response_selector_ui <- reactive({
    tryCatch(
      {
        plan_obj <- current_plan()
        if (is.null(plan_obj)) {
          return(p("Carga un workbook DOE o genera un plan para habilitar la seleccion de respuesta."))
        }

        df <- execution_data()
        if (is.null(df)) {
          return(p("Carga datos de ejecucion para habilitar la seleccion de respuesta."))
        }

        response_choices <- recommended_response_columns(df, plan_obj$factors)
        if (length(response_choices) < 1) {
          return(p("No hay columnas numericas de respuesta disponibles todavia."))
        }

        selected_response <- if (!is.null(input$response_col) && input$response_col %in% response_choices) {
          input$response_col
        } else {
          response_choices[[1]]
        }

        selectInput("response_col", "Respuesta a analizar", choices = response_choices, selected = selected_response)
      },
      error = function(e) {
        p(style = "color: #b91c1c;", paste("Error en selector de respuesta:", conditionMessage(e)))
      }
    )
  })

  output$analysis_response_selector <- renderUI({
    response_selector_ui()
  })

  observe({
    if (is.null(current_plan()) || is.null(execution_data())) {
      return()
    }

    response_choices <- recommended_response_columns(execution_data(), current_plan()$factors)
    if (length(response_choices) < 1) {
      return()
    }

    selected_response <- if (!is.null(input$response_col) && input$response_col %in% response_choices) {
      input$response_col
    } else {
      response_choices[[1]]
    }

    updateSelectInput(session, "response_col", choices = response_choices, selected = selected_response)
  })

  current_analysis <- reactive({
    req(analysis_state(), isTRUE(analysis_state()$ok))
    analysis_state()$analysis
  })

  output$analysis_status <- renderUI({
    tryCatch(
      {
        plan_obj <- current_plan()
        if (is.null(plan_obj)) {
          return(p("Genera un plan DOE o carga un workbook DOE para habilitar el analisis."))
        }

        dataset <- analysis_dataset()
        if (is.null(dataset)) {
          return(p("Carga datos de ejecucion para habilitar el analisis."))
        }

        available_responses <- recommended_response_columns(dataset, plan_obj$factors)
        if (length(available_responses) < 1) {
          return(p(style = "color: #b91c1c;", "No hay respuestas numericas disponibles para analizar."))
        }

        if (is.null(analysis_state())) {
          selected_response <- if (!is.null(input$response_col) && input$response_col %in% available_responses) {
            input$response_col
          } else {
            available_responses[[1]]
          }
          return(p(sprintf("Listo para analizar. Respuesta actual: %s.", selected_response)))
        }

        if (isTRUE(analysis_state()$ok)) {
          return(p(style = "color: #0f766e;", "Analisis DOE completado."))
        }

        p(style = "color: #b91c1c;", analysis_state()$message)
      },
      error = function(e) {
        p(style = "color: #b91c1c;", paste("Error en estado de analisis:", conditionMessage(e)))
      }
    )
  })

  observeEvent(input$run_analysis, {
    if (is.null(current_plan()) || is.null(analysis_dataset())) {
      analysis_state(list(ok = FALSE, message = "Falta plan o datos de ejecucion para correr el analisis."))
      showNotification("Falta plan o datos de ejecucion para correr el analisis.", type = "error", duration = 8)
      return()
    }

    available_responses <- recommended_response_columns(analysis_dataset(), current_plan()$factors)
    selected_response <- if (!is.null(input$response_col) && nzchar(input$response_col) && input$response_col %in% available_responses) {
      input$response_col
    } else if (length(available_responses) > 0) {
      available_responses[[1]]
    } else {
      NULL
    }

    if (is.null(selected_response)) {
      analysis_state(list(ok = FALSE, message = "No hay respuestas numericas disponibles para analizar."))
      showNotification("No hay respuestas numericas disponibles para analizar.", type = "error")
      return()
    }

    result <- withProgress(message = "Analizando DOE", value = 0.3, {
      tryCatch(
        {
          analysis <- run_doe_analysis(
            plan_result = current_plan(),
            execution_df = analysis_dataset(),
            response_col = selected_response
          )
          incProgress(0.7)
          list(ok = TRUE, analysis = analysis)
        },
        error = function(e) {
          list(ok = FALSE, message = conditionMessage(e))
        }
      )
    })

    analysis_state(result)

    if (isTRUE(result$ok)) {
      try(updateTabsetPanel(session, "main_tabs", selected = "Analisis"), silent = TRUE)
      showNotification("Analisis DOE completado.", type = "message")
    } else {
      try(updateTabsetPanel(session, "main_tabs", selected = "Analisis"), silent = TRUE)
      showNotification(result$message, type = "error", duration = 8)
    }
  }, ignoreInit = TRUE)

  output$analysis_summary <- renderTable({
    if (is.null(analysis_state()) || !isTRUE(analysis_state()$ok)) {
      return(NULL)
    }
    data_frame_from_named_list(current_analysis()$summary)
  }, rownames = FALSE)

  output$analysis_log <- renderText({
    if (is.null(analysis_state()) || !isTRUE(analysis_state()$ok)) {
      return("")
    }
    current_analysis()$log
  })

  output$analysis_plots <- renderUI({
    if (is.null(analysis_state()) || !isTRUE(analysis_state()$ok)) {
      return(p("No hay graficos de analisis para mostrar todavia."))
    }

    plots <- current_analysis()$plots
    if (is.null(plots) || length(plots) < 1) {
      return(p("No hay graficos disponibles para este analisis."))
    }

    tagList(lapply(seq_along(plots), function(index) {
      plot_name <- names(plots)[[index]]
      outfile <- tempfile(fileext = ".png")
      build_plot_export(plots[[index]], outfile)
      image_src <- base64enc::dataURI(file = outfile, mime = "image/png")

      tagList(
        h4(plot_name),
        tags$div(
          style = "width: 100%; margin-bottom: 1.5rem;",
          tags$img(
            src = image_src,
            alt = plot_name,
            style = "display:block; width:100%; max-width:100%; height:auto;"
          )
        )
      )
    }))
  })

  output$analysis_tables <- renderUI({
    if (is.null(analysis_state()) || !isTRUE(analysis_state()$ok)) {
      return(p("No hay resultados de analisis para mostrar todavia."))
    }
    tables <- current_analysis()$tables

    tagList(lapply(seq_along(tables), function(index) {
      output_id <- sprintf("analysis_table_%s", index)
      output[[output_id]] <- renderTable({
        current_analysis()$tables[[index]]
      }, rownames = FALSE)
      tagList(h4(names(tables)[[index]]), tableOutput(output_id))
    }))
  })

  outputOptions(output, "analysis_response_selector", suspendWhenHidden = FALSE)
  outputOptions(output, "analysis_status", suspendWhenHidden = FALSE)
  outputOptions(output, "analysis_summary", suspendWhenHidden = FALSE)
  outputOptions(output, "analysis_log", suspendWhenHidden = FALSE)
  outputOptions(output, "analysis_tables", suspendWhenHidden = FALSE)
  outputOptions(output, "analysis_plots", suspendWhenHidden = FALSE)

  interpretation_attempt <- eventReactive(input$run_interpretation, {
    req(current_analysis())

    tryCatch(
      {
        plot_files <- character()
        if (!is.null(current_analysis()$plots) && length(current_analysis()$plots) > 0) {
          plot_files <- vapply(seq_along(current_analysis()$plots), function(index) {
            plot_file <- tempfile(fileext = ".png")
            build_plot_export(current_analysis()$plots[[index]], plot_file)
            plot_file
          }, character(1))
        } else {
          plot_file <- tempfile(fileext = ".png")
          current_analysis()$build_plot(plot_file)
          plot_files <- plot_file
        }

        list(
          ok = TRUE,
          text = openai_interpret_analysis(
            analysis_result = current_analysis(),
            plot_paths = plot_files,
            language = "es",
            extra_instructions = or_default(input$interpretation_instructions, "")
          )
        )
      },
      error = function(e) {
        list(ok = FALSE, text = conditionMessage(e))
      }
    )
  })

  output$interpretation_status <- renderText({
    if (is.null(interpretation_attempt())) {
      return("Sin interpretacion generada.")
    }

    if (isTRUE(interpretation_attempt()$ok)) {
      "Interpretacion generada."
    } else {
      "No se pudo generar la interpretacion."
    }
  })

  output$interpretation_text <- renderUI({
    req(interpretation_attempt())
    card(
      card_body(
        tags$div(
          style = paste(
            "white-space: pre-wrap; line-height: 1.5;",
            if (isTRUE(interpretation_attempt()$ok)) "" else "color: #b91c1c;"
          ),
          interpretation_attempt()$text
        )
      )
    )
  })

  output$package_status <- renderTable({
    package_status_table()
  }, rownames = FALSE)

  output$download_execution_workbook_ui <- renderUI({
    if (is.null(current_plan())) {
      return(p(style = "color: #6b7280;", "Genera un plan DOE para descargar la plantilla de ejecucion."))
    }

    tagList(
      p(style = "color: #4b5563; margin-bottom: 0.5rem;", "Descarga la plantilla operativa para capturar resultados y luego volver a cargarla en Ejecucion."),
      downloadButton("download_execution_workbook", "Descargar workbook de ejecucion")
    )
  })

  output$download_execution_workbook <- downloadHandler(
    filename = function() {
      req(current_plan())
      family_id <- current_plan()$design_family
      sprintf("doe-ejecucion-%s-%s.xlsx", family_id, Sys.Date())
    },
    content = function(file) {
      req(current_plan())
      write_doe_execution_workbook(
        path = file,
        plan_result = current_plan()
      )
    }
  )

  output$download_workbook_ui <- renderUI({
    if (is.null(current_plan())) {
      return(p(style = "color: #6b7280; margin-bottom: 0;", "Genera un plan DOE o carga un workbook DOE para habilitar la descarga del reporte."))
    }

    downloadButton("download_workbook", "Descargar reporte DOE")
  })

  output$download_workbook <- downloadHandler(
    filename = function() {
      req(current_plan())
      family_id <- current_plan()$design_family
      sprintf("doe-workbench-%s-%s.xlsx", family_id, Sys.Date())
    },
    content = function(file) {
      req(current_plan())
      interpretation_text <- if (!is.null(interpretation_attempt()) && isTRUE(interpretation_attempt()$ok)) {
        interpretation_attempt()$text
      } else {
        NULL
      }

      write_doe_export_workbook(
        path = file,
        plan_result = current_plan(),
        analysis_result = if (!is.null(analysis_state()) && isTRUE(analysis_state()$ok)) current_analysis() else NULL,
        interpretation_text = interpretation_text
      )
    }
  )
}
