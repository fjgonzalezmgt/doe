#' @title Server DOE workbench
NULL

server <- function(input, output, session) {
  output$design_family_help <- renderUI({
    meta <- design_family_catalog[[input$design_family]]
    req(meta)
    tagList(
      p(tags$strong(meta$label)),
      p(style = "color: #4b5563; margin-bottom: 0.25rem;", meta$objective),
      p(style = "font-size: 0.9rem; color: #6b7280;", meta$notes)
    )
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
    selectInput("execution_sheet", "Hoja", choices = sheets, selected = sheets[[1]])
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
    req(plan_attempt(), isTRUE(plan_attempt()$ok))
    plan_attempt()$plan
  })

  output$plan_status <- renderUI({
    if (is.null(plan_attempt())) {
      return(p("Define factores y genera el plan DOE."))
    }

    if (isTRUE(plan_attempt()$ok)) {
      return(p(style = "color: #0f766e;", "Plan DOE generado."))
    }

    p(style = "color: #b91c1c;", plan_attempt()$message)
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

  execution_data <- reactive({
    if (is.null(input$execution_file)) {
      return(NULL)
    }

    read_input_data(
      path = input$execution_file$datapath,
      header = input$execution_header,
      sep = or_default(input$execution_sep, ","),
      quote = or_default(input$execution_quote, '"'),
      dec = or_default(input$execution_dec, "."),
      sheet = if (tolower(tools::file_ext(input$execution_file$name)) %in% c("xls", "xlsx")) input$execution_sheet else NULL
    )
  })

  analysis_dataset <- reactive({
    if (is.null(plan_attempt()) || !isTRUE(plan_attempt()$ok)) {
      return(NULL)
    }
    coerce_analysis_dataset(current_plan(), execution_data())
  })

  output$execution_status <- renderUI({
    if (is.null(current_plan())) {
      return(p("Primero genera un plan DOE."))
    }

    if (is.null(execution_data())) {
      return(p("Puedes descargar el workbook DOE, capturar resultados y luego cargar el archivo completado."))
    }

    p(sprintf("Archivo cargado con %s filas y %s columnas.", nrow(execution_data()), ncol(execution_data())))
  })

  output$execution_preview <- renderTable({
    req(analysis_dataset())
    utils::head(analysis_dataset(), 15)
  }, rownames = FALSE)

  output$response_selector <- renderUI({
    req(current_plan())
    df <- analysis_dataset()
    req(df)
    response_choices <- recommended_response_columns(df, current_plan()$factors)
    if (length(response_choices) < 1) {
      return(p("No hay columnas numericas de respuesta disponibles todavia."))
    }

    selectInput("response_col", "Respuesta a analizar", choices = response_choices, selected = response_choices[[1]])
  })

  analysis_attempt <- eventReactive(input$run_analysis, {
    req(current_plan(), analysis_dataset(), input$response_col)
    tryCatch(
      {
        list(
          ok = TRUE,
          analysis = run_doe_analysis(
            plan_result = current_plan(),
            execution_df = analysis_dataset(),
            response_col = input$response_col
          )
        )
      },
      error = function(e) {
        list(ok = FALSE, message = conditionMessage(e))
      }
    )
  })

  current_analysis <- reactive({
    req(analysis_attempt(), isTRUE(analysis_attempt()$ok))
    analysis_attempt()$analysis
  })

  output$analysis_status <- renderUI({
    if (is.null(analysis_attempt())) {
      return(p("Carga datos ejecutados y corre el analisis DOE."))
    }

    if (isTRUE(analysis_attempt()$ok)) {
      return(p(style = "color: #0f766e;", "Analisis DOE completado."))
    }

    p(style = "color: #b91c1c;", analysis_attempt()$message)
  })

  output$analysis_summary <- renderTable({
    req(current_analysis())
    data_frame_from_named_list(current_analysis()$summary)
  }, rownames = FALSE)

  output$analysis_log <- renderText({
    req(current_analysis())
    current_analysis()$log
  })

  output$analysis_plot <- renderPlot({
    req(current_analysis())
    print(current_analysis()$plot_obj)
  })

  output$analysis_tables <- renderUI({
    req(current_analysis())
    tables <- current_analysis()$tables

    tagList(lapply(seq_along(tables), function(index) {
      output_id <- sprintf("analysis_table_%s", index)
      output[[output_id]] <- renderTable({
        current_analysis()$tables[[index]]
      }, rownames = FALSE)
      tagList(h4(names(tables)[[index]]), tableOutput(output_id))
    }))
  })

  interpretation_attempt <- eventReactive(input$run_interpretation, {
    req(current_analysis())

    tryCatch(
      {
        plot_file <- tempfile(fileext = ".png")
        current_analysis()$build_plot(plot_file)
        list(
          ok = TRUE,
          text = openai_interpret_analysis(
            analysis_result = current_analysis(),
            plot_path = plot_file,
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

  output$download_workbook <- downloadHandler(
    filename = function() {
      family_id <- if (!is.null(current_plan())) current_plan()$design_family else "doe"
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
        analysis_result = if (!is.null(analysis_attempt()) && isTRUE(analysis_attempt()$ok)) current_analysis() else NULL,
        interpretation_text = interpretation_text
      )
    }
  )
}
