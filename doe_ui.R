#' @title UI DOE workbench
NULL

library(shiny)
library(bslib)

ui <- page_sidebar(
  title = "DOE Industrial Workbench",
  fillable = TRUE,
  sidebar = sidebar(
    width = 380,
    h4("Planeacion"),
    selectInput(
      "design_family",
      "Familia DOE",
      choices = design_family_choices(),
      selected = "full_factorial"
    ),
    uiOutput("design_family_help"),
    textAreaInput(
      "factor_definitions",
      "Factores",
      value = paste(
        "Temperatura, 180, 220",
        "Presion, 45, 65",
        "Catalizador, A, B",
        sep = "\n"
      ),
      rows = 6,
      width = "100%",
      placeholder = "Una linea por factor: nombre, nivel_bajo, nivel_alto"
    ),
    textInput("response_names", "Respuestas esperadas", value = "Yield, DefectRate"),
    checkboxInput("randomize_runs", "Aleatorizar corridas", TRUE),
    numericInput("random_seed", "Semilla", value = 123, step = 1, min = 1),
    uiOutput("design_controls"),
    actionButton("generate_plan", "Generar plan DOE", class = "btn-primary"),
    tags$hr(),
    h4("Ejecucion"),
    fileInput("execution_file", "Cargar resultados ejecutados", accept = c(".csv", ".txt", ".xls", ".xlsx")),
    checkboxInput("execution_header", "La primera fila es encabezado", TRUE),
    uiOutput("execution_sheet_control"),
    conditionalPanel(
      condition = "input.execution_file && !/\\.(xls|xlsx)$/i.test(input.execution_file.name || '')",
      selectInput("execution_sep", "Separador", choices = c("Coma" = ",", "Punto y coma" = ";", "Tab" = "\t"), selected = ",")
    ),
    conditionalPanel(
      condition = "input.execution_file && !/\\.(xls|xlsx)$/i.test(input.execution_file.name || '')",
      selectInput("execution_dec", "Separador decimal", choices = c("Punto" = ".", "Coma" = ","), selected = ".")
    ),
    conditionalPanel(
      condition = "input.execution_file && !/\\.(xls|xlsx)$/i.test(input.execution_file.name || '')",
      selectInput("execution_quote", "Comillas", choices = c('Doble comilla' = '"', "Simple comilla" = "'", "Ninguna" = ""), selected = '"')
    ),
    uiOutput("response_selector"),
    actionButton("run_analysis", "Analizar DOE", class = "btn-primary")
  ),
  card(
    full_screen = TRUE,
    navset_card_tab(
      nav_panel(
        "Plan",
        br(),
        uiOutput("plan_status"),
        tableOutput("plan_summary"),
        tags$div(style = "margin: 1rem 0;", downloadButton("download_workbook", "Descargar workbook DOE")),
        h4("Factores"),
        tableOutput("factor_table"),
        h4("Plan de corridas"),
        tableOutput("plan_table")
      ),
      nav_panel(
        "Ejecucion",
        br(),
        uiOutput("execution_status"),
        tableOutput("execution_preview")
      ),
      nav_panel(
        "Analisis",
        br(),
        uiOutput("analysis_status"),
        tableOutput("analysis_summary"),
        h4("Traza del modelo"),
        verbatimTextOutput("analysis_log"),
        h4("Grafico"),
        plotOutput("analysis_plot", height = "480px"),
        h4("Tablas"),
        uiOutput("analysis_tables")
      ),
      nav_panel(
        "Interpretacion",
        br(),
        p("La interpretacion asistida por OpenAI es opcional y se ejecuta bajo demanda."),
        textAreaInput(
          "interpretation_instructions",
          "Instrucciones adicionales",
          rows = 4,
          width = "100%",
          placeholder = "Ejemplo: enfocate en factores criticos, riesgo de escalamiento y proxima iteracion experimental."
        ),
        actionButton("run_interpretation", "Generar interpretacion", class = "btn-primary"),
        verbatimTextOutput("interpretation_status"),
        uiOutput("interpretation_text")
      ),
      nav_panel(
        "Ayuda",
        br(),
        p("Esta version esta orientada a DOEs industriales: planear, ejecutar con plantilla y analizar efectos o superficies de respuesta."),
        h4("Stack recomendado"),
        tableOutput("package_status"),
        p("La app funciona con Base R para factorial completo y usa paquetes opcionales para fraccionales y RSM.")
      )
    )
  )
)
