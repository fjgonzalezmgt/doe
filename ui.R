#' @title UI DOE workbench
NULL

library(shiny)
library(bslib)

ui <- page_sidebar(
  title = "DOE Industrial Workbench",
  fillable = TRUE,
  tags$head(
    tags$style(HTML("
      #analysis_plot {
        width: 100%;
      }
      #analysis_plot img {
        display: block;
        width: 100% !important;
        max-width: 100%;
        height: auto !important;
      }
    "))
  ),
  sidebar = sidebar(
    width = 380,
    h4("Planeacion"),
    p(
      style = "font-size: 0.9rem; color: #6b7280;",
      "Hint: empieza con pocos factores bien definidos. Si estas en screening, prioriza amplitud; si ya conoces el proceso, enfocate en optimizacion."
    ),
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
    p(
      style = "font-size: 0.85rem; color: #6b7280; margin-top: -0.5rem;",
      "Hint: usa niveles realmente operables. Si el factor es numerico, define un bajo y alto que el proceso pueda sostener."
    ),
    textInput("response_names", "Respuestas esperadas", value = "Yield, DefectRate"),
    p(
      style = "font-size: 0.85rem; color: #6b7280; margin-top: -0.5rem;",
      "Hint: separa respuestas con coma. Aunque declares varias, el analisis se corre una respuesta a la vez."
    ),
    checkboxInput("randomize_runs", "Aleatorizar corridas", TRUE),
    numericInput("random_seed", "Semilla", value = 123, step = 1, min = 1),
    p(
      style = "font-size: 0.85rem; color: #6b7280; margin-top: -0.5rem;",
      "Hint: aleatoriza salvo que exista una restriccion real de planta. Si necesitas repetir el mismo orden, conserva la semilla."
    ),
    uiOutput("design_controls"),
    actionButton("generate_plan", "Generar plan DOE", class = "btn-primary"),
    tags$hr(),
    h4("Ejecucion"),
    p(
      style = "font-size: 0.9rem; color: #6b7280;",
      "Hint: si generaste el plan aqui mismo, lo mas seguro es descargar el workbook de ejecucion, capturar resultados ahi y volverlo a cargar."
    ),
    selectInput(
      "execution_design_family",
      "Tipo de DOE para ejecutar/analizar",
      choices = design_family_choices(),
      selected = "full_factorial"
    ),
    uiOutput("execution_family_help"),
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
    actionButton("run_analysis", "Analizar DOE", class = "btn-primary")
  ),
  card(
    full_screen = TRUE,
    tags$div(
      style = "padding: 1rem 1rem 0 1rem;",
      uiOutput("download_workbook_ui")
    ),
    navset_card_tab(
      id = "main_tabs",
      nav_panel(
        "Plan",
        br(),
        p(
          style = "color: #4b5563;",
          "Hint: usa esta pestana para validar factores, numero de corridas y orden de ejecucion. `std_order` es el orden teorico del diseno; `run_order` es el orden sugerido para correrlo."
        ),
        uiOutput("download_execution_workbook_ui"),
        uiOutput("plan_status"),
        tableOutput("plan_summary"),
        h4("Factores"),
        tableOutput("factor_table"),
        h4("Plan de corridas"),
        tableOutput("plan_table")
      ),
      nav_panel(
        "Ejecucion",
        br(),
        p(
          style = "color: #4b5563;",
          "Hint: revisa aqui que el archivo cargado conserve nombres de columnas y que las respuestas tengan valores numericos donde corresponde."
        ),
        uiOutput("execution_status"),
        tableOutput("execution_preview")
      ),
      nav_panel(
        "Analisis",
        br(),
        p(
          style = "color: #4b5563;",
          "Hint: primero elige una respuesta. Empieza leyendo resumen, luego ANOVA y coeficientes; deja la traza del modelo para validar el detalle."
        ),
        uiOutput("analysis_status"),
        uiOutput("analysis_response_selector"),
        tableOutput("analysis_summary"),
        h4("Traza del modelo"),
        verbatimTextOutput("analysis_log"),
        h4("Tablas"),
        uiOutput("analysis_tables")
      ),
      nav_panel(
        "Graficos",
        br(),
        p(
          style = "color: #4b5563;",
          "Hint: si `Observado vs ajustado` se dispersa demasiado o los residuales muestran patron, no tomes decisiones solo con p-values."
        ),
        uiOutput("analysis_plots")
      ),
      nav_panel(
        "Interpretacion",
        br(),
        p("La interpretacion asistida por OpenAI es opcional y se ejecuta bajo demanda."),
        p(
          style = "color: #4b5563;",
          "Hint: usala para resumir hallazgos y siguientes pasos, no para reemplazar el criterio estadistico o de proceso."
        ),
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
        h4("Flujo sugerido"),
        tags$ol(
          tags$li("Define factores y respuestas en la barra lateral."),
          tags$li("Genera el plan y valida corridas, niveles y aleatorizacion."),
          tags$li("Descarga el workbook de ejecucion y captura resultados reales."),
          tags$li("Carga el workbook completado en Ejecucion."),
          tags$li("Analiza una respuesta a la vez y revisa tablas y graficos antes de concluir.")
        ),
        h4("Pistas rapidas"),
        tags$ul(
          tags$li("Si estas explorando muchos factores, empieza por factorial fraccional o Plackett-Burman."),
          tags$li("Si buscas optimizar y sospechas curvatura, usa CCD o Box-Behnken."),
          tags$li("Si el archivo cargado no analiza, revisa nombres de columnas, hoja correcta y respuestas numericas."),
          tags$li("Un buen R2 no basta: revisa tambien residuales y orden de corrida.")
        ),
        h4("Stack recomendado"),
        tableOutput("package_status"),
        p("La app funciona con Base R para factorial completo y usa paquetes opcionales para fraccionales y RSM.")
      )
    )
  )
)
