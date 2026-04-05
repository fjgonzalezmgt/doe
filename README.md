# DOE Industrial Workbench

![R](https://img.shields.io/badge/R-4.5+-276DC3?logo=r&logoColor=white)
![Shiny](https://img.shields.io/badge/Shiny-App-75AADB?logo=rstudioide&logoColor=white)
![bslib](https://img.shields.io/badge/UI-bslib-4B5563)
![DOE](https://img.shields.io/badge/Domain-DOE-0F766E)
![OpenAI](https://img.shields.io/badge/AI-OpenAI-412991?logo=openai&logoColor=white)
![Excel](https://img.shields.io/badge/Output-Excel-217346?logo=microsoft-excel&logoColor=white)

Aplicacion Shiny para planear, ejecutar y analizar disenos de experimentos (DOE) a nivel industrial. La app ayuda a definir factores y niveles, generar el plan experimental, cargar resultados ejecutados, ajustar modelos DOE y exportar un workbook consolidado.

## Alcance

La app esta orientada a tres momentos del trabajo experimental:

1. Planeacion del DOE.
2. Ejecucion con plantilla exportable.
3. Analisis estadistico e interpretacion.

## Disenos soportados

- factorial completo de 2 niveles
- factorial fraccional de 2 niveles
- Plackett-Burman para screening
- Central Composite Design (CCD)
- Box-Behnken

## Flujo de trabajo

1. Define factores en formato `nombre, nivel_bajo, nivel_alto`.
2. Define una o varias respuestas esperadas.
3. Selecciona la familia DOE y sus controles.
4. Genera el plan con aleatorizacion opcional.
5. Descarga el workbook DOE.
6. Ejecuta las corridas y captura resultados.
7. Carga el archivo ejecutado en la app.
8. Selecciona una respuesta y corre el analisis.
9. Revisa ANOVA, coeficientes, grafico e interpretacion opcional.
10. Exporta el workbook consolidado.

## Archivo de ejemplo

El proyecto incluye `doe_example_workbook.xlsx`, un ejemplo listo para probar la app.

Contenido:

- `Resumen_plan`: resumen del DOE ejemplo
- `Factores`: definicion de factores y niveles
- `Plan_corridas`: corridas con respuestas capturadas
- `Instrucciones`: pasos sugeridos para validar el flujo

Para probar la app con este archivo, carga la hoja `Plan_corridas` como archivo de ejecucion y analiza `Yield` o `DefectRate`.

## Funcionalidades

- definicion rapida de factores numericos y categoricos de dos niveles
- generacion de plan con orden estandar y orden de corrida
- columnas operativas para ejecucion: estado, lote, operador, timestamp y comentarios
- soporte para respuestas numericas capturadas en el workbook
- ajuste de modelos DOE para efectos e interacciones
- ajuste de modelos de segundo orden para disenos RSM
- exportacion a Excel del plan y del analisis
- interpretacion opcional con OpenAI

## Stack

Paquetes base de la app:

```r
install.packages(c(
  "shiny",
  "bslib",
  "ggplot2",
  "readxl",
  "openxlsx",
  "httr2",
  "jsonlite",
  "base64enc"
))
```

Paquetes opcionales para ampliar tipos de DOE:

```r
install.packages(c(
  "FrF2",
  "DoE.base",
  "rsm"
))
```

Rol de cada paquete:

- `shiny`: interfaz web interactiva
- `bslib`: layout y componentes visuales
- `ggplot2`: visualizacion de efectos
- `readxl`: lectura de archivos Excel
- `openxlsx`: exportacion de workbooks
- `httr2`, `jsonlite`, `base64enc`: integracion con OpenAI
- `FrF2`: factoriales fraccionales y Plackett-Burman
- `DoE.base`: infraestructura DOE complementaria
- `rsm`: superficie de respuesta, CCD y Box-Behnken

## Ejecucion

Desde la carpeta del proyecto:

```r
shiny::runApp()
```

## Formato de factores

Usa una linea por factor:

```text
Temperatura, 180, 220
Presion, 45, 65
Catalizador, A, B
```

Reglas:

- si ambos niveles son numericos, el factor se interpreta como numerico
- si alguno no es numerico, el factor se interpreta como categorico
- para factores numericos, el nivel alto debe ser mayor que el nivel bajo

## Analisis

La app ajusta modelos lineales sobre las corridas ejecutadas:

- en factoriales y screening, estima efectos principales y, cuando aplica, interacciones de segundo orden
- en CCD y Box-Behnken, ajusta un modelo de segundo orden

Las salidas actuales incluyen:

- resumen del modelo
- tabla ANOVA
- tabla de coeficientes
- corridas analizadas
- grafico de coeficientes
- traza textual del modelo

## OpenAI

La interpretacion asistida es opcional.

Configura un archivo `.Renviron` en la raiz del proyecto usando como referencia `.Renviron.example`:

```env
OPENAI_API_KEY=tu_api_key
OPENAI_MODEL=gpt-5-mini
```

Cuando se solicita interpretacion, la app envia:

- resumen del analisis
- tablas del modelo
- grafico exportado

## Estructura del proyecto

- `global.R`: punto de entrada global
- `ui.R`: punto de entrada de interfaz
- `server.R`: punto de entrada de servidor
- `doe_engine.R`: motor DOE
- `doe_ui.R`: interfaz DOE
- `doe_server.R`: logica reactiva DOE
- `doe_openai.R`: interpretacion DOE con OpenAI
- `openai_helpers.R`: wrapper de compatibilidad para helpers de OpenAI

## Limitaciones actuales

- el analisis es monorespuesta por corrida; cada respuesta se analiza por separado
- no hay manejo formal de bloques, split-plot o restricciones de aleatorizacion
- la app no calcula todavia estructura de alias ni diagnosticos avanzados de residuos
- los caminos `FrF2` y `rsm` dependen de que esos paquetes esten instalados en el entorno
- la interpretacion con OpenAI no sustituye criterio estadistico o de proceso

## Licencia

Este proyecto se distribuye bajo la licencia Creative Commons Attribution 4.0 International (`CC BY 4.0`). Consulta `LICENSE.md`.
