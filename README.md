# DOE Industrial Workbench

![R](https://img.shields.io/badge/R-4.5+-276DC3?logo=r&logoColor=white)
![Shiny](https://img.shields.io/badge/Shiny-App-75AADB?logo=rstudioide&logoColor=white)
![bslib](https://img.shields.io/badge/UI-bslib-4B5563)
![DOE](https://img.shields.io/badge/Domain-DOE-0F766E)
![OpenAI](https://img.shields.io/badge/AI-OpenAI-412991?logo=openai&logoColor=white)
![Excel](https://img.shields.io/badge/Output-Excel-217346?logo=microsoft-excel&logoColor=white)

Aplicacion Shiny para planear, ejecutar y analizar disenos de experimentos (DOE) a nivel industrial. La app ayuda a definir factores y niveles, generar el plan experimental, descargar una plantilla operativa para ejecucion, cargar resultados ejecutados, ajustar modelos DOE y exportar un reporte consolidado.

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
5. Descarga el workbook de ejecucion desde la pestana `Plan`.
6. Ejecuta las corridas y captura resultados.
7. Carga el archivo ejecutado en la app.
8. Selecciona una respuesta y corre el analisis.
9. Revisa tablas, diagnosticos y graficos en las pestanas `Analisis` y `Graficos`.
10. Genera interpretacion opcional con OpenAI.
11. Exporta el reporte DOE consolidado.

## Archivos de ejemplo

El proyecto incluye workbooks listos para probar cada familia de diseno:

- `doe_example_full_factorial.xlsx`
- `doe_example_fractional_factorial.xlsx`
- `doe_example_plackett_burman.xlsx`
- `doe_example_ccd.xlsx`
- `doe_example_box_behnken.xlsx`

Contenido:

- `Resumen_plan`: resumen del DOE ejemplo
- `Factores`: definicion de factores y niveles
- `Plan_corridas`: corridas con respuestas capturadas
- `Instrucciones`: pasos sugeridos para validar el flujo

Los ejemplos usan respuestas simuladas para producir senales mas creibles segun el tipo de DOE. Para probar la app, carga cualquiera de esos archivos en `Ejecucion`, selecciona la hoja `Plan_corridas` y analiza la respuesta principal del workbook.

## Funcionalidades

- definicion rapida de factores numericos y categoricos de dos niveles
- generacion de plan con orden estandar y orden de corrida
- columnas operativas para ejecucion: estado, lote, operador, timestamp y comentarios
- descarga de workbook de ejecucion desde `Plan`, con formato cargable por la propia app
- soporte para respuestas numericas capturadas en el workbook
- ajuste de modelos DOE para efectos e interacciones
- ajuste de modelos de segundo orden para disenos RSM
- pestana dedicada de `Graficos` con multiples diagnosticos
- exportacion a Excel del plan de ejecucion y del reporte de analisis
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
- observado vs ajustado
- residuales vs ajustados
- residuales por orden de corrida
- QQ plot de residuales
- traza textual del modelo

## Exportaciones

La app ofrece dos salidas distintas:

- `Descargar workbook de ejecucion`: genera la plantilla operativa desde la pestana `Plan`. Ese archivo es el que luego se vuelve a cargar en `Ejecucion`.
- `Descargar reporte DOE`: genera un workbook consolidado con resumen del plan, tablas del analisis, todos los graficos disponibles e interpretacion si existe.

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
- todos los graficos del analisis disponibles

## Estructura del proyecto

- `global.R`: punto de entrada global
- `ui.R`: punto de entrada de interfaz
- `server.R`: punto de entrada de servidor
- `doe_engine.R`: motor DOE
- `doe_openai.R`: interpretacion DOE con OpenAI

## Limitaciones actuales

- el analisis es monorespuesta por corrida; cada respuesta se analiza por separado
- no hay manejo formal de bloques, split-plot o restricciones de aleatorizacion
- la app no calcula todavia estructura formal de alias ni bloqueo industrial
- los caminos `FrF2` y `rsm` dependen de que esos paquetes esten instalados en el entorno
- la interpretacion con OpenAI no sustituye criterio estadistico o de proceso

## Licencia

Este proyecto se distribuye bajo la licencia Creative Commons Attribution 4.0 International (`CC BY 4.0`). Consulta `LICENSE.md`.
