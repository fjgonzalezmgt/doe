#' @title Punto de entrada global para DOE Industrial Workbench
#' @description Carga dependencias base y sourcea los archivos de logica compartida
#'   para la aplicacion Shiny.
#' @keywords internal
NULL

library(shiny)

source("doe_engine.R", local = TRUE)
source("doe_openai.R", local = TRUE)
