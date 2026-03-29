#' excel2r: Migrate Excel Workbooks to R
#'
#' Converts Excel formula workbooks into standalone R scripts. Extracts
#' every formula, translates 62 Excel functions to R equivalents, resolves
#' cross-sheet references, determines execution order via topological sort,
#' and optionally exports raw data as tidy CSVs for a fully Excel-free workflow.
#'
#' @section Main functions:
#' \describe{
#'   \item{\code{\link{migrate}}}{One-step migration: process, export, verify}
#'   \item{\code{\link{process}}}{Process workbook without writing files}
#'   \item{\code{\link{verify}}}{Compare R results against Excel cached values}
#'   \item{\code{\link{supported_functions}}}{List the 62 supported Excel functions}
#'   \item{\code{\link{run_app}}}{Launch the interactive Shiny interface}
#' }
#'
#' @importFrom stats setNames
#' @importFrom utils write.csv
#' @docType package
#' @name excel2r-package
"_PACKAGE"
