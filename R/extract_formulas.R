# =============================================================================
# extract_formulas.R — Extract formulas and dimensions from any Excel file
# =============================================================================

#' Extract all formulas from an Excel file
#' Auto-detects all sheets if none specified
#' @param file_path Path to .xlsx file
#' @param sheet_names Optional character vector of sheet names to process
#' @return Data frame with columns: Sheet, Cell, Row, Col, Formula
extract_all_formulas <- function(file_path, sheet_names = NULL) {
  if (!requireNamespace("tidyxl", quietly = TRUE)) {
    stop("Package 'tidyxl' is required. Install with: install.packages('tidyxl')")
  }

  # Auto-detect sheets if not specified
  if (is.null(sheet_names)) {
    sheet_names <- readxl::excel_sheets(file_path)
  }

  # Read all cells
  cells <- tidyxl::xlsx_cells(file_path, sheets = sheet_names)

  # Filter for cells with formulas
  formula_cells <- cells[!is.na(cells$formula),
                         c("sheet", "address", "row", "col", "formula")]

  if (nrow(formula_cells) == 0) {
    message("No formulas found in the specified sheets.")
    return(data.frame(Sheet = character(), Cell = character(),
                      Row = integer(), Col = integer(),
                      Formula = character(), stringsAsFactors = FALSE))
  }

  colnames(formula_cells) <- c("Sheet", "Cell", "Row", "Col", "Formula")
  formula_cells$Row <- as.integer(formula_cells$Row)
  formula_cells$Col <- as.integer(formula_cells$Col)

  # Reset row names
  rownames(formula_cells) <- NULL
  formula_cells
}

#' Detect actual dimensions (max row, max col) for each sheet
#' @param file_path Path to .xlsx file
#' @param sheet_names Optional character vector of sheet names
#' @return Named list: sheet_name -> list(max_row, max_col)
detect_sheet_dimensions <- function(file_path, sheet_names = NULL) {
  if (is.null(sheet_names)) {
    sheet_names <- readxl::excel_sheets(file_path)
  }

  cells <- tidyxl::xlsx_cells(file_path, sheets = sheet_names)

  dims <- list()
  for (sheet in sheet_names) {
    sheet_cells <- cells[cells$sheet == sheet, ]
    if (nrow(sheet_cells) > 0) {
      dims[[sheet]] <- list(
        max_row = max(sheet_cells$row, na.rm = TRUE),
        max_col = max(sheet_cells$col, na.rm = TRUE)
      )
    } else {
      dims[[sheet]] <- list(max_row = 0, max_col = 0)
    }
  }
  dims
}

#' Identify which Excel functions are used in the formulas
#' @param formula_data Data frame from extract_all_formulas()
#' @return Named integer vector: function_name -> count
detect_used_functions <- function(formula_data) {
  # Match function names: word chars followed by (
  all_funcs <- unlist(regmatches(formula_data$Formula,
                                  gregexpr("[A-Z][A-Z0-9.]*(?=\\()", formula_data$Formula, perl = TRUE, ignore.case = TRUE)))
  if (length(all_funcs) == 0) return(integer(0))
  sort(table(all_funcs), decreasing = TRUE)
}
