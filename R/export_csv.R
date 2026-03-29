# =============================================================================
# export_csv.R — Export sheet data as tidy long-format CSVs
# =============================================================================

#' Export each sheet's raw (non-formula) cell values as tidy CSV
#'
#' @param file_path Path to .xlsx file
#' @param sheet_names Sheets to export
#' @param output_dir Directory to write CSVs to
#' @return Named character vector: sheet_name -> csv_path
export_sheet_csvs <- function(file_path, sheet_names, output_dir) {
  dir.create(output_dir, showWarnings = FALSE, recursive = TRUE)

  all_cells <- tidyxl::xlsx_cells(file_path, sheets = sheet_names)

  paths <- character(0)

  for (s in sheet_names) {
    sheet_cells <- all_cells[all_cells$sheet == s, ]

    # Keep only non-formula cells
    value_cells <- sheet_cells[is.na(sheet_cells$formula), ]

    # Pre-allocate vectors
    n <- nrow(value_cells)
    rows_vec <- integer(n)
    cols_vec <- character(n)
    vals_vec <- character(n)
    keep <- logical(n)

    for (i in seq_len(n)) {
      cell <- value_cells[i, ]

      val <- NA_character_
      if (!is.na(cell$character))    val <- cell$character
      else if (!is.na(cell$numeric)) val <- as.character(cell$numeric)
      else if (!is.na(cell$logical)) val <- as.character(cell$logical)
      else if (!is.na(cell$date))    val <- as.character(cell$date)

      if (is.na(val)) next

      rows_vec[i] <- cell$row
      cols_vec[i] <- index_to_col_letter(cell$col)
      vals_vec[i] <- val
      keep[i] <- TRUE
    }

    tidy_rows <- data.frame(
      row = rows_vec[keep],
      col = cols_vec[keep],
      value = vals_vec[keep],
      stringsAsFactors = FALSE
    )

    filename <- paste0(sanitize_sheet_name(s), ".csv")
    csv_path <- file.path(output_dir, filename)
    write.csv(tidy_rows, csv_path, row.names = FALSE, quote = TRUE)
    paths[s] <- csv_path
  }

  paths
}
