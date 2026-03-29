# =============================================================================
# detect_tables.R -- Detect Excel named tables (ListObjects)
# =============================================================================

#' @keywords internal
#' @noRd
parse_table_ref <- function(ref) {
  parts <- strsplit(gsub("\\$", "", ref), ":")[[1]]
  col_start <- gsub("[0-9]", "", parts[1])
  row_start <- as.integer(gsub("[A-Z]", "", parts[1]))
  col_end <- gsub("[0-9]", "", parts[2])
  row_end <- as.integer(gsub("[A-Z]", "", parts[2]))

  list(
    col_start = col_start,
    col_end = col_end,
    row_start = row_start,
    row_end = row_end,
    col_start_idx = col_letter_to_index(col_start),
    col_end_idx = col_letter_to_index(col_end)
  )
}

#' @keywords internal
#' @noRd
detect_named_tables <- function(file_path, sheet_names = NULL) {
  wb <- openxlsx2::wb_load(file_path)

  if (is.null(sheet_names)) {
    sheet_names <- wb$get_sheet_names()
  }

  tables <- data.frame(
    sheet = character(),
    table_name = character(),
    ref = character(),
    header_row = integer(),
    data_start_row = integer(),
    data_end_row = integer(),
    col_start = character(),
    col_end = character(),
    stringsAsFactors = FALSE
  )

  for (s in sheet_names) {
    t <- tryCatch(wb$get_tables(sheet = s), error = function(e) NULL)
    if (is.null(t) || !is.data.frame(t) || nrow(t) == 0) next

    for (i in seq_len(nrow(t))) {
      ref <- t$tab_ref[i]
      parsed <- parse_table_ref(ref)

      # Read header row to get column names
      header_data <- tryCatch(
        openxlsx2::read_xlsx(
          file_path, sheet = s,
          rows = parsed$row_start,
          cols = parsed$col_start_idx:parsed$col_end_idx,
          col_names = FALSE
        ),
        error = function(e) NULL
      )

      if (!is.null(header_data) && nrow(header_data) > 0) {
        col_names <- as.character(unlist(header_data[1, ]))
      } else {
        # Fallback: generate column letters
        n_cols <- parsed$col_end_idx - parsed$col_start_idx + 1
        col_names <- generate_col_names(n_cols)
      }

      tables <- rbind(tables, data.frame(
        sheet = s,
        table_name = t$tab_name[i],
        ref = ref,
        header_row = parsed$row_start,
        data_start_row = parsed$row_start + 1,
        data_end_row = parsed$row_end,
        col_start = parsed$col_start,
        col_end = parsed$col_end,
        col_names = I(list(col_names)),
        stringsAsFactors = FALSE
      ))
    }
  }

  tables
}
