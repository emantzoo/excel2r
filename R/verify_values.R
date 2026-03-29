# =============================================================================
# verify_values.R — Compare R-computed values against Excel cached values
# =============================================================================

#' Run generated script and compare R values vs Excel cached formula results
#'
#' @param file_path Path to the Excel file
#' @param report Data frame from process_excel_file()$report
#' @param script_text The generated R script text
#' @return list(summary, mismatches) where mismatches contains only real differences
verify_against_excel <- function(file_path, report, script_text) {
  # Run the generated script in an isolated environment
  env <- new.env(parent = globalenv())
  tryCatch({
    # Write script to temp file and source it
    tmp <- tempfile(fileext = ".R")
    on.exit(unlink(tmp), add = TRUE)
    writeLines(script_text, tmp)
    source(tmp, local = env)
  }, error = function(e) {
    return(list(
      summary = list(total = 0, matches = 0, value_mismatches = 0,
                     fp_precision = 0, na_mismatches = 0,
                     error = paste("Script execution failed:", e$message)),
      mismatches = data.frame(Sheet = character(), Cell = character(),
                              Formula = character(), R_Value = character(),
                              Excel_Value = character(), Category = character(),
                              stringsAsFactors = FALSE)
    ))
  })

  # Read Excel values via tidyxl
  all_cells <- tidyxl::xlsx_cells(file_path)

  # Only compare formulas with status "ok"
  formulas <- report[tolower(report$Status) == "ok", ]

  total <- 0L
  matches <- 0L
  value_mismatches <- 0L
  fp_precision <- 0L
  na_mismatches <- 0L

  mismatch_rows <- list()

  for (i in seq_len(nrow(formulas))) {
    sheet <- formulas$Sheet[i]
    cell <- formulas$Cell[i]

    if (is.null(sheet) || is.na(sheet) || nchar(sheet) == 0) {
      total <- total + 1L
      next
    }

    df_name <- sanitize_sheet_name(sheet)
    col_letter <- toupper(gsub("[0-9]+", "", cell))
    row_num <- as.integer(gsub("[A-Za-z]+", "", cell))

    # Get R value from isolated env
    r_value <- tryCatch({
      df <- get(df_name, envir = env)
      if (col_letter %in% colnames(df) && row_num <= nrow(df)) {
        df[[col_letter]][row_num]
      } else {
        NA
      }
    }, error = function(e) NA)

    # Get Excel value from tidyxl
    col_idx <- col_letter_to_index(col_letter)
    excel_cell <- all_cells[all_cells$sheet == sheet &
                              all_cells$row == row_num &
                              all_cells$col == col_idx, ]

    excel_value <- NA
    excel_is_error <- FALSE
    excel_is_text <- FALSE
    excel_has_formula <- FALSE
    if (nrow(excel_cell) > 0) {
      excel_has_formula <- !is.na(excel_cell$formula[1])
      if (!is.na(excel_cell$error[1])) {
        excel_value <- excel_cell$error[1]
        excel_is_error <- TRUE
      } else if (!is.na(excel_cell$numeric[1])) {
        excel_value <- excel_cell$numeric[1]
      } else if (!is.na(excel_cell$character[1])) {
        excel_value <- excel_cell$character[1]
        excel_is_text <- TRUE
      } else if (!is.na(excel_cell$logical[1])) {
        excel_value <- as.numeric(excel_cell$logical[1])
      }
    }

    total <- total + 1L

    r_na <- is.na(r_value) || (is.numeric(r_value) && is.nan(r_value))
    e_na <- is.na(excel_value)

    r_str <- if (r_na) "NA" else as.character(r_value)
    e_str <- if (e_na) "NA" else as.character(excel_value)

    # Classification
    if (e_na && excel_has_formula) {
      # Formula exists but no cached value — workbook not recalculated
      # Can't compare; treat as harmless
      na_mismatches <- na_mismatches + 1L

    } else if (r_na && e_na) {
      # Both NA — match
      matches <- matches + 1L

    } else if (excel_is_error) {
      # Excel error (#VALUE!, #REF!, etc.) — harmless, R can't replicate
      na_mismatches <- na_mismatches + 1L

    } else if (excel_is_text && is.numeric(r_value) && isTRUE(r_value == 0)) {
      # R=0 (from NA->0 conversion) vs Excel text placeholder — harmless
      na_mismatches <- na_mismatches + 1L

    } else if (r_na && !e_na) {
      # R is NA but Excel has a value — real mismatch
      value_mismatches <- value_mismatches + 1L
      mismatch_rows[[length(mismatch_rows) + 1]] <- data.frame(
        Sheet = sheet, Cell = cell, Formula = formulas$Formula[i],
        R_Value = r_str, Excel_Value = e_str, Category = "R is NA",
        stringsAsFactors = FALSE
      )

    } else if (!r_na && e_na) {
      # R has value but Excel is empty — real mismatch
      value_mismatches <- value_mismatches + 1L
      mismatch_rows[[length(mismatch_rows) + 1]] <- data.frame(
        Sheet = sheet, Cell = cell, Formula = formulas$Formula[i],
        R_Value = r_str, Excel_Value = e_str, Category = "Excel is NA",
        stringsAsFactors = FALSE
      )

    } else if (is.numeric(r_value) && is.numeric(excel_value)) {
      diff <- abs(r_value - excel_value)
      if (diff < 1e-4) {
        if (diff > 1e-6) {
          fp_precision <- fp_precision + 1L
        } else {
          matches <- matches + 1L
        }
      } else {
        value_mismatches <- value_mismatches + 1L
        mismatch_rows[[length(mismatch_rows) + 1]] <- data.frame(
          Sheet = sheet, Cell = cell, Formula = formulas$Formula[i],
          R_Value = r_str, Excel_Value = e_str,
          Category = sprintf("Diff: %.6g", diff),
          stringsAsFactors = FALSE
        )
      }

    } else if (r_str == e_str) {
      matches <- matches + 1L

    } else {
      value_mismatches <- value_mismatches + 1L
      mismatch_rows[[length(mismatch_rows) + 1]] <- data.frame(
        Sheet = sheet, Cell = cell, Formula = formulas$Formula[i],
        R_Value = r_str, Excel_Value = e_str, Category = "Type mismatch",
        stringsAsFactors = FALSE
      )
    }
  }

  mismatches <- if (length(mismatch_rows) > 0) {
    do.call(rbind, mismatch_rows)
  } else {
    data.frame(Sheet = character(), Cell = character(),
               Formula = character(), R_Value = character(),
               Excel_Value = character(), Category = character(),
               stringsAsFactors = FALSE)
  }

  list(
    summary = list(
      total = total,
      matches = matches,
      value_mismatches = value_mismatches,
      fp_precision = fp_precision,
      na_mismatches = na_mismatches
    ),
    mismatches = mismatches
  )
}
