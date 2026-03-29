# =============================================================================
# export_csv.R -- Export sheet data as tidy long-format CSVs
# =============================================================================

#' @keywords internal
#' @noRd
extract_referenced_cells <- function(formula_data, sheet_names, sheet_dims = NULL) {
  referenced <- setNames(
    lapply(sheet_names, function(s) list(cells = character(0), whole_cols = character(0))),
    sheet_names
  )

  for (i in seq_len(nrow(formula_data))) {
    formula_sheet <- formula_data$Sheet[i]
    formula_text  <- formula_data$Formula[i]
    if (is.na(formula_text) || formula_text == "") next

    # --- Step 1: Extract range references using existing parser ---
    range_result <- extract_ranges(formula_text)
    placeholder_formula <- range_result$placeholder_formula
    placeholder_map     <- range_result$placeholder_map

    # --- Step 2: Process each range reference ---
    for (ph in names(placeholder_map)) {
      range_ref <- placeholder_map[[ph]]

      # Determine target sheet
      target_sheet <- formula_sheet
      cell_part <- range_ref
      if (grepl("!", range_ref)) {
        parts <- strsplit(range_ref, "!", fixed = FALSE)[[1]]
        sheet_part <- gsub("^'+|'+$", "", parts[1])
        # Match against known sheet names
        for (s in sheet_names) {
          if (s == sheet_part || sanitize_sheet_name(s) == sanitize_sheet_name(sheet_part)) {
            target_sheet <- s
            break
          }
        }
        cell_part <- parts[2]
      }

      # Parse the range endpoints
      clean_part <- gsub("\\$", "", cell_part)
      range_parts <- strsplit(clean_part, ":")[[1]]
      if (length(range_parts) != 2) next

      p1 <- parse_cell_address(range_parts[1])
      p2 <- parse_cell_address(range_parts[2])

      if (is.na(p1$row) || is.na(p2$row)) {
        # Whole-column reference like A:A or B:D -- mark columns, don't expand
        col1_idx <- col_letter_to_index(p1$col)
        col2_idx <- col_letter_to_index(p2$col)
        for (ci in col1_idx:col2_idx) {
          referenced[[target_sheet]]$whole_cols <- c(
            referenced[[target_sheet]]$whole_cols,
            index_to_col_letter(ci)
          )
        }
      } else {
        # Normal range: expand all cells within rectangle
        expanded <- expand_range_to_cells(cell_part)
        referenced[[target_sheet]]$cells <- c(referenced[[target_sheet]]$cells, expanded)
      }
    }

    # --- Step 3: Extract single-cell references from placeholder formula ---
    # (Ranges have been replaced by <RANGE_N> placeholders)

    work_formula <- placeholder_formula

    # Handle cross-sheet single refs: 'Sheet Name'!A1
    quoted_cross <- regmatches(
      work_formula,
      gregexpr("'[^']+'!\\$?[A-Z]{1,3}\\$?[0-9]+", work_formula)
    )[[1]]
    for (ref in quoted_cross) {
      parts <- strsplit(ref, "!", fixed = FALSE)[[1]]
      sheet_part <- gsub("^'+|'+$", "", parts[1])
      target_sheet <- formula_sheet
      for (s in sheet_names) {
        if (s == sheet_part || sanitize_sheet_name(s) == sanitize_sheet_name(sheet_part)) {
          target_sheet <- s
          break
        }
      }
      cell_addr <- gsub("\\$", "", parts[2])
      referenced[[target_sheet]]$cells <- c(referenced[[target_sheet]]$cells, cell_addr)
    }
    # Remove matched cross-sheet refs from formula
    work_formula <- gsub("'[^']+'!\\$?[A-Z]{1,3}\\$?[0-9]+", "", work_formula)

    # Remove placeholders
    work_formula <- gsub("<RANGE_[0-9]+>", "", work_formula)

    # Remove string literals to avoid false matches
    work_formula <- gsub('"[^"]*"', "", work_formula)

    # Find remaining same-sheet single cell references
    single_matches <- regmatches(
      work_formula,
      gregexpr("[A-Z]{1,3}\\$?[0-9]+", work_formula)
    )[[1]]

    for (sm in single_matches) {
      sm_clean <- gsub("\\$", "", sm)
      # Validate: looks like a real cell ref (1-3 letters + digits)
      if (grepl("^[A-Z]{1,3}[0-9]+$", sm_clean)) {
        referenced[[formula_sheet]]$cells <- c(referenced[[formula_sheet]]$cells, sm_clean)
      }
    }
  }

  # Deduplicate per sheet
  for (s in sheet_names) {
    referenced[[s]]$cells <- unique(referenced[[s]]$cells)
    referenced[[s]]$whole_cols <- unique(referenced[[s]]$whole_cols)
  }

  referenced
}

#' @keywords internal
#' @noRd
export_sheet_csvs <- function(file_path, sheet_names, output_dir,
                               formula_data = NULL, sheet_dims = NULL) {
  dir.create(output_dir, showWarnings = FALSE, recursive = TRUE)

  all_cells <- tidyxl::xlsx_cells(file_path, sheets = sheet_names)

  # Build set of referenced cells if formula_data provided
  ref_sets <- NULL
  if (!is.null(formula_data) && nrow(formula_data) > 0) {
    if (is.null(sheet_dims)) {
      sheet_dims <- detect_sheet_dimensions(file_path, sheet_names)
    }
    ref_sets <- extract_referenced_cells(formula_data, sheet_names, sheet_dims)
  }

  paths <- character(0)

  for (s in sheet_names) {
    sheet_cells <- all_cells[all_cells$sheet == s, ]

    # Keep only non-formula cells
    value_cells <- sheet_cells[is.na(sheet_cells$formula), ]

    # Build address for each cell
    cell_addrs <- paste0(
      vapply(value_cells$col, index_to_col_letter, character(1)),
      value_cells$row
    )

    # Filter to referenced cells if we have that info
    sheet_ref <- NULL
    sheet_ref_set <- NULL
    if (!is.null(ref_sets)) {
      sheet_ref <- ref_sets[[s]]
      if (is.null(sheet_ref)) sheet_ref <- list(cells = character(0), whole_cols = character(0))

      # A cell is included if it's in the explicit cell set OR in a whole-column
      cell_cols <- vapply(value_cells$col, index_to_col_letter, character(1))
      keep_mask <- (cell_addrs %in% sheet_ref$cells) |
                   (cell_cols %in% sheet_ref$whole_cols)

      value_cells <- value_cells[keep_mask, ]
      cell_addrs <- cell_addrs[keep_mask]
      sheet_ref_set <- sheet_ref$cells  # for blank-slot logic below
    }

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

      if (is.na(val)) {
        # If this cell is explicitly referenced but blank, include as fillable slot
        if (!is.null(sheet_ref) && cell_addrs[i] %in% sheet_ref$cells) {
          rows_vec[i] <- cell$row
          cols_vec[i] <- index_to_col_letter(cell$col)
          vals_vec[i] <- ""
          keep[i] <- TRUE
        }
        # Skip blank cells that only matched via whole-column (no need for blank slots)
        next
      }

      rows_vec[i] <- cell$row
      cols_vec[i] <- index_to_col_letter(cell$col)
      vals_vec[i] <- val
      keep[i] <- TRUE
    }

    # Also add referenced cells that don't exist at all in tidyxl
    # (completely empty cells -- not even in all_cells)
    if (!is.null(sheet_ref_set) && length(sheet_ref_set) > 0) {
      existing_addrs <- cell_addrs
      # Also exclude formula cells from the "missing" set
      formula_cells <- sheet_cells[!is.na(sheet_cells$formula), ]
      formula_addrs <- paste0(
        vapply(formula_cells$col, index_to_col_letter, character(1)),
        formula_cells$row
      )
      all_existing <- c(existing_addrs, formula_addrs)
      missing <- setdiff(sheet_ref_set, all_existing)
      if (length(missing) > 0) {
        for (addr in missing) {
          pa <- parse_cell_address(addr)
          if (!is.na(pa$row)) {
            rows_vec <- c(rows_vec, pa$row)
            cols_vec <- c(cols_vec, pa$col)
            vals_vec <- c(vals_vec, "")
            keep <- c(keep, TRUE)
          }
        }
      }
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
