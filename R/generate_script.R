# =============================================================================
# generate_script.R — Assembles the downloadable .R output script
# =============================================================================

#' Main entry point: process an Excel file and generate an R script
#' @param file_path Path to the uploaded .xlsx file
#' @param sheet_names Optional: which sheets to include (NULL = all)
#' @param wrap_trycatch Wrap each formula in tryCatch? (default TRUE)
#' @param include_comments Include original Excel formulas as comments? (default TRUE)
#' @param excel_path_in_script Path to write in the generated script's read_xlsx() calls
#' @param progress_callback Optional function(step, detail) for Shiny progress
#' @return list(script = character string, report = data.frame, warnings = character vector)
process_excel_file <- function(file_path,
                                sheet_names = NULL,
                                wrap_trycatch = TRUE,
                                include_comments = TRUE,
                                include_named_tables = TRUE,
                                data_source = "excel",
                                excel_path_in_script = NULL,
                                progress_callback = NULL) {

  update_progress <- function(step, detail) {
    if (!is.null(progress_callback)) progress_callback(step, detail)
  }

  all_warnings <- character(0)

  # --- Step 1: Extract formulas ---
  update_progress(1, "Extracting formulas from Excel file...")
  formula_data <- extract_all_formulas(file_path, sheet_names)

  if (nrow(formula_data) == 0) {
    return(list(
      script = "# No formulas found in the uploaded Excel file.\n",
      report = data.frame(),
      warnings = "No formulas found.",
      named_tables = NULL
    ))
  }

  # Detect sheets from formulas if not specified
  if (is.null(sheet_names)) {
    sheet_names <- readxl::excel_sheets(file_path)
  }

  # --- Step 2: Detect dimensions ---
  update_progress(2, "Detecting sheet dimensions...")
  sheet_dims <- detect_sheet_dimensions(file_path, sheet_names)

  # Detect used functions
  used_functions <- detect_used_functions(formula_data)

  # Detect named tables
  named_tables <- NULL
  if (include_named_tables) {
    named_tables <- tryCatch(
      detect_named_tables(file_path, sheet_names),
      error = function(e) NULL
    )
    if (!is.null(named_tables) && nrow(named_tables) == 0) named_tables <- NULL
  }

  # --- Step 3: Transform formulas ---
  update_progress(3, "Transforming formulas to R code...")

  report <- formula_data
  report$R_Code <- character(nrow(report))
  report$Status <- character(nrow(report))
  report$Warnings <- character(nrow(report))

  for (i in seq_len(nrow(report))) {
    sheet_sanitized <- sanitize_sheet_name(report$Sheet[i])
    formula <- report$Formula[i]

    tryCatch({
      # Step 3a: Transform references
      ref_transformed <- transform_cell_references(formula, sheet_sanitized, sheet_dims)

      # Step 3b: Transform functions
      func_result <- transform_all_functions(ref_transformed)

      report$R_Code[i] <- func_result$formula
      if (length(func_result$warnings) > 0) {
        report$Status[i] <- "warning"
        report$Warnings[i] <- paste(func_result$warnings, collapse = "; ")
        all_warnings <- c(all_warnings, func_result$warnings)
      } else {
        report$Status[i] <- "ok"
      }
    }, error = function(e) {
      report$R_Code[i] <<- sprintf("NA  # Transform error: %s", e$message)
      report$Status[i] <<- "error"
      report$Warnings[i] <<- e$message
      all_warnings <<- c(all_warnings, sprintf("Error in %s!%s: %s",
                                                report$Sheet[i], report$Cell[i], e$message))
    })
  }

  # --- Step 4: Determine execution order ---
  update_progress(4, "Analyzing dependencies and execution order...")
  exec_order <- determine_execution_order(formula_data)
  all_warnings <- c(all_warnings, exec_order$warnings)

  # Assign order numbers to report
  sheet_order_map <- setNames(seq_along(exec_order$sheet_order), exec_order$sheet_order)
  report$Sheet_Order <- NA_integer_
  report$Cell_Order <- NA_integer_

  for (i in seq_len(nrow(report))) {
    s <- report$Sheet[i]
    report$Sheet_Order[i] <- if (!is.null(sheet_order_map[s])) sheet_order_map[s] else 999L

    cell_order <- exec_order$cell_orders[[s]]
    if (!is.null(cell_order)) {
      idx <- match(report$Cell[i], cell_order)
      report$Cell_Order[i] <- if (!is.na(idx)) idx else 999L
    }
  }

  # --- Step 5: Generate script ---
  update_progress(5, "Generating R script...")

  if (is.null(excel_path_in_script)) {
    excel_path_in_script <- basename(file_path)
  }

  script <- generate_r_script(
    excel_path = excel_path_in_script,
    formula_data = report,
    sheet_names = sheet_names,
    sheet_dims = sheet_dims,
    exec_order = exec_order,
    used_functions = used_functions,
    wrap_trycatch = wrap_trycatch,
    include_comments = include_comments,
    named_tables = named_tables,
    data_source = data_source
  )

  list(script = script, report = report, warnings = unique(all_warnings),
       named_tables = named_tables)
}


#' Generate the self-contained R script
generate_r_script <- function(excel_path, formula_data, sheet_names, sheet_dims,
                               exec_order, used_functions, wrap_trycatch, include_comments,
                               named_tables = NULL, data_source = "excel") {
  L <- character(0)  # lines accumulator
  add <- function(...) L <<- c(L, paste0(...))
  blank <- function() L <<- c(L, "")

  n_formulas <- nrow(formula_data)
  n_ok <- sum(formula_data$Status == "ok")
  n_warn <- sum(formula_data$Status == "warning")
  n_err <- sum(formula_data$Status == "error")

  # === SECTION 1: Header ===
  add("# ============================================================")
  add("# Auto-generated R script from Excel formulas")
  add("# Source file: ", excel_path)
  add("# Generated:  ", format(Sys.time(), "%Y-%m-%d %H:%M:%S"))
  add("# Sheets:     ", length(sheet_names))
  add("# Formulas:   ", n_formulas, " (", n_ok, " OK, ", n_warn, " warnings, ", n_err, " errors)")
  add("# ============================================================")
  blank()

  # === SECTION 2: Required Packages ===
  uses_ifs <- any(grepl("dplyr::case_when", formula_data$R_Code, fixed = TRUE))

  if (data_source == "csv") {
    if (uses_ifs) {
      add("# --- Required Packages ---")
      add('if (!requireNamespace("dplyr", quietly = TRUE)) install.packages("dplyr")')
      blank()
      add("library(dplyr)")
    } else {
      add("# --- No external packages required (base R only) ---")
    }
  } else {
    add("# --- Required Packages ---")
    add('if (!requireNamespace("openxlsx2", quietly = TRUE)) install.packages("openxlsx2")')
    if (uses_ifs) {
      add('if (!requireNamespace("dplyr", quietly = TRUE)) install.packages("dplyr")')
    }
    blank()
    add("library(openxlsx2)")
    if (uses_ifs) add("library(dplyr)")
  }
  blank()

  # === SECTION 3: Configuration ===
  add("# --- Configuration ---")
  if (data_source == "csv") {
    add('data_dir <- "data"')
  } else {
    add('excel_file <- "', excel_path, '"')
  }
  blank()

  # === SECTION 4: Helper Functions ===
  add("# --- Helper Functions ---")
  blank()

  # Always include basic helpers
  add("# Column letter to index (A=1, Z=26, AA=27, ...)")
  add("col_letter_to_index <- function(col) {")
  add("  col <- toupper(col)")
  add("  chars <- strsplit(col, '')[[1]]")
  add("  idx <- 0")
  add("  for (ch in chars) idx <- idx * 26 + match(ch, LETTERS)")
  add("  idx")
  add("}")
  blank()

  add("# Index to column letter")
  add("index_to_col_letter <- function(idx) {")
  add("  result <- ''")
  add("  while (idx > 0) {")
  add("    remainder <- (idx - 1) %% 26")
  add("    result <- paste0(LETTERS[remainder + 1], result)")
  add("    idx <- (idx - 1) %/% 26")
  add("  }")
  add("  result")
  add("}")
  blank()

  add("# Generate column names")
  add("generate_col_names <- function(n) vapply(1:n, index_to_col_letter, character(1))")
  blank()

  if (data_source == "csv") {
    add("# Reconstruct sheet grid from tidy CSV (row, col, value)")
    add("reconstruct_grid <- function(csv_path, max_row, max_col) {")
    add("  raw <- read.csv(csv_path, stringsAsFactors = FALSE, colClasses = 'character')")
    add("  col_names <- generate_col_names(max_col)")
    add("  grid <- data.frame(matrix(NA, nrow = max_row, ncol = max_col))")
    add("  colnames(grid) <- col_names")
    add("  for (i in seq_len(nrow(raw))) {")
    add("    r <- as.integer(raw$row[i])")
    add("    c <- raw$col[i]")
    add("    if (r <= max_row && c %in% col_names) grid[[c]][r] <- raw$value[i]")
    add("  }")
    add("  grid")
    add("}")
    blank()
  }

  # Lookup helpers if needed
  uses_vlookup <- any(grepl("excel_vlookup", formula_data$R_Code, fixed = TRUE))
  uses_hlookup <- any(grepl("excel_hlookup", formula_data$R_Code, fixed = TRUE))
  uses_match <- any(grepl("excel_match", formula_data$R_Code, fixed = TRUE))
  uses_index <- any(grepl("excel_index", formula_data$R_Code, fixed = TRUE))
  uses_xlookup <- any(grepl("excel_xlookup", formula_data$R_Code, fixed = TRUE))
  uses_roundup <- any(grepl("excel_roundup", formula_data$R_Code, fixed = TRUE))
  uses_rounddown <- any(grepl("excel_rounddown", formula_data$R_Code, fixed = TRUE))

  if (uses_vlookup) {
    add("# VLOOKUP equivalent")
    add("excel_vlookup <- function(lookup_val, table_range, col_idx, exact = FALSE) {")
    add("  if (!(is.data.frame(table_range) || is.matrix(table_range))) return(NA)")
    add("  first_col <- table_range[, 1]")
    add("  if (isTRUE(exact) || identical(exact, 0)) {")
    add("    # Exact match")
    add("    idx <- match(lookup_val, first_col)")
    add("  } else {")
    add("    # Approximate match: sorted ascending, find largest <= lookup_val")
    add("    idx <- match(lookup_val, first_col)")
    add("    if (is.na(idx)) {")
    add("      candidates <- which(first_col <= lookup_val)")
    add("      idx <- if (length(candidates) > 0) max(candidates) else NA")
    add("    }")
    add("  }")
    add("  if (is.na(idx)) return(NA)")
    add("  table_range[idx, col_idx]")
    add("}")
    blank()
  }

  if (uses_hlookup) {
    add("# HLOOKUP equivalent")
    add("excel_hlookup <- function(lookup_val, table_range, row_idx, exact = FALSE) {")
    add("  if (is.data.frame(table_range) || is.matrix(table_range)) {")
    add("    idx <- match(lookup_val, table_range[1, ])")
    add("    if (is.na(idx)) return(NA)")
    add("    return(table_range[row_idx, idx])")
    add("  }")
    add("  NA")
    add("}")
    blank()
  }

  if (uses_match) {
    add("# MATCH equivalent")
    add("excel_match <- function(lookup_val, lookup_range, match_type = 0) {")
    add("  if (match_type == 0) {")
    add("    idx <- match(lookup_val, lookup_range)")
    add("    return(if (is.na(idx)) NA else idx)")
    add("  }")
    add("  # match_type 1 or -1: approximate match")
    add("  idx <- match(lookup_val, lookup_range)")
    add("  if (!is.na(idx)) return(idx)")
    add("  if (match_type == 1) {")
    add("    candidates <- which(lookup_range <= lookup_val)")
    add("    return(if (length(candidates) == 0) NA else max(candidates))")
    add("  } else {")
    add("    candidates <- which(lookup_range >= lookup_val)")
    add("    return(if (length(candidates) == 0) NA else min(candidates))")
    add("  }")
    add("}")
    blank()
  }

  if (uses_index) {
    add("# INDEX equivalent")
    add("excel_index <- function(array_range, row_num, col_num = NULL) {")
    add("  if (is.null(col_num)) {")
    add("    return(array_range[row_num])")
    add("  }")
    add("  if (is.data.frame(array_range) || is.matrix(array_range)) {")
    add("    return(array_range[row_num, col_num])")
    add("  }")
    add("  array_range[row_num]")
    add("}")
    blank()
  }

  if (uses_xlookup) {
    add("# XLOOKUP equivalent")
    add("excel_xlookup <- function(lookup_val, lookup_array, return_array, if_not_found = NA) {")
    add("  idx <- match(lookup_val, lookup_array)")
    add("  if (is.na(idx)) return(if_not_found)")
    add("  return_array[idx]")
    add("}")
    blank()
  }

  if (uses_roundup) {
    add("# ROUNDUP equivalent")
    add("excel_roundup <- function(x, digits = 0) {")
    add("  mult <- 10^digits")
    add("  ceiling(x * mult) / mult")
    add("}")
    blank()
  }

  if (uses_rounddown) {
    add("# ROUNDDOWN equivalent")
    add("excel_rounddown <- function(x, digits = 0) {")
    add("  mult <- 10^digits")
    add("  floor(x * mult) / mult")
    add("}")
    blank()
  }

  # Conditional aggregation helpers (based on Excel operator parsing)
  # Shared criteria parser needed by SUMIF/SUMIFS/COUNTIF/COUNTIFS/AVERAGEIF/AVERAGEIFS
  cond_agg_funcs <- c("SUMIF", "SUMIFS", "COUNTIF", "COUNTIFS", "AVERAGEIF", "AVERAGEIFS")
  uses_cond_agg <- any(toupper(names(used_functions)) %in% cond_agg_funcs)

  if (uses_cond_agg) {
    add("# Parse Excel criteria string into operator + value")
    add("# Handles: \">=10\", \">5\", \"<>0\", \"YES\", 42, etc.")
    add(".parse_criterion <- function(crit) {")
    add("  if (is.numeric(crit)) return(list(op = '==', val = crit))")
    add("  if (is.logical(crit) && isTRUE(crit)) return(list(op = 'TRUE', val = TRUE))")
    add("  s <- as.character(crit)")
    add("  if (grepl('^<>', s)) return(list(op = '!=', val = type.convert(sub('^<>', '', s), as.is = TRUE)))")
    add("  if (grepl('^>=', s)) return(list(op = '>=', val = as.numeric(sub('^>=', '', s))))")
    add("  if (grepl('^<=', s)) return(list(op = '<=', val = as.numeric(sub('^<=', '', s))))")
    add("  if (grepl('^>', s))  return(list(op = '>',  val = as.numeric(sub('^>', '', s))))")
    add("  if (grepl('^<', s))  return(list(op = '<',  val = as.numeric(sub('^<', '', s))))")
    add("  # Plain value — try numeric first")
    add("  num <- suppressWarnings(as.numeric(s))")
    add("  if (!is.na(num)) return(list(op = '==', val = num))")
    add("  list(op = '==', val = s)")
    add("}")
    blank()

    add("# Apply a parsed criterion to a range, returns logical vector")
    add(".apply_criterion <- function(range_vec, parsed) {")
    add("  if (identical(parsed$op, 'TRUE')) return(rep(TRUE, length(range_vec)))")
    add("  get(parsed$op)(range_vec, parsed$val)")
    add("}")
    blank()
  }

  uses_sumifs <- any(grepl("\\bSUMIFS\\b", names(used_functions)))
  uses_sumif <- any(grepl("\\bSUMIF\\b", names(used_functions))) && !uses_sumifs
  uses_countifs <- any(grepl("\\bCOUNTIFS\\b", names(used_functions)))
  uses_countif <- any(grepl("\\bCOUNTIF\\b", names(used_functions))) && !uses_countifs
  uses_averageifs <- any(grepl("\\bAVERAGEIFS\\b", names(used_functions)))
  uses_averageif <- any(grepl("\\bAVERAGEIF\\b", names(used_functions))) && !uses_averageifs

  if (uses_sumifs || any(toupper(names(used_functions)) == "SUMIFS")) {
    add("# SUMIFS: sum_range, criteria_range1, criteria1, criteria_range2, criteria2, ...")
    add("SUMIFS <- function(sum_range, ...) {")
    add("  args <- list(...)")
    add("  mask <- rep(TRUE, length(sum_range))")
    add("  for (i in seq(1, length(args), by = 2)) {")
    add("    crit_range <- args[[i]]")
    add("    crit_val <- args[[i + 1]]")
    add("    parsed <- .parse_criterion(crit_val)")
    add("    mask <- mask & .apply_criterion(crit_range, parsed)")
    add("  }")
    add("  sum(sum_range[mask], na.rm = TRUE)")
    add("}")
    blank()
  }

  if (uses_sumif || any(toupper(names(used_functions)) == "SUMIF")) {
    add("# SUMIF: criteria_range, criteria, sum_range")
    add("SUMIF <- function(criteria_range, criteria, sum_range) {")
    add("  parsed <- .parse_criterion(criteria)")
    add("  mask <- .apply_criterion(criteria_range, parsed)")
    add("  sum(sum_range[mask], na.rm = TRUE)")
    add("}")
    blank()
  }

  if (uses_countifs || any(toupper(names(used_functions)) == "COUNTIFS")) {
    add("# COUNTIFS: criteria_range1, criteria1, criteria_range2, criteria2, ...")
    add("COUNTIFS <- function(...) {")
    add("  args <- list(...)")
    add("  n <- length(args[[1]])")
    add("  mask <- rep(TRUE, n)")
    add("  for (i in seq(1, length(args), by = 2)) {")
    add("    crit_range <- args[[i]]")
    add("    crit_val <- args[[i + 1]]")
    add("    parsed <- .parse_criterion(crit_val)")
    add("    mask <- mask & .apply_criterion(crit_range, parsed)")
    add("  }")
    add("  sum(mask, na.rm = TRUE)")
    add("}")
    blank()
  }

  if (uses_countif || any(toupper(names(used_functions)) == "COUNTIF")) {
    add("# COUNTIF: criteria_range, criteria")
    add("COUNTIF <- function(criteria_range, criteria) {")
    add("  parsed <- .parse_criterion(criteria)")
    add("  mask <- .apply_criterion(criteria_range, parsed)")
    add("  sum(mask, na.rm = TRUE)")
    add("}")
    blank()
  }

  if (uses_averageifs || any(toupper(names(used_functions)) == "AVERAGEIFS")) {
    add("# AVERAGEIFS: avg_range, criteria_range1, criteria1, criteria_range2, criteria2, ...")
    add("AVERAGEIFS <- function(avg_range, ...) {")
    add("  args <- list(...)")
    add("  mask <- rep(TRUE, length(avg_range))")
    add("  for (i in seq(1, length(args), by = 2)) {")
    add("    crit_range <- args[[i]]")
    add("    crit_val <- args[[i + 1]]")
    add("    parsed <- .parse_criterion(crit_val)")
    add("    mask <- mask & .apply_criterion(crit_range, parsed)")
    add("  }")
    add("  mean(avg_range[mask], na.rm = TRUE)")
    add("}")
    blank()
  }

  if (uses_averageif || any(toupper(names(used_functions)) == "AVERAGEIF")) {
    add("# AVERAGEIF: criteria_range, criteria, avg_range")
    add("AVERAGEIF <- function(criteria_range, criteria, avg_range) {")
    add("  parsed <- .parse_criterion(criteria)")
    add("  mask <- .apply_criterion(criteria_range, parsed)")
    add("  mean(avg_range[mask], na.rm = TRUE)")
    add("}")
    blank()
  }

  # === SECTION 5: Load Data ===
  add("# ============================================================")
  if (data_source == "csv") {
    add("# Load Data from CSV files")
  } else {
    add("# Load Excel Data")
  }
  add("# ============================================================")
  blank()

  # Identify criteria columns per sheet (from SUMIF/SUMIFS formulas)
  criteria_cols <- identify_criteria_columns(formula_data, sheet_names)

  for (s in exec_order$sheet_order) {
    s_sanitized <- sanitize_sheet_name(s)
    dims <- sheet_dims[[s]]
    if (is.null(dims)) next

    max_row <- dims$max_row
    max_col <- dims$max_col

    add("# Sheet: \"", s, "\" -> ", s_sanitized)

    if (data_source == "csv") {
      csv_filename <- paste0(s_sanitized, ".csv")
      add(s_sanitized, " <- reconstruct_grid(")
      add("  file.path(data_dir, \"", csv_filename, "\"),")
      add("  max_row = ", max_row, ", max_col = ", max_col)
      add(")")
    } else {
      add(s_sanitized, " <- as.data.frame(openxlsx2::read_xlsx(")
      add("  excel_file, sheet = \"", s, "\",")
      add("  rows = 1:", max_row, ", cols = 1:", max_col, ",")
      add("  skip_empty_rows = FALSE, skip_empty_cols = FALSE, col_names = FALSE")
      add("))")
      blank()

      add("# Ensure correct column count and names")
      add("if (ncol(", s_sanitized, ") < ", max_col, ") {")
      add("  ", s_sanitized, "[(ncol(", s_sanitized, ") + 1):", max_col, "] <- NA")
      add("}")
      add(s_sanitized, " <- ", s_sanitized, "[, 1:", max_col, ", drop = FALSE]")
      add("colnames(", s_sanitized, ") <- generate_col_names(", max_col, ")")
    }
    blank()

    # Type conversion
    sheet_crit_cols <- criteria_cols[[s]]
    if (length(sheet_crit_cols) > 0) {
      crit_str <- paste0('"', sheet_crit_cols, '"', collapse = ", ")
      add("# Convert to numeric, preserving criteria columns as character")
      add("# Blank/text cells become 0 in numeric columns (matches Excel behavior)")
      add("for (.col in colnames(", s_sanitized, ")) {")
      add("  if (.col %in% c(", crit_str, ")) {")
      add("    ", s_sanitized, "[[.col]] <- as.character(", s_sanitized, "[[.col]])")
      add("  } else {")
      add("    ", s_sanitized, "[[.col]] <- suppressWarnings(as.numeric(as.character(", s_sanitized, "[[.col]])))")
      add("    ", s_sanitized, "[[.col]][is.na(", s_sanitized, "[[.col]])] <- 0")
      add("  }")
      add("}")
    } else {
      add("# Convert all columns to numeric")
      add("# Blank/text cells become 0 (matches Excel behavior)")
      add("for (.col in colnames(", s_sanitized, ")) {")
      add("  ", s_sanitized, "[[.col]] <- suppressWarnings(as.numeric(as.character(", s_sanitized, "[[.col]])))")
      add("  ", s_sanitized, "[[.col]][is.na(", s_sanitized, "[[.col]])] <- 0")
      add("}")
    }

    # Pad rows if needed
    add("if (nrow(", s_sanitized, ") < ", max_row, ") {")
    add("  .padding <- data.frame(matrix(NA, nrow = ", max_row, " - nrow(", s_sanitized, "), ncol = ncol(", s_sanitized, ")))")
    add("  colnames(.padding) <- colnames(", s_sanitized, ")")
    add("  ", s_sanitized, " <- rbind(", s_sanitized, ", .padding)")
    add("}")
    blank()
  }

  # === SECTION 5.5: Named Tables ===
  if (!is.null(named_tables) && nrow(named_tables) > 0) {
    add("# ============================================================")
    add("# Named Tables")
    add("# ============================================================")
    blank()

    for (i in seq_len(nrow(named_tables))) {
      tbl <- named_tables[i, ]
      s_sanitized <- sanitize_sheet_name(tbl$sheet)
      tbl_name <- sanitize_sheet_name(tbl$table_name)
      col_start_idx <- col_letter_to_index(tbl$col_start)
      col_end_idx <- col_letter_to_index(tbl$col_end)
      col_names <- tbl$col_names[[1]]

      add("# Table \"", tbl$table_name, "\" on sheet \"", tbl$sheet, "\" (", tbl$ref, ")")
      add(tbl_name, " <- ", s_sanitized, "[", tbl$data_start_row, ":", tbl$data_end_row,
          ", ", col_start_idx, ":", col_end_idx, "]")
      col_names_str <- paste0('"', col_names, '"', collapse = ", ")
      add("colnames(", tbl_name, ") <- c(", col_names_str, ")")
      blank()
    }
  }

  # === SECTION 6: Apply Formulas ===
  add("# ============================================================")
  add("# Apply Formulas (dependency-ordered)")
  add("# ============================================================")
  blank()

  # Sort by execution order
  sorted_formulas <- formula_data[order(formula_data$Sheet_Order, formula_data$Cell_Order), ]

  current_sheet <- ""
  for (i in seq_len(nrow(sorted_formulas))) {
    row <- sorted_formulas[i, ]
    s_sanitized <- sanitize_sheet_name(row$Sheet)

    if (row$Sheet != current_sheet) {
      blank()
      add("# --- Sheet: ", row$Sheet, " (Order: ", row$Sheet_Order, ") ---")
      current_sheet <- row$Sheet
    }

    cell_parsed <- parse_cell_address(row$Cell)
    target <- sprintf("%s$%s[%d]", s_sanitized, cell_parsed$col, cell_parsed$row)

    if (include_comments) {
      add("# ", row$Cell, " = ", row$Formula)
    }

    if (row$Status == "error") {
      add(target, " <- NA  # Transform error: ", row$Warnings)
    } else if (wrap_trycatch) {
      add(target, " <- tryCatch(")
      add("  ", row$R_Code, ",")
      add("  error = function(e) { message('Error in ", row$Sheet, "!", row$Cell, ": ', e$message); NA }")
      add(")")
    } else {
      add(target, " <- ", row$R_Code)
    }
  }

  # === SECTION 7: Summary ===
  blank()
  add("# ============================================================")
  add("# Verification Summary")
  add("# ============================================================")
  blank()
  add('cat("\\n=== Script execution complete ===\\n")')
  add('cat("Data frames created:\\n")')

  for (s in exec_order$sheet_order) {
    s_sanitized <- sanitize_sheet_name(s)
    add('cat(sprintf("  ', s_sanitized, ': %d rows x %d cols\\n", nrow(', s_sanitized, '), ncol(', s_sanitized, ')))')
  }

  if (!is.null(named_tables) && nrow(named_tables) > 0) {
    blank()
    add('cat("\\nNamed tables:\\n")')
    for (i in seq_len(nrow(named_tables))) {
      tbl_name <- sanitize_sheet_name(named_tables$table_name[i])
      add('cat(sprintf("  ', tbl_name, ': %d rows x %d cols\\n", nrow(', tbl_name, '), ncol(', tbl_name, ')))')
    }
  }

  paste(L, collapse = "\n")
}


#' Resolve which sheet a range reference targets
#' @param rng Range string, possibly with sheet prefix like 'Sheet1'!A:A or Sheet1!A:A
#' @param default_sheet The sheet the formula lives in (fallback)
#' @param sheet_names All known sheet names
#' @return The target sheet name (original, not sanitized)
resolve_range_sheet <- function(rng, default_sheet, sheet_names) {
  if (!grepl("!", rng)) return(default_sheet)
  sheet_part <- strsplit(rng, "!")[[1]][1]
  # Strip quotes
  sheet_part <- gsub("^'+|'+$", "", sheet_part)
  # Match against known sheet names
  if (sheet_part %in% sheet_names) return(sheet_part)
  # Try sanitized match
  for (s in sheet_names) {
    if (sanitize_sheet_name(s) == sanitize_sheet_name(sheet_part)) return(s)
  }
  default_sheet
}

#' Identify which columns are used as criteria in SUMIF/SUMIFS formulas
#' These must be kept as character type when loading data
identify_criteria_columns <- function(formula_data, sheet_names) {
  criteria_cols <- setNames(vector("list", length(sheet_names)), sheet_names)

  sumif_rows <- formula_data[grepl("^(SUMIFS?|COUNTIFS?|AVERAGEIFS?)", formula_data$Formula, ignore.case = TRUE), ]

  for (i in seq_len(nrow(sumif_rows))) {
    sheet <- sumif_rows$Sheet[i]
    formula <- sumif_rows$Formula[i]

    # Extract all ranges
    ranges <- regmatches(formula, gregexpr("(?:'[^']+'!)?[A-Z]+\\$?[0-9]*:[A-Z]+\\$?[0-9]*", formula))[[1]]

    # Filter out ranges inside quotes
    valid_ranges <- character(0)
    for (rng in ranges) {
      rng_pos <- regexpr(rng, formula, fixed = TRUE)[1]
      if (!is_within_quotes(formula, rng_pos)) {
        valid_ranges <- c(valid_ranges, rng)
      }
    }

    if (grepl("^SUMIF\\(", formula, ignore.case = TRUE) && length(valid_ranges) >= 1) {
      # SUMIF(range, criteria, sum_range) — range is the criteria range
      rng <- valid_ranges[1]
      target_sheet <- resolve_range_sheet(rng, sheet, sheet_names)
      rng_cell <- if (grepl("!", rng)) strsplit(rng, "!")[[1]][2] else rng
      col <- gsub("[0-9$:]+", "", strsplit(rng_cell, ":")[[1]][1])
      if (target_sheet %in% sheet_names) {
        criteria_cols[[target_sheet]] <- unique(c(criteria_cols[[target_sheet]], col))
      }
    } else if (grepl("^SUMIFS\\(", formula, ignore.case = TRUE) && length(valid_ranges) > 1) {
      # SUMIFS(sum_range, crit_range1, crit1, crit_range2, crit2, ...)
      for (j in 2:length(valid_ranges)) {
        rng <- valid_ranges[j]
        target_sheet <- resolve_range_sheet(rng, sheet, sheet_names)
        rng_cell <- if (grepl("!", rng)) strsplit(rng, "!")[[1]][2] else rng
        col <- gsub("[0-9$:]+", "", strsplit(rng_cell, ":")[[1]][1])
        if (target_sheet %in% sheet_names) {
          criteria_cols[[target_sheet]] <- unique(c(criteria_cols[[target_sheet]], col))
        }
      }
    } else if (grepl("^COUNTIFS?\\(", formula, ignore.case = TRUE)) {
      for (j in seq_along(valid_ranges)) {
        rng <- valid_ranges[j]
        target_sheet <- resolve_range_sheet(rng, sheet, sheet_names)
        rng_cell <- if (grepl("!", rng)) strsplit(rng, "!")[[1]][2] else rng
        col <- gsub("[0-9$:]+", "", strsplit(rng_cell, ":")[[1]][1])
        if (target_sheet %in% sheet_names) {
          criteria_cols[[target_sheet]] <- unique(c(criteria_cols[[target_sheet]], col))
        }
      }
    }
  }

  criteria_cols
}
