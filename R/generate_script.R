# =============================================================================
# generate_script.R -- Assembles the downloadable .R output script
# =============================================================================

#' @keywords internal
#' @noRd
process_excel_file <- function(file_path,
                                sheet_names = NULL,
                                wrap_trycatch = TRUE,
                                include_comments = TRUE,
                                include_named_tables = TRUE,
                                data_source = "excel",
                                excel_path_in_script = NULL,
                                progress_callback = NULL,
                                input_cells = NULL) {
  # input_cells: "Sheet!A1" cells whose formula is NOT translated; their value
  # is read from the workbook at run time (cells a caller overwrites with
  # its own values). Formulas that reference another workbook ('[1]Sheet'!A1)
  # are treated the same way, since that workbook is not available.

  update_progress <- function(step, detail) {
    if (!is.null(progress_callback)) progress_callback(step, detail)
  }

  all_warnings <- character(0)

  # --- Step 1: Extract formulas ---
  update_progress(1, "Extracting formulas from Excel file...")
  formula_data <- extract_all_formulas(file_path, sheet_names)

  # Formulas kept out of the translation (become inputs)
  key <- paste0(formula_data$Sheet, "!", formula_data$Cell)
  external <- grepl("\\[[0-9]+\\][^!]*!", formula_data$Formula)
  declared <- key %in% gsub("$", "", input_cells, fixed = TRUE)
  inputs <- formula_data[external | declared, c("Sheet", "Cell", "Formula")]
  if (nrow(inputs) > 0) {
    inputs$Reason <- ifelse(external[external | declared], "external link", "input cell")
    formula_data <- formula_data[!(external | declared), ]
    rownames(formula_data) <- NULL
    if (any(external)) {
      all_warnings <- c(all_warnings, sprintf(
        "%d formula(s) reference another workbook and are read as inputs (e.g. %s)",
        sum(external), key[which(external)[1]]))
    }
  }

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
  n_f <- nrow(report)
  r_code <- character(n_f)
  status <- character(n_f)
  warns <- character(n_f)

  # Columns used as *IF criteria get a text copy; criteria arguments are
  # pointed at it during the transform (see to_text_shadow)
  criteria_cols <- text_columns(formula_data, sheet_names)
  criteria_cols <- criteria_cols[lengths(criteria_cols) > 0]
  .excel2r_state$shadow <- setNames(criteria_cols, sanitize_sheet_name(names(criteria_cols)))
  on.exit(.excel2r_state$shadow <- NULL, add = TRUE)

  sheet_sanitized_all <- sanitize_sheet_name(report$Sheet)
  # Formulas as Excel evaluates them (implicit intersection applied)
  eval_formulas <- formulas_for_evaluation(formula_data)
  for (i in seq_len(n_f)) {
    res <- tryCatch({
      # Step 3a: Transform references; Step 3b: transform functions
      ref_transformed <- transform_cell_references(eval_formulas[i], sheet_sanitized_all[i], sheet_dims)
      transform_all_functions(ref_transformed)
    }, error = function(e) e)

    if (inherits(res, "error")) {
      r_code[i] <- sprintf("NA  # Transform error: %s", conditionMessage(res))
      status[i] <- "error"
      warns[i] <- conditionMessage(res)
      all_warnings <- c(all_warnings, sprintf("Error in %s!%s: %s",
                                              report$Sheet[i], report$Cell[i], conditionMessage(res)))
    } else {
      r_code[i] <- res$formula
      if (length(res$warnings) > 0) {
        status[i] <- "warning"
        warns[i] <- paste(res$warnings, collapse = "; ")
        all_warnings <- c(all_warnings, res$warnings)
      } else {
        status[i] <- "ok"
      }
    }
  }
  report$R_Code <- r_code
  report$Status <- status
  report$Warnings <- warns

  # --- Step 4: Determine execution order ---
  update_progress(4, "Analyzing dependencies and execution order...")
  exec_order <- determine_execution_order(formula_data, all_sheets = sheet_names)
  all_warnings <- c(all_warnings, exec_order$warnings)

  # Assign order numbers to report. Exec_Order is the global (cross-sheet)
  # position and is what the script follows; Sheet_Order/Cell_Order are kept
  # for display.
  sheet_order_map <- setNames(seq_along(exec_order$sheet_order), exec_order$sheet_order)
  report$Sheet_Order <- unname(sheet_order_map[report$Sheet])
  report$Sheet_Order[is.na(report$Sheet_Order)] <- 999L
  report$Exec_Order <- NA_integer_
  report$Exec_Order[exec_order$global_order] <- seq_along(exec_order$global_order)
  report$Cell_Order <- NA_integer_
  for (s in names(exec_order$cell_orders)) {
    rows <- which(report$Sheet == s)
    report$Cell_Order[rows] <- match(report$Cell[rows], exec_order$cell_orders[[s]])
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
       named_tables = named_tables,
       inputs = if (exists("inputs", inherits = FALSE) && nrow(inputs)) inputs else NULL)
}


#' @keywords internal
#' @noRd
generate_r_script <- function(excel_path, formula_data, sheet_names, sheet_dims,
                               exec_order, used_functions, wrap_trycatch, include_comments,
                               named_tables = NULL, data_source = "excel") {
  # Lines accumulator: preallocated buffer that doubles when full (appending
  # with c() copied the whole script on every line -- quadratic)
  L <- character(4096)
  n_L <- 0L
  push <- function(line) {
    if (n_L == length(L)) L <<- c(L, character(length(L)))
    n_L <<- n_L + 1L
    L[n_L] <<- line
  }
  add <- function(...) push(paste0(...))
  blank <- function() push("")

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
    add("# Set excel_file before sourcing to run on another workbook of this layout")
    # encodeString escapes backslashes (Windows paths) and quotes
    add('if (!exists("excel_file", inherits = FALSE)) excel_file <- ',
        encodeString(excel_path, quote = '"'))
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
    add("  # Plain value -- try numeric first")
    add("  num <- suppressWarnings(as.numeric(s))")
    add("  if (!is.na(num)) return(list(op = '==', val = num))")
    add("  list(op = '==', val = s)")
    add("}")
    blank()

    add("# Apply a parsed criterion to a range, returns logical vector")
    add("# Excel semantics: text criteria match case-insensitively; numeric")
    add("# criteria compare numerically (text that parses as a number counts);")
    add("# cells that cannot be compared never match, except under '<>'.")
    add(".apply_criterion <- function(range_vec, parsed) {")
    add("  if (identical(parsed$op, 'TRUE')) return(rep(TRUE, length(range_vec)))")
    add("  if (is.character(parsed$val)) {")
    add("    res <- get(parsed$op)(toupper(as.character(range_vec)), toupper(parsed$val))")
    add("  } else {")
    add("    rv <- if (is.numeric(range_vec)) range_vec else suppressWarnings(as.numeric(as.character(range_vec)))")
    add("    res <- get(parsed$op)(rv, parsed$val)")
    add("  }")
    add("  res[is.na(res)] <- identical(parsed$op, '!=')")
    add("  res")
    add("}")
    blank()
    add("# Sum that tolerates character columns (text is ignored, as in Excel)")
    add(".sum_num <- function(x) {")
    add("  if (!is.numeric(x)) x <- suppressWarnings(as.numeric(as.character(x)))")
    add("  sum(x, na.rm = TRUE)")
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
    add("  .sum_num(sum_range[mask])")
    add("}")
    blank()
  }

  if (uses_sumif || any(toupper(names(used_functions)) == "SUMIF")) {
    add("# SUMIF: criteria_range, criteria, sum_range")
    add("SUMIF <- function(criteria_range, criteria, sum_range = criteria_range) {")
    add("  parsed <- .parse_criterion(criteria)")
    add("  mask <- .apply_criterion(criteria_range, parsed)")
    add("  .sum_num(sum_range[mask])")
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

  # Columns with a text copy: *IF criteria + what pure references feed them
  criteria_cols <- text_columns(formula_data, sheet_names)
  pure_refs <- pure_reference_formulas(formula_data, sheet_names)
  pure_refs$key <- if (nrow(pure_refs)) {
    paste0(formula_data$Sheet[pure_refs$i], "!", formula_data$Cell[pure_refs$i])
  } else character(0)

  # Load every sheet, not only those holding formulas: data-only sheets
  # (inputs) are referenced by formulas elsewhere.
  load_sheets <- unique(c(exec_order$sheet_order, sheet_names))
  for (s in load_sheets) {
    s_sanitized <- sanitize_sheet_name(s)
    dims <- sheet_dims[[s]]
    if (is.null(dims) || dims$max_row < 1) next

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

    # Pad rows if needed
    add("if (nrow(", s_sanitized, ") < ", max_row, ") {")
    add("  .padding <- data.frame(matrix(NA, nrow = ", max_row, " - nrow(", s_sanitized, "), ncol = ncol(", s_sanitized, ")))")
    add("  colnames(.padding) <- colnames(", s_sanitized, ")")
    add("  ", s_sanitized, " <- rbind(", s_sanitized, ", .padding)")
    add("}")

    # Text copy of the columns used as *IF criteria, so text keys ("1.a",
    # "55.10") still match after the numeric conversion below
    sheet_crit_cols <- criteria_cols[[s]]
    has_shadow <- length(sheet_crit_cols) > 0
    if (has_shadow) {
      add("# Text copy of criteria columns (read by SUMIF/COUNTIF criteria)")
      add(s_sanitized, "__txt <- as.data.frame(lapply(", s_sanitized,
          ", as.character), stringsAsFactors = FALSE)")
    }

    add("# Convert all columns to numeric")
    add("# Blank/text cells become 0 (matches Excel behavior)")
    add("for (.col in colnames(", s_sanitized, ")) {")
    add("  ", s_sanitized, "[[.col]] <- suppressWarnings(as.numeric(as.character(", s_sanitized, "[[.col]])))")
    add("  ", s_sanitized, "[[.col]][is.na(", s_sanitized, "[[.col]])] <- 0")
    add("}")

    # Excel mode reads the workbook with its cached formula results. Clear
    # them so every formula is recomputed from inputs only: a cell computed
    # out of order then shows up as a mismatch instead of silently reusing
    # the cached value. (CSV mode never contains formula cells.)
    if (data_source != "csv") {
      f_rows <- formula_data[formula_data$Sheet == s, c("Cell", "Row", "Col")]
      if (nrow(f_rows) > 0) {
        f_letters <- gsub("[0-9$]", "", toupper(f_rows$Cell))
        f_rownum <- if (!is.null(f_rows$Row)) as.integer(f_rows$Row) else
          as.integer(gsub("[^0-9]", "", f_rows$Cell))
        add("# Clear cached formula results (recomputed below)")
        for (cl in unique(f_letters)) {
          rr <- sort(unique(f_rownum[f_letters == cl]))
          add(s_sanitized, "$", cl, "[", compact_int_vector(rr), "] <- NA")
          if (has_shadow && cl %in% sheet_crit_cols) {
            add(s_sanitized, "__txt$", cl, "[", compact_int_vector(rr), "] <- NA")
          }
        }
      }
    }
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
  sorted_formulas <- if (!is.null(formula_data$Exec_Order)) formula_data[order(formula_data$Exec_Order), ] else formula_data[order(formula_data$Sheet_Order, formula_data$Cell_Order), ]

  current_sheet <- ""
  for (i in seq_len(nrow(sorted_formulas))) {
    row <- sorted_formulas[i, ]
    s_sanitized <- sanitize_sheet_name(row$Sheet)

    if (row$Sheet != current_sheet) {
      blank()
      add("# --- Sheet: ", row$Sheet, " ---")
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

    # Keep the text copy in step when a formula writes into a text column;
    # a pure reference copies the source's text (codes survive), anything
    # else stores its computed value as text
    if (cell_parsed$col %in% criteria_cols[[row$Sheet]]) {
      txt_target <- sprintf("%s__txt$%s[%d]", s_sanitized, cell_parsed$col, cell_parsed$row)
      src <- pure_refs[pure_refs$key == paste0(row$Sheet, "!", row$Cell), ]
      if (nrow(src) == 1 && src$src_col %in% criteria_cols[[src$src_sheet]]) {
        add(txt_target, " <- ", sprintf("%s__txt$%s[%d]",
            sanitize_sheet_name(src$src_sheet), src$src_col, src$src_row))
      } else {
        add(txt_target, " <- as.character(", target, ")")
      }
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

  paste(L[seq_len(n_L)], collapse = "\n")
}


#' @keywords internal
#' @noRd
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

#' Pure single-cell reference formulas (=Sheet!D16): target and source
#' @keywords internal
#' @noRd
pure_reference_formulas <- function(formula_data, sheet_names) {
  f <- trimws(sub("^=", "", formula_data$Formula))
  cand <- which(grepl("^('[^']+'|[A-Za-z_][A-Za-z0-9_.]*)?!?\\$?[A-Z]{1,3}\\$?[0-9]+$", f))
  out <- data.frame(i = integer(0), src_sheet = character(0), src_col = character(0),
                    src_row = integer(0), stringsAsFactors = FALSE)
  lookup <- setNames(sheet_names, toupper(sheet_names))
  for (i in cand) {
    r <- scan_references(f[i])
    if (nrow(r) != 1 || r$is_range || r$text != f[i]) next
    s <- if (is.na(r$sheet)) formula_data$Sheet[i] else unname(lookup[toupper(r$sheet)])
    if (is.na(s)) next
    out[nrow(out) + 1, ] <- list(i, s, index_to_col_letter(r$c1), r$r1)
  }
  out
}

#' Columns that need a text copy (`<Sheet>__txt`)
#'
#' The *IF criteria columns, plus -- transitively -- every column that a pure
#' reference formula copies into one of them: a criterion cell such as
#' F75 = T6_MAKE!D16 must carry the text code of D16 ("0210-0240"), which
#' the numeric frames turn into 0.
#' @keywords internal
#' @noRd
text_columns <- function(formula_data, sheet_names) {
  cols <- identify_criteria_columns(formula_data, sheet_names)
  refs <- pure_reference_formulas(formula_data, sheet_names)
  if (nrow(refs) == 0) return(cols)
  tgt_sheet <- formula_data$Sheet[refs$i]
  tgt_col <- gsub("[0-9$]", "", toupper(formula_data$Cell[refs$i]))
  repeat {
    in_text <- mapply(function(s, cl) cl %in% cols[[s]], tgt_sheet, tgt_col)
    add <- refs[in_text & !mapply(function(s, cl) cl %in% cols[[s]], refs$src_sheet, refs$src_col), ]
    if (nrow(add) == 0) break
    for (k in seq_len(nrow(add))) {
      s <- add$src_sheet[k]
      cols[[s]] <- unique(c(cols[[s]], add$src_col[k]))
    }
  }
  cols
}

#' @keywords internal
#' @noRd
identify_criteria_columns <- function(formula_data, sheet_names) {
  # Columns that must stay character so text keys (e.g. "1.a", "68.31") keep
  # matching: every criteria range of a *IF/*IFS call, plus the column of a
  # criterion argument when that argument is a plain cell reference.
  # Calls are found anywhere in the formula (nested, negated, summed).
  criteria_cols <- setNames(vector("list", length(sheet_names)), sheet_names)

  # Positions (1-based) of criteria-range and criterion arguments per function
  arg_roles <- function(fname, n_args) {
    switch(fname,
      SUMIF = , COUNTIF = , AVERAGEIF = list(ranges = 1L, crits = 2L),
      SUMIFS = , AVERAGEIFS = list(ranges = seq(2L, n_args, by = 2L),
                                  crits = seq(3L, n_args, by = 2L)),
      COUNTIFS = list(ranges = seq(1L, n_args, by = 2L),
                      crits = seq(2L, n_args, by = 2L)),
      NULL)
  }

  # Collect (sheet, column index) pairs; letters are resolved once at the end
  found_sheet <- character(0)
  found_col <- integer(0)
  add_col <- function(ref_sheet, c1, c2, formula_sheet) {
    target <- if (is.na(ref_sheet)) formula_sheet else ref_sheet
    if (!target %in% sheet_names) return()
    found_sheet <<- c(found_sheet, rep(target, abs(c2 - c1) + 1L))
    found_col <<- c(found_col, seq(c1, c2))
  }

  has_if <- grepl("(SUMIFS?|COUNTIFS?|AVERAGEIFS?)\\(", formula_data$Formula, ignore.case = TRUE)

  for (i in which(has_if)) {
    sheet <- formula_data$Sheet[i]
    formula <- formula_data$Formula[i]
    for (call in find_function_calls(formula)) {
      fname <- toupper(call$name)
      args <- split_function_args(call$args_str)
      roles <- arg_roles(fname, length(args))
      if (is.null(roles)) next
      want_ranges <- roles$ranges[roles$ranges <= length(args)]
      want_crits <- roles$crits[roles$crits <= length(args)]
      for (j in c(want_ranges, want_crits)) {
        refs <- scan_references(args[j])
        # A criterion only counts when it is exactly one reference
        if (j %in% want_crits && !(nrow(refs) == 1 && refs$text == trimws(args[j]))) next
        for (k in seq_len(nrow(refs))) add_col(refs$sheet[k], refs$c1[k], refs$c2[k], sheet)
      }
    }
  }

  pairs <- unique(data.frame(s = found_sheet, c = found_col, stringsAsFactors = FALSE))
  for (s in unique(pairs$s)) {
    cols <- sort(pairs$c[pairs$s == s])
    criteria_cols[[s]] <- vapply(cols, index_to_col_letter, character(1))
  }
  criteria_cols
}
