# =============================================================================
# migrate.R -- Public API for excel2r package
# =============================================================================

#' Migrate an Excel workbook to R
#'
#' Extracts all formulas, translates them to R code, exports raw data as
#' tidy CSVs, and generates a standalone R script. The Excel file is no
#' longer needed after migration.
#'
#' @param file_path Path to the .xlsx file
#' @param output_dir Directory to write output (created if needed)
#' @param mode Output mode: "csv" (standalone, default) or "excel" (requires .xlsx at runtime)
#' @param include_named_tables Detect and export named tables? (default TRUE)
#' @param wrap_trycatch Wrap formulas in tryCatch? (default TRUE)
#' @param include_comments Include original Excel formulas as comments? (default TRUE)
#' @param verify Run verification against Excel cached values? (default TRUE)
#' @return Invisibly returns a list with script, report, named_tables, verification results
#' @export
#'
#' @examples
#' \dontrun{
#' # Full standalone migration
#' excel2r::migrate("workbook.xlsx", "my_project/")
#'
#' # Excel-dependent mode (lighter, no CSVs)
#' excel2r::migrate("workbook.xlsx", "output/", mode = "excel")
#'
#' # Skip verification
#' excel2r::migrate("workbook.xlsx", "output/", verify = FALSE)
#' }
migrate <- function(file_path,
                    output_dir = ".",
                    mode = c("csv", "excel"),
                    include_named_tables = TRUE,
                    wrap_trycatch = TRUE,
                    include_comments = TRUE,
                    verify = TRUE) {
  mode <- match.arg(mode)

  # Process
  result <- process_excel_file(
    file_path = file_path,
    wrap_trycatch = wrap_trycatch,
    include_comments = include_comments,
    include_named_tables = include_named_tables,
    data_source = mode,
    excel_path_in_script = basename(file_path)
  )

  # Write script
  dir.create(output_dir, showWarnings = FALSE, recursive = TRUE)
  script_path <- file.path(output_dir, "generated_script.R")
  writeLines(result$script, script_path)
  message(sprintf("Script written: %s", script_path))

  # Export CSVs if standalone mode
  if (mode == "csv") {
    data_dir <- file.path(output_dir, "data")
    sheets <- unique(result$report$Sheet)
    export_sheet_csvs(file_path, sheets, data_dir,
                      formula_data = result$report)
    message(sprintf("CSV data exported: %s/ (%d sheets)", data_dir, length(sheets)))
  }

  # Verify
  verification <- NULL
  if (verify) {
    # Always verify with Excel-mode script
    verify_result <- process_excel_file(
      file_path = file_path,
      wrap_trycatch = wrap_trycatch,
      include_comments = include_comments,
      include_named_tables = include_named_tables,
      data_source = "excel",
      excel_path_in_script = file_path
    )
    verification <- verify_against_excel(file_path, verify_result$report,
                                          verify_result$script)
    s <- verification$summary
    message(sprintf(
      "Verification: %d/%d match, %d mismatches, %d precision diffs, %d harmless NA/error",
      s$matches, s$total, s$value_mismatches, s$fp_precision, s$na_mismatches
    ))
  }

  # Summary
  report <- result$report
  n_ok <- sum(report$Status == "ok")
  n_warn <- sum(report$Status == "warning")
  n_err <- sum(report$Status == "error")
  message(sprintf(
    "Done: %d formulas (%d ok, %d warnings, %d errors) across %d sheets",
    nrow(report), n_ok, n_warn, n_err, length(unique(report$Sheet))
  ))

  invisible(list(
    script = result$script,
    report = result$report,
    named_tables = result$named_tables,
    verification = verification
  ))
}

#' Process an Excel workbook without writing files
#'
#' Extracts formulas, transforms them to R code, detects named tables,
#' and generates the script text. Does not write any files.
#'
#' @param file_path Path to the .xlsx file
#' @param mode Output mode: "csv" or "excel"
#' @param sheet_names Sheets to process (NULL = all)
#' @param include_named_tables Detect named tables? (default TRUE)
#' @param wrap_trycatch Wrap formulas in tryCatch? (default TRUE)
#' @param include_comments Include original formulas as comments? (default TRUE)
#' @return List with: script (character), report (data.frame), named_tables (data.frame or NULL)
#' @export
#'
#' @examples
#' \dontrun{
#' result <- excel2r::process("workbook.xlsx")
#' cat(result$script)
#' View(result$report)
#' }
process <- function(file_path,
                    mode = c("csv", "excel"),
                    sheet_names = NULL,
                    include_named_tables = TRUE,
                    wrap_trycatch = TRUE,
                    include_comments = TRUE) {
  mode <- match.arg(mode)
  process_excel_file(
    file_path = file_path,
    sheet_names = sheet_names,
    wrap_trycatch = wrap_trycatch,
    include_comments = include_comments,
    include_named_tables = include_named_tables,
    data_source = mode,
    excel_path_in_script = basename(file_path)
  )
}

#' Verify R-computed values against Excel cached results
#'
#' Runs the generated script in an isolated environment and compares
#' every formula cell's computed value against what Excel had cached.
#'
#' @param file_path Path to the original .xlsx file
#' @param result Output from process() or migrate()
#' @return List with: summary (match counts), mismatches (data.frame of differences)
#' @export
#'
#' @examples
#' \dontrun{
#' result <- excel2r::process("workbook.xlsx", mode = "excel")
#' v <- excel2r::verify("workbook.xlsx", result)
#' print(v$summary)
#' View(v$mismatches)
#' }
verify <- function(file_path, result) {
  # If result was generated in CSV mode, regenerate Excel-mode for verification
  if (!grepl("read_xlsx", result$script)) {
    excel_result <- process_excel_file(
      file_path = file_path,
      wrap_trycatch = TRUE,
      include_comments = FALSE,
      data_source = "excel",
      excel_path_in_script = file_path
    )
    script_text <- excel_result$script
    report <- excel_result$report
  } else {
    script_text <- result$script
    report <- result$report
  }

  verify_against_excel(file_path, report, script_text)
}

#' List supported Excel functions
#'
#' Returns the names of all Excel functions that excel2r can translate to R.
#'
#' @return Character vector of function names
#' @export
#'
#' @examples
#' excel2r::supported_functions()
supported_functions <- function() {
  get_supported_functions()
}

#' Launch the Excel2R Shiny app
#'
#' Opens the interactive 5-step workflow in your browser.
#'
#' @param ... Arguments passed to shiny::runApp()
#' @export
#'
#' @examples
#' \dontrun{
#' excel2r::run_app()
#' }
run_app <- function(...) {
  if (!requireNamespace("shiny", quietly = TRUE)) {
    stop("Package 'shiny' is required to run the app. Install it with: install.packages('shiny')")
  }
  app_dir <- system.file("app", package = "excel2r")
  if (app_dir == "") {
    stop("Could not find app directory. Try reinstalling the package.")
  }
  shiny::runApp(app_dir, ...)
}
