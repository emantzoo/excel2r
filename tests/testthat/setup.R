# Setup: load functions for tests
# Works in both package mode (devtools::test) and standalone mode (run_tests.R)

if (!exists("sanitize_sheet_name", mode = "function")) {
  # Standalone mode: source all R modules
  r_dir <- file.path(dirname(dirname(getwd())), "R")
  if (!dir.exists(r_dir)) {
    r_dir <- file.path(getwd(), "..", "..", "R")
  }
  if (!dir.exists(r_dir)) {
    r_dir <- "R"
  }
  for (f in list.files(r_dir, pattern = "\\.R$", full.names = TRUE)) {
    source(f)
  }
}

find_demo_file <- function() {
  # Package mode: use system.file
  f <- system.file("extdata/sales_report_demo.xlsx", package = "excel2r")
  if (f != "") return(f)
  # Standalone mode: resolve relative to working directory
  f <- file.path(
    normalizePath(file.path(getwd(), "..", ".."), winslash = "/"),
    "inst/extdata/sales_report_demo.xlsx"
  )
  if (file.exists(f)) return(f)
  NULL
}

find_named_tables_file <- function() {
  f <- system.file("extdata/test_named_tables.xlsx", package = "excel2r")
  if (f != "") return(f)
  f <- file.path(
    normalizePath(file.path(getwd(), "..", ".."), winslash = "/"),
    "inst/extdata/test_named_tables.xlsx"
  )
  if (file.exists(f)) return(f)
  NULL
}
