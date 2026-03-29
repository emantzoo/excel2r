# =============================================================================
# Tests for CSV export mode
# =============================================================================

find_demo_file <- function() {
  f <- file.path(
    normalizePath(file.path(getwd(), "..", ".."), winslash = "/"),
    "inst/demo/sales_report_demo.xlsx"
  )
  if (file.exists(f)) f else NULL
}

# --- export_sheet_csvs ---

test_that("export_sheet_csvs creates tidy CSV files", {
  demo_file <- find_demo_file()
  skip_if(is.null(demo_file), "Demo workbook not found")

  tmp <- file.path(tempdir(), "test_csvs")
  on.exit(unlink(tmp, recursive = TRUE), add = TRUE)

  sheets <- readxl::excel_sheets(demo_file)
  paths <- export_sheet_csvs(demo_file, sheets, tmp)

  expect_equal(length(paths), length(sheets))
  for (p in paths) expect_true(file.exists(p))

  # Check tidy format: must have row, col, value columns
  csv <- read.csv(paths[1], stringsAsFactors = FALSE)
  expect_true(all(c("row", "col", "value") %in% colnames(csv)))
  expect_true(nrow(csv) > 0)
})

test_that("tidy CSV excludes formula cells", {
  demo_file <- find_demo_file()
  skip_if(is.null(demo_file), "Demo workbook not found")

  tmp <- file.path(tempdir(), "test_exclude")
  on.exit(unlink(tmp, recursive = TRUE), add = TRUE)

  sheets <- readxl::excel_sheets(demo_file)
  paths <- export_sheet_csvs(demo_file, sheets, tmp)

  # Read formula cells via tidyxl
  all_cells <- tidyxl::xlsx_cells(demo_file)
  formula_cells <- all_cells[!is.na(all_cells$formula), ]

  for (s in sheets) {
    csv <- read.csv(paths[s], stringsAsFactors = FALSE)
    sheet_formulas <- formula_cells[formula_cells$sheet == s, ]
    for (j in seq_len(nrow(sheet_formulas))) {
      fc <- sheet_formulas[j, ]
      col_letter <- index_to_col_letter(fc$col)
      matches <- csv$row == fc$row & csv$col == col_letter
      expect_false(any(matches),
        info = sprintf("Formula cell %s!%s%d should not be in CSV",
                       s, col_letter, fc$row))
    }
  }
})

# --- CSV mode script generation ---

test_that("CSV mode script uses base R only", {
  demo_file <- find_demo_file()
  skip_if(is.null(demo_file), "Demo workbook not found")

  result <- process_excel_file(
    file_path = demo_file,
    data_source = "csv",
    excel_path_in_script = "demo.xlsx"
  )

  expect_false(grepl("openxlsx2", result$script))
  expect_false(grepl("read_xlsx", result$script))
  expect_true(grepl("reconstruct_grid", result$script))
  expect_true(grepl("read.csv", result$script))
  expect_true(grepl("data_dir", result$script))
})

test_that("CSV mode produces same formula R_Code as Excel mode", {
  demo_file <- find_demo_file()
  skip_if(is.null(demo_file), "Demo workbook not found")

  result_excel <- process_excel_file(
    file_path = demo_file, data_source = "excel"
  )
  result_csv <- process_excel_file(
    file_path = demo_file, data_source = "csv"
  )

  expect_equal(result_excel$report$R_Code, result_csv$report$R_Code)
  expect_equal(result_excel$report$Status, result_csv$report$Status)
})

test_that("reconstruct_grid rebuilds grid correctly", {
  tmp_csv <- tempfile(fileext = ".csv")
  on.exit(unlink(tmp_csv), add = TRUE)

  write.csv(data.frame(
    row = c(1, 1, 2, 2),
    col = c("A", "B", "A", "B"),
    value = c("Name", "Price", "Widget", "10.5")
  ), tmp_csv, row.names = FALSE)

  # Define reconstruct_grid locally (same as generated script)
  reconstruct_grid <- function(csv_path, max_row, max_col) {
    raw <- read.csv(csv_path, stringsAsFactors = FALSE,
                    colClasses = "character")
    col_names <- generate_col_names(max_col)
    grid <- data.frame(matrix(NA, nrow = max_row, ncol = max_col))
    colnames(grid) <- col_names
    for (i in seq_len(nrow(raw))) {
      r <- as.integer(raw$row[i])
      c <- raw$col[i]
      if (r <= max_row && c %in% col_names) grid[[c]][r] <- raw$value[i]
    }
    grid
  }

  grid <- reconstruct_grid(tmp_csv, max_row = 3, max_col = 2)

  expect_equal(ncol(grid), 2)
  expect_equal(nrow(grid), 3)
  expect_equal(colnames(grid), c("A", "B"))
  expect_equal(grid$A[1], "Name")
  expect_equal(grid$B[2], "10.5")
  expect_true(is.na(grid$A[3]))
})

test_that("CSV mode end-to-end: export + run script", {
  demo_file <- find_demo_file()
  skip_if(is.null(demo_file), "Demo workbook not found")

  tmp_dir <- file.path(tempdir(), "test_e2e")
  data_dir <- file.path(tmp_dir, "data")
  on.exit(unlink(tmp_dir, recursive = TRUE), add = TRUE)

  result <- process_excel_file(
    file_path = demo_file,
    data_source = "csv",
    excel_path_in_script = "demo.xlsx"
  )

  sheets <- unique(result$report$Sheet)
  export_sheet_csvs(demo_file, sheets, data_dir)

  script_path <- file.path(tmp_dir, "generated_script.R")
  writeLines(result$script, script_path)

  # Run in isolated env
  env <- new.env(parent = globalenv())
  old_wd <- setwd(tmp_dir)
  on.exit(setwd(old_wd), add = TRUE)
  source(script_path, local = env)

  # Check that data frames were created
  for (s in sheets) {
    df_name <- sanitize_sheet_name(s)
    expect_true(exists(df_name, envir = env),
                info = sprintf("Data frame %s should exist", df_name))
  }
})
