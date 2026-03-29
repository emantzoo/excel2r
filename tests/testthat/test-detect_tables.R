# =============================================================================
# Tests for R/detect_tables.R
# =============================================================================

# --- parse_table_ref ---

test_that("parse_table_ref handles standard refs", {
  result <- parse_table_ref("A1:E151")
  expect_equal(result$col_start, "A")
  expect_equal(result$col_end, "E")
  expect_equal(result$row_start, 1)
  expect_equal(result$row_end, 151)
  expect_equal(result$col_start_idx, 1)
  expect_equal(result$col_end_idx, 5)
})

test_that("parse_table_ref strips absolute refs", {
  result <- parse_table_ref("$A$1:$E$151")
  expect_equal(result$col_start, "A")
  expect_equal(result$col_end, "E")
  expect_equal(result$row_start, 1)
  expect_equal(result$row_end, 151)
})

test_that("parse_table_ref handles multi-letter columns", {
  result <- parse_table_ref("AA5:AZ100")
  expect_equal(result$col_start, "AA")
  expect_equal(result$col_end, "AZ")
  expect_equal(result$row_start, 5)
  expect_equal(result$row_end, 100)
  expect_equal(result$col_start_idx, 27)
  expect_equal(result$col_end_idx, 52)
})

# --- detect_named_tables ---

test_that("detect_named_tables finds tables in test workbook", {
  test_file <- find_named_tables_file()
  skip_if(is.null(test_file), "Test named tables file not found")

  result <- detect_named_tables(test_file)
  expect_true(is.data.frame(result))
  expect_true(nrow(result) >= 2)

  expect_true("productlist" %in% result$table_name)
  expect_true("salesdata" %in% result$table_name)

  # Check ProductList details
  pl <- result[result$table_name == "productlist", ]
  expect_equal(pl$sheet, "Products")
  expect_equal(pl$header_row, 1)
  expect_equal(pl$data_start_row, 2)

  # Check column names were extracted
  expect_true("Product_ID" %in% pl$col_names[[1]])
  expect_true("Unit_Price" %in% pl$col_names[[1]])
})

test_that("detect_named_tables returns empty data frame when no tables", {
  # Use the demo file which has no named tables
  candidates <- c(
    "inst/extdata/sales_report_demo.xlsx",
    "../../inst/extdata/sales_report_demo.xlsx",
    file.path(getwd(), "inst/extdata/sales_report_demo.xlsx"),
    file.path(getwd(), "../../inst/extdata/sales_report_demo.xlsx")
  )
  demo_file <- NULL
  for (f in candidates) {
    if (file.exists(f)) { demo_file <- normalizePath(f); break }
  }
  skip_if(is.null(demo_file), "Demo file not found")

  result <- detect_named_tables(demo_file)
  expect_true(is.data.frame(result))
  expect_equal(nrow(result), 0)
})

# --- Integration: generate_r_script with named tables ---

test_that("generated script includes named tables section", {
  test_file <- find_named_tables_file()
  skip_if(is.null(test_file), "Test named tables file not found")

  result <- process_excel_file(
    file_path = test_file,
    include_named_tables = TRUE,
    wrap_trycatch = TRUE,
    include_comments = TRUE,
    excel_path_in_script = "test_named_tables.xlsx"
  )

  expect_true(!is.null(result$named_tables))
  expect_true(nrow(result$named_tables) >= 2)

  # Script should contain named tables section
  expect_true(grepl("Named Tables", result$script))
  expect_true(grepl("productlist", result$script, ignore.case = TRUE))
  expect_true(grepl("salesdata", result$script, ignore.case = TRUE))
  expect_true(grepl("colnames\\(", result$script))
})

test_that("generated script omits named tables when disabled", {
  test_file <- find_named_tables_file()
  skip_if(is.null(test_file), "Test named tables file not found")

  result <- process_excel_file(
    file_path = test_file,
    include_named_tables = FALSE,
    wrap_trycatch = TRUE,
    include_comments = TRUE,
    excel_path_in_script = "test_named_tables.xlsx"
  )

  expect_true(is.null(result$named_tables))
  expect_false(grepl("Named Tables", result$script))
})
