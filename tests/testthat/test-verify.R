# =============================================================================
# Tests for R/verify_values.R
# =============================================================================

test_that("verify_against_excel works on demo workbook", {
  demo_file <- file.path(
    normalizePath(file.path(getwd(), "..", ".."), winslash = "/"),
    "inst/demo/sales_report_demo.xlsx"
  )
  skip_if_not(file.exists(demo_file), "Demo workbook not found")

  result <- process_excel_file(
    file_path = demo_file,
    wrap_trycatch = TRUE,
    include_comments = TRUE,
    include_named_tables = FALSE,
    excel_path_in_script = demo_file
  )

  v <- verify_against_excel(
    file_path = demo_file,
    report = result$report,
    script_text = result$script
  )

  expect_true(is.list(v))
  expect_true(is.list(v$summary))
  expect_true(is.data.frame(v$mismatches))
  expect_true(v$summary$total > 0)
  expect_equal(v$summary$value_mismatches, 0)
})

test_that("verify_against_excel returns correct structure", {
  demo_file <- file.path(
    normalizePath(file.path(getwd(), "..", ".."), winslash = "/"),
    "inst/demo/sales_report_demo.xlsx"
  )
  skip_if_not(file.exists(demo_file), "Demo workbook not found")

  result <- process_excel_file(
    file_path = demo_file,
    wrap_trycatch = TRUE,
    include_comments = TRUE,
    include_named_tables = FALSE,
    excel_path_in_script = demo_file
  )

  v <- verify_against_excel(
    file_path = demo_file,
    report = result$report,
    script_text = result$script
  )

  # Summary has expected fields
  expect_true("total" %in% names(v$summary))
  expect_true("matches" %in% names(v$summary))
  expect_true("value_mismatches" %in% names(v$summary))
  expect_true("fp_precision" %in% names(v$summary))
  expect_true("na_mismatches" %in% names(v$summary))

  # Mismatches data frame has expected columns
  expect_true(all(c("Sheet", "Cell", "Formula", "R_Value", "Excel_Value", "Category")
                  %in% colnames(v$mismatches)))

  # Total = matches + mismatches + fp + na
  expect_equal(
    v$summary$total,
    v$summary$matches + v$summary$value_mismatches +
      v$summary$fp_precision + v$summary$na_mismatches
  )
})
