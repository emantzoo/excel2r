# =============================================================================
# Regression tests for patterns found in large statistical workbooks (TSA):
# unquoted sheet names, sheet-level loops without cell cycles, implicit
# intersection in SUMIF criteria, text keys, text passed through references,
# empty arguments, nested SUMIF inside SUM, data-only input sheets.
# =============================================================================

# --- Reference scanner ---

test_that("scan_references handles unquoted, quoted and external sheet prefixes", {
  r <- scan_references("T6_MAKE!AZ8+T6")
  expect_equal(r$text, c("T6_MAKE!AZ8", "T6"))
  expect_equal(r$sheet, c("T6_MAKE", NA))

  r <- scan_references("SUMIF('T6 USE'!B:B,\"A1\",C:C)+LOG10(A1)")
  expect_equal(r$text, c("'T6 USE'!B:B", "C:C", "A1"))  # "A1" literal skipped
  expect_true(all(is.na(r$r1[1:2])))                    # whole columns

  r <- scan_references("'[1]GR-TSA Table 2'!$V$18")
  expect_equal(r$sheet, "[1]GR-TSA Table 2")
  expect_equal(c(r$c1, r$r1), c(22L, 18L))
})

test_that("unquoted cross-sheet single refs translate to valid R", {
  out <- transform_cell_references("T6_MAKE!AZ8+SUM(T1_Final!J36)", "T4_Main")
  expect_equal(out, "T6_MAKE$AZ[8]+SUM(T1_Final$J[36])")
  expect_silent(parse(text = out))
})

test_that("implicit intersection rewrites whole-column criteria to the formula row", {
  expect_equal(rewrite_implicit_intersection("SUMIF(M!$D:$D,$F:$F,M!AZ:AZ)", 11L),
               "SUMIF(M!$D:$D,F11,M!AZ:AZ)")
  expect_equal(rewrite_implicit_intersection("COUNTIF(D:D,'My S'!F:F)+1", 7L),
               "COUNTIF(D:D,'My S'!F7)+1")
  # a literal criterion is left alone
  expect_equal(rewrite_implicit_intersection("SUMIF(A1:A9,\">5\",B1:B9)", 3L),
               "SUMIF(A1:A9,\">5\",B1:B9)")
})

test_that("nested SUMIF inside SUM translates both functions", {
  out <- transform_all_functions("SUM(SUMIF(S$A[1:3], 1, S$B[1:3]), 2)")$formula
  expect_false(grepl("SUM\\(", sub("SUMIF\\(", "", out)))  # outer SUM became sum()
  expect_true(grepl("^sum\\(", out))
})

test_that("empty function arguments become 0", {
  out <- transform_all_functions("SUM(S$A[1],S$A[2],)")$formula
  expect_equal(eval(parse(text = out), list(S = data.frame(A = c(1, 2)))), 3)
})

# --- Execution order ---

test_that("global order handles sheet-level loops that are not cell cycles", {
  # A!B1 <- B!A1 <- A!A1 : sheets A and B read each other, cells do not loop
  data <- data.frame(
    Sheet = c("A", "B", "A"), Cell = c("B1", "A1", "A1"),
    Row = c(1L, 1L, 1L), Col = c(2L, 1L, 1L),
    Formula = c("B!A1*2", "A!A1+1", "5"), stringsAsFactors = FALSE)
  res <- determine_execution_order(data)
  expect_length(res$warnings, 0)
  pos <- match(c(3L, 2L, 1L), res$global_order)
  expect_true(pos[1] < pos[2] && pos[2] < pos[3])
})

# --- End to end on a small workbook ---

make_tsa_like_workbook <- function(path) {
  wb <- openxlsx2::wb_workbook()
  # Data-only input sheet with text codes (one numeric-looking)
  wb$add_worksheet("Codes")
  wb$add_data("Codes", x = data.frame(code = c("1.a", "1.b", "2", "1.a"),
                                      val = c(10, 20, 30, 5)),
              dims = "A1", col_names = FALSE)
  # Result sheet: criteria come from a pure reference to a text cell, the
  # criterion is passed as a whole column (implicit intersection)
  wb$add_worksheet("Out")
  wb$add_formula("Out", x = "Codes!A1", dims = "A1")             # "1.a"
  wb$add_formula("Out", x = "Codes!A2", dims = "A2")             # "1.b"
  wb$add_formula("Out", x = "SUMIF(Codes!$A:$A,$A:$A,Codes!B:B)", dims = "B1")
  wb$add_formula("Out", x = "SUMIF(Codes!$A:$A,$A:$A,Codes!B:B)", dims = "B2")
  wb$add_formula("Out", x = "SUM(B1,B2,)", dims = "B3")
  wb$add_formula("Out", x = "SUM(SUMIF(Codes!A1:A4,\"2\",Codes!B1:B4),1)", dims = "B4")
  wb$save(path)
}

test_that("generated script computes text-keyed SUMIFs across sheets", {
  path <- tempfile(fileext = ".xlsx")
  on.exit(unlink(path), add = TRUE)
  make_tsa_like_workbook(path)

  res <- process_excel_file(path, data_source = "excel", excel_path_in_script = path,
                            include_comments = FALSE)
  expect_true(all(res$report$Status == "ok"))

  env <- new.env()
  capture.output(suppressMessages(source(textConnection(res$script), local = env)))
  expect_equal(env$Out$B[1], 15)   # "1.a": 10 + 5
  expect_equal(env$Out$B[2], 20)   # "1.b"
  expect_equal(env$Out$B[3], 35)   # SUM with an empty argument
  expect_equal(env$Out$B[4], 31)   # nested SUMIF on a numeric-looking key
})

test_that("input_cells and external links are read, not recomputed", {
  path <- tempfile(fileext = ".xlsx")
  on.exit(unlink(path), add = TRUE)
  make_tsa_like_workbook(path)
  res <- process_excel_file(path, data_source = "excel", excel_path_in_script = path,
                            input_cells = "Out!B4")
  expect_false("B4" %in% res$report$Cell[res$report$Sheet == "Out"])
  expect_equal(res$inputs$Reason, "input cell")
})
