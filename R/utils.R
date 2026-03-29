# =============================================================================
# utils.R — Shared utility functions for Excel-to-R conversion
# =============================================================================

#' Sanitize a sheet name into a valid R identifier
#' "Non-Residents Tour Expenditure" -> "NonResidents_Tour_Expenditure"
sanitize_sheet_name <- function(sheet) {
  sheet <- gsub(" ", "_", sheet)
  sheet <- gsub("[^[:alnum:]_]", "", sheet)
  sheet
}

#' Convert column letter(s) to numeric index: A=1, B=2, ..., Z=26, AA=27, ...
col_letter_to_index <- function(col) {
  col <- toupper(col)
  chars <- strsplit(col, "")[[1]]
  idx <- 0
  for (ch in chars) {
    idx <- idx * 26 + (match(ch, LETTERS))
  }
  idx
}

#' Convert numeric index to column letter(s): 1=A, 26=Z, 27=AA, ...
index_to_col_letter <- function(idx) {
  result <- ""
  while (idx > 0) {
    remainder <- (idx - 1) %% 26
    result <- paste0(LETTERS[remainder + 1], result)
    idx <- (idx - 1) %/% 26
  }
  result
}

#' Generate column names for n columns: A, B, ..., Z, AA, AB, ...
generate_col_names <- function(num_cols) {
  if (num_cols < 1) stop("Number of columns must be at least 1")
  vapply(1:num_cols, index_to_col_letter, character(1))
}

#' Check if a character position in a formula is inside a quoted string
is_within_quotes <- function(formula, pos) {
  quote_positions <- gregexpr('"', formula)[[1]]
  if (quote_positions[1] == -1) return(FALSE)
  # Only consider paired quotes; drop trailing unpaired quote
  n_quotes <- length(quote_positions)
  if (n_quotes %% 2 != 0) n_quotes <- n_quotes - 1
  if (n_quotes < 2) return(FALSE)
  for (i in seq(1, n_quotes, by = 2)) {
    start_quote <- quote_positions[i]
    end_quote <- quote_positions[i + 1]
    if (pos > start_quote && pos < end_quote) return(TRUE)
  }
  FALSE
}

#' Parse a cell address into column and row components
#' "D10" -> list(col="D", row=10), "$AB$5" -> list(col="AB", row=5)
parse_cell_address <- function(cell) {
  col <- gsub("[0-9$]+", "", cell)
  row_str <- gsub("[^0-9]", "", cell)
  row <- if (row_str == "") NA_real_ else as.numeric(row_str)
  list(col = col, row = row)
}

#' Expand a range like "B3:E10" into all individual cell addresses
expand_range_to_cells <- function(range_str) {
  # Remove sheet prefix if present
  cell_part <- range_str
  if (grepl("!", range_str)) {
    cell_part <- strsplit(range_str, "!")[[1]][2]
  }

  parts <- strsplit(gsub("\\$", "", cell_part), ":")[[1]]
  if (length(parts) != 2) return(character(0))

  p1 <- parse_cell_address(parts[1])
  p2 <- parse_cell_address(parts[2])

  if (is.na(p1$row) || is.na(p2$row)) return(character(0))

  col1_idx <- col_letter_to_index(p1$col)
  col2_idx <- col_letter_to_index(p2$col)

  cells <- character(0)
  for (ci in col1_idx:col2_idx) {
    col_letter <- index_to_col_letter(ci)
    for (ri in p1$row:p2$row) {
      cells <- c(cells, paste0(col_letter, ri))
    }
  }
  cells
}
