# =============================================================================
# parse_formula.R -- Tokenizer and parser for Excel formulas
# Handles balanced parentheses and nested function calls
# =============================================================================

#' @keywords internal
#' @noRd
# Paren depth after each character, ignoring parens inside "strings".
# Vectorised: the character loops this replaces cost milliseconds per
# formula, which added up to minutes on workbooks with 10k+ formulas.
.paren_depth <- function(chars) {
  is_q <- chars == '"'
  in_str <- (cumsum(is_q) %% 2 == 1) & !is_q
  step <- ifelse(in_str | is_q, 0L, ifelse(chars == "(", 1L, ifelse(chars == ")", -1L, 0L)))
  list(depth = cumsum(step), in_str = in_str | is_q)
}

#' @keywords internal
#' @noRd
find_matching_paren <- function(formula, open_pos) {
  chars <- strsplit(formula, "", fixed = TRUE)[[1]]
  n <- length(chars)
  if (open_pos >= n) return(NA_integer_)
  tail_chars <- chars[(open_pos + 1):n]
  pd <- .paren_depth(tail_chars)
  # depth relative to the opening paren: 1 + running depth
  hit <- which(1L + pd$depth == 0L & tail_chars == ")" & !pd$in_str)
  if (length(hit) == 0) return(NA_integer_)
  as.integer(open_pos + hit[1])
}

#' @keywords internal
#' @noRd
split_function_args <- function(content) {
  if (is.na(content) || nchar(content) == 0) return(character(0))
  chars <- strsplit(content, "", fixed = TRUE)[[1]]
  pd <- .paren_depth(chars)
  # depth *before* each char equals depth after it for a comma
  commas <- which(chars == "," & pd$depth == 0L & !pd$in_str)
  if (length(commas) == 0) return(trimws(content))
  starts <- c(1L, commas + 1L)
  ends <- c(commas - 1L, length(chars))
  trimws(substring(content, starts, ends))
}

#' @keywords internal
#' @noRd
find_function_calls <- function(formula) {
  # Pattern: function name immediately followed by (
  # Function names: one or more uppercase letters/digits/dots, starting with letter
  # Negative lookbehind prevents matching mid-word (e.g. "Catch" in "tryCatch")
  matches <- gregexpr("(?<![A-Za-z._])[A-Z][A-Za-z0-9.]*\\(", formula, perl = TRUE)[[1]]

  if (matches[1] == -1) return(list())

  calls <- list()
  match_lengths <- attr(matches, "match.length")

  for (idx in seq_along(matches)) {
    start_pos <- matches[idx]
    match_len <- match_lengths[idx]
    func_name <- substr(formula, start_pos, start_pos + match_len - 2)  # exclude '('
    open_paren <- start_pos + match_len - 1

    # Skip if inside a string
    if (is_within_quotes(formula, start_pos)) next

    # Find matching close paren
    close_paren <- find_matching_paren(formula, open_paren)
    if (is.na(close_paren)) next

    # Extract the arguments string (between parens)
    args_str <- substr(formula, open_paren + 1, close_paren - 1)

    calls[[length(calls) + 1]] <- list(
      name = func_name,
      start = start_pos,
      open_paren = open_paren,
      close_paren = close_paren,
      full_match = substr(formula, start_pos, close_paren),
      args_str = args_str
    )
  }

  # Sort right-to-left for safe replacement
  if (length(calls) > 0) {
    starts <- sapply(calls, function(x) x$start)
    calls <- calls[order(starts, decreasing = TRUE)]
  }

  calls
}

#' @keywords internal
#' @noRd
extract_ranges <- function(formula) {
  colon_positions <- gregexpr(":", formula)[[1]]
  placeholder_map <- list()
  placeholders <- list()
  i <- 1

  if (colon_positions[1] != -1) {
    for (col_pos in colon_positions) {
      if (is_within_quotes(formula, col_pos)) next

      # Find left boundary
      left_pos <- col_pos - 1
      while (left_pos > 0) {
        char <- substr(formula, left_pos, left_pos)
        if (char %in% c(",", "(", ")", ";", "-", "+", " ")) break
        # Allow sheet!ref -- include the sheet prefix
        if (char == "!") {
          if (left_pos > 1 && substr(formula, left_pos - 1, left_pos - 1) == "'") {
            # Part of 'Sheet Name'! -- scan left past closing quote to find opening quote
            left_pos <- left_pos - 2  # skip past the closing quote
            while (left_pos > 1 && substr(formula, left_pos, left_pos) != "'") {
              left_pos <- left_pos - 1
            }
            # left_pos is now at the opening quote -- stop scanning
            left_pos <- left_pos - 1
            break
          }
          # Unquoted sheet ref like Sheet1!A1:B10 -- continue scanning left
          left_pos <- left_pos - 1
          while (left_pos > 0) {
            ch2 <- substr(formula, left_pos, left_pos)
            if (ch2 %in% c(",", "(", ")", ";", "-", "+", " ")) break
            left_pos <- left_pos - 1
          }
          next
        }
        left_pos <- left_pos - 1
      }
      left_boundary <- if (left_pos == 0) 1 else left_pos + 1

      # Find right boundary
      right_pos <- col_pos + 1
      while (right_pos <= nchar(formula)) {
        char <- substr(formula, right_pos, right_pos)
        if (char %in% c(",", "(", ")", ";", "-", "+", " ")) break
        right_pos <- right_pos + 1
      }
      right_boundary <- if (right_pos > nchar(formula)) nchar(formula) else right_pos - 1

      range_ref <- substr(formula, left_boundary, right_boundary)

      # Validate: looks like an Excel range (with optional sheet prefix)
      range_pattern <- "^'?[^']*'?!?\\$?[A-Z]{1,3}\\$?[0-9]*:\\$?[A-Z]{1,3}\\$?[0-9]*$"
      if (grepl(range_pattern, range_ref)) {
        placeholder <- sprintf("<RANGE_%d>", i)
        placeholder_map[[placeholder]] <- range_ref
        placeholders[[length(placeholders) + 1]] <- list(
          start = left_boundary, end = right_boundary, ph = placeholder
        )
        i <- i + 1
      }
    }
  }

  # Sort by start position descending (right-to-left replacement)
  if (length(placeholders) > 0) {
    placeholders <- placeholders[order(sapply(placeholders, function(x) x$start),
                                       decreasing = TRUE)]
  }

  # Replace from right to left
  placeholder_formula <- formula
  for (pl in placeholders) {
    placeholder_formula <- paste0(
      substr(placeholder_formula, 1, pl$start - 1),
      pl$ph,
      substr(placeholder_formula, pl$end + 1, nchar(placeholder_formula))
    )
  }

  list(placeholder_formula = placeholder_formula, placeholder_map = placeholder_map)
}

# Regex for one cell or range reference with an optional sheet prefix.
# Sheet prefix: 'Quoted Name'! (with '' escapes) or UnquotedName!
# Reference:    $A$1, A1:B10, A:C (whole columns)
# The lookbehind stops matches inside names (T6_MAKE, LOG10, Sheet1!A1 -> A1);
# the lookahead stops matches on function names (LOG10() and sheet names (T6!).
.ref_regex <- paste0(
  "(?<![A-Za-z0-9_.$'!\\]])",
  "(?:'((?:[^']|'')+)'!|([A-Za-z_][A-Za-z0-9_.]*)!)?",
  "(\\$?[A-Z]{1,3}\\$?[0-9]+(?::\\$?[A-Z]{1,3}\\$?[0-9]+)?|\\$?[A-Z]{1,3}:\\$?[A-Z]{1,3})",
  "(?![A-Za-z0-9_(!.])"
)

#' Find every cell/range reference in a formula
#'
#' Returns one row per reference that is not inside a string literal, with
#' its position, the sheet name (NA when the reference is local) and the
#' column/row bounds (rows are NA for whole-column references).
#' @keywords internal
#' @noRd
scan_references <- function(formula) {
  empty <- data.frame(start = integer(0), end = integer(0), text = character(0),
                      sheet = character(0), c1 = integer(0), r1 = integer(0),
                      c2 = integer(0), r2 = integer(0), is_range = logical(0),
                      stringsAsFactors = FALSE)
  if (is.na(formula) || !nzchar(formula)) return(empty)

  m <- gregexpr(.ref_regex, formula, perl = TRUE)[[1]]
  if (m[1] == -1) return(empty)

  starts <- as.integer(m)
  lens <- attr(m, "match.length")
  cap_s <- attr(m, "capture.start")
  cap_l <- attr(m, "capture.length")
  sub_cap <- function(k) {
    out <- substring(formula, cap_s[, k], cap_s[, k] + cap_l[, k] - 1L)
    out[cap_l[, k] <= 0] <- NA_character_
    out
  }
  quoted <- sub_cap(1)
  unquoted <- sub_cap(2)
  ref <- gsub("$", "", sub_cap(3), fixed = TRUE)

  # Drop matches inside "string literals"
  q <- gregexpr('"', formula, fixed = TRUE)[[1]]
  keep <- rep(TRUE, length(starts))
  if (q[1] != -1 && length(q) >= 2) {
    nq <- length(q) - length(q) %% 2
    for (k in seq(1, nq, by = 2)) {
      keep <- keep & !(starts > q[k] & starts < q[k + 1])
    }
  }

  sheet <- ifelse(!is.na(quoted), gsub("''", "'", quoted), unquoted)
  is_range <- grepl(":", ref, fixed = TRUE)
  first <- sub(":.*$", "", ref)
  second <- sub("^.*:", "", ref)
  row_of <- function(x) suppressWarnings(as.integer(gsub("[^0-9]", "", x)))

  # Built directly as a data.frame (data.frame() itself is slow and this
  # runs several times per formula)
  out <- list(
    start = starts[keep], end = (starts + lens - 1L)[keep],
    text = substring(formula, starts, starts + lens - 1L)[keep],
    sheet = sheet[keep],
    c1 = .col_index(gsub("[0-9]", "", first))[keep], r1 = row_of(first)[keep],
    c2 = .col_index(gsub("[0-9]", "", second))[keep], r2 = row_of(second)[keep],
    is_range = is_range[keep]
  )
  structure(out, class = "data.frame", row.names = seq_len(sum(keep)))
}

#' Apply Excel implicit intersection to *IF criterion arguments
#'
#' In a normal (non-array) formula, a single-column range given where one
#' value is expected -- the criterion of SUMIF/SUMIFS/COUNTIF(S)/AVERAGEIF(S),
#' e.g. SUMIF(D:D, $F:$F, AZ:AZ) -- means the cell of that column in the
#' formula's own row ($F11 for a formula in row 11). Rewriting the formula
#' text makes both the dependency graph and the translation see one cell;
#' left as a whole column it creates false dependency cycles and a vector
#' criterion.
#' @keywords internal
#' @noRd
rewrite_implicit_intersection <- function(formula, row) {
  if (is.na(row) || is.na(formula) ||
      !grepl("(SUMIFS?|COUNTIFS?|AVERAGEIFS?)\\(", formula, ignore.case = TRUE)) {
    return(formula)
  }
  crit_positions <- function(fname, n) {
    switch(toupper(fname),
      SUMIF = , COUNTIF = , AVERAGEIF = 2L,
      SUMIFS = , AVERAGEIFS = seq(3L, n, by = 2L),
      COUNTIFS = seq(2L, n, by = 2L),
      integer(0))
  }
  for (guard in seq_len(200)) {
    changed <- FALSE
    for (call in find_function_calls(formula)) {
      crits <- crit_positions(call$name, 99L)
      if (!length(crits)) next
      content <- call$args_str
      chars <- strsplit(content, "", fixed = TRUE)[[1]]
      if (!length(chars)) next
      pd <- .paren_depth(chars)
      commas <- which(chars == "," & pd$depth == 0L & !pd$in_str)
      a_start <- c(1L, commas + 1L)
      a_end <- c(commas - 1L, length(chars))
      for (j in crits[crits <= length(a_start)]) {
        raw <- substr(content, a_start[j], a_end[j])
        arg <- trimws(raw)
        refs <- scan_references(arg)
        if (nrow(refs) != 1 || refs$text != arg || !refs$is_range || refs$c1 != refs$c2) next
        if (!is.na(refs$r1) && (row < min(refs$r1, refs$r2) || row > max(refs$r1, refs$r2))) next
        prefix <- if (grepl("!", arg, fixed = TRUE)) sub("^(.*!).*$", "\\1", arg) else ""
        new_arg <- paste0(prefix, index_to_col_letter(refs$c1), row)
        # content[k] is formula[open_paren + k]
        pos <- call$open_paren + a_start[j] + regexpr(arg, raw, fixed = TRUE) - 1L
        formula <- paste0(substr(formula, 1, pos - 1L), new_arg,
                          substr(formula, pos + nchar(arg), nchar(formula)))
        changed <- TRUE
        break
      }
      if (changed) break
    }
    if (!changed) break
  }
  formula
}

#' Formulas as evaluated: implicit intersection applied (see above)
#' @keywords internal
#' @noRd
formulas_for_evaluation <- function(formula_data) {
  rows <- if (!is.null(formula_data$Row)) formula_data$Row else
    as.integer(gsub("[^0-9]", "", formula_data$Cell))
  vapply(seq_len(nrow(formula_data)), function(i) {
    rewrite_implicit_intersection(formula_data$Formula[i], rows[i])
  }, character(1))
}

# Column letters -> index, vectorised ("A" -> 1, "AA" -> 27)
.col_index <- function(letters_vec) {
  vapply(letters_vec, function(s) {
    v <- utf8ToInt(s) - 64L
    as.integer(sum(v * 26L^(rev(seq_along(v)) - 1L)))
  }, integer(1), USE.NAMES = FALSE)
}
