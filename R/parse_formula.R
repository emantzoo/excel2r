# =============================================================================
# parse_formula.R — Tokenizer and parser for Excel formulas
# Handles balanced parentheses and nested function calls
# =============================================================================

#' Find the matching closing parenthesis for an opening parenthesis
#' Respects string literals (double-quoted)
#' @param formula The formula string
#' @param open_pos Position of the opening '('
#' @return Position of the matching ')'
find_matching_paren <- function(formula, open_pos) {
  n <- nchar(formula)
  depth <- 1
  i <- open_pos + 1
  in_string <- FALSE

  while (i <= n && depth > 0) {
    ch <- substr(formula, i, i)
    if (ch == '"') {
      in_string <- !in_string
    } else if (!in_string) {
      if (ch == '(') depth <- depth + 1
      else if (ch == ')') depth <- depth - 1
    }
    if (depth > 0) i <- i + 1
  }

  if (depth != 0) return(NA_integer_)
  i
}

#' Split function arguments at top-level commas (depth 0)
#' Respects nested parentheses and string literals
#' @param content String inside the outermost parentheses (without the parens themselves)
#' @return Character vector of argument strings
split_function_args <- function(content) {
  args <- character(0)
  depth <- 0
  in_string <- FALSE
  current <- ""
  n <- nchar(content)

  if (n == 0) return(character(0))

  for (i in 1:n) {
    ch <- substr(content, i, i)
    if (ch == '"') {
      in_string <- !in_string
      current <- paste0(current, ch)
    } else if (in_string) {
      current <- paste0(current, ch)
    } else if (ch == '(') {
      depth <- depth + 1
      current <- paste0(current, ch)
    } else if (ch == ')') {
      depth <- depth - 1
      current <- paste0(current, ch)
    } else if (ch == ',' && depth == 0) {
      args <- c(args, trimws(current))
      current <- ""
    } else {
      current <- paste0(current, ch)
    }
  }
  # Last argument
  if (nchar(trimws(current)) > 0 || length(args) > 0) {
    args <- c(args, trimws(current))
  }
  args
}

#' Find all top-level function calls in a formula
#' Returns a list of list(name, start, open_paren, close_paren, args_str)
#' Processes from right-to-left for safe replacement
#' @param formula The formula string
#' @return List of function call descriptors, ordered right-to-left
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

#' Identify and extract range references from a formula, replacing with placeholders
#' Handles cross-sheet ranges like 'Sheet Name'!A1:B10
#' @param formula The formula string
#' @return list(placeholder_formula, placeholder_map)
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
        # Allow sheet!ref — include the sheet prefix
        if (char == "!") {
          if (left_pos > 1 && substr(formula, left_pos - 1, left_pos - 1) == "'") {
            # Part of 'Sheet Name'! — scan further left to include the quoted sheet
            left_pos <- left_pos - 1
            # Now find the opening quote
            while (left_pos > 1 && substr(formula, left_pos, left_pos) != "'") {
              left_pos <- left_pos - 1
            }
            # left_pos is now at the opening quote
            next
          }
          # Unquoted sheet ref like Sheet1!A1:B10 — continue scanning left
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
