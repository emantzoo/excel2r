# =============================================================================
# transform_references.R -- Transform Excel cell/range references to R syntax
# =============================================================================

#' @keywords internal
#' @noRd
transform_single_cell <- function(ref, current_sheet) {
  if (grepl("!", ref)) {
    parts <- strsplit(ref, "!")[[1]]
    sheet <- sanitize_sheet_name(gsub("'", "", parts[1]))
    cell <- parts[2]
  } else {
    sheet <- current_sheet
    cell <- ref
  }

  col <- gsub("[0-9$]+", "", cell)
  row <- as.numeric(gsub("[^0-9]", "", cell))
  if (is.na(row)) {
    warning(sprintf("Invalid cell reference: %s", ref))
    return(ref)
  }

  sprintf("%s$%s[%d]", sheet, col, row)
}

#' @keywords internal
#' @noRd
transform_range <- function(ref, current_sheet, sheet_dims = NULL) {
  if (grepl("!", ref)) {
    parts <- strsplit(ref, "!")[[1]]
    sheet <- sanitize_sheet_name(gsub("'", "", parts[1]))
    cell <- parts[2]
    # Look up dimensions using original sheet name
    original_sheet <- gsub("'", "", parts[1])
  } else {
    sheet <- current_sheet
    cell <- ref
    original_sheet <- NULL
  }

  range_parts <- strsplit(gsub("\\$", "", cell), ":")[[1]]
  col1 <- gsub("[0-9]+", "", range_parts[1])
  col2 <- gsub("[0-9]+", "", range_parts[2])
  row1_str <- gsub("[^0-9]", "", range_parts[1])
  row2_str <- gsub("[^0-9]", "", range_parts[2])
  row1 <- if (row1_str == "") NA else as.numeric(row1_str)
  row2 <- if (row2_str == "") NA else as.numeric(row2_str)

  # Determine max_row from sheet_dims
  max_row <- 1000  # fallback
  if (!is.null(sheet_dims)) {
    # Try original sheet name first, then sanitized
    if (!is.null(original_sheet) && !is.null(sheet_dims[[original_sheet]])) {
      max_row <- sheet_dims[[original_sheet]]$max_row
    } else {
      # Search by sanitized name
      for (sn in names(sheet_dims)) {
        if (sanitize_sheet_name(sn) == sheet) {
          max_row <- sheet_dims[[sn]]$max_row
          break
        }
      }
    }
  }

  if (is.na(row1) && is.na(row2)) {
    # Whole-column reference (e.g., A:A or A:D)
    if (col1 == col2) {
      return(sprintf("%s$%s[1:%d]", sheet, col1, max_row))
    } else {
      # Multi-column whole-column range
      col1_idx <- col_letter_to_index(col1)
      col2_idx <- col_letter_to_index(col2)
      col_letters <- vapply(col1_idx:col2_idx, index_to_col_letter, character(1))
      cols_str <- paste0('"', col_letters, '"', collapse = ", ")
      return(sprintf("unlist(%s[1:%d, c(%s)])", sheet, max_row, cols_str))
    }
  } else if (col1 == col2 && !is.na(row1) && !is.na(row2)) {
    # Same-column range (e.g., D10:D12) -- most common case
    return(sprintf("%s$%s[%d:%d]", sheet, col1, row1, row2))
  } else if (col1 != col2 && !is.na(row1) && !is.na(row2)) {
    # Multi-column range (e.g., A1:D10)
    col1_idx <- col_letter_to_index(col1)
    col2_idx <- col_letter_to_index(col2)
    col_letters <- vapply(col1_idx:col2_idx, index_to_col_letter, character(1))
    cols_str <- paste0('"', col_letters, '"', collapse = ", ")
    return(sprintf("unlist(%s[%d:%d, c(%s)])", sheet, row1, row2, cols_str))
  } else {
    warning(sprintf("Unrecognized range format: %s", ref))
    return(ref)
  }
}

#' @keywords internal
#' @noRd
transform_cell_references <- function(formula, sheet, sheet_dims = NULL) {
  # Step 1: Extract ranges and replace with placeholders
  extracted <- extract_ranges(formula)
  placeholder_formula <- extracted$placeholder_formula
  placeholder_map <- extracted$placeholder_map

  # Step 2: Transform single cell references (not inside placeholders)
  # Pattern matches: optional 'Sheet Name'! followed by $?COL$?ROW
  single_ref_pattern <- "'[^']+'!\\$?[A-Z]{1,3}\\$?[0-9]+|\\$?[A-Z]{1,3}\\$?[0-9]+"
  single_refs <- unique(unlist(regmatches(
    placeholder_formula,
    gregexpr(single_ref_pattern, placeholder_formula)
  )))

  # Filter out placeholders and pure numbers
  single_refs <- single_refs[!grepl("<RANGE_", single_refs)]
  # Filter out things that look like they could be part of function names
  # by checking if preceded by a letter -- but simpler: just check format
  single_refs <- single_refs[grepl("^'?[^']*'?!?\\$?[A-Z]{1,3}\\$?[0-9]+$", single_refs)]

  # Sort by length descending to avoid partial replacements (e.g., replacing "A1" inside "AA1")
  single_refs <- single_refs[order(nchar(single_refs), decreasing = TRUE)]

  for (ref in single_refs) {
    if (ref == "") next
    transformed_ref <- transform_single_cell(ref, sheet)
    # Use word boundaries to avoid replacing substrings
    # Escape all regex special characters in the ref
    escaped_ref <- gsub("([\\\\\\[\\](){}^$.*+?|])", "\\\\\\1", ref, perl = TRUE)
    placeholder_formula <- gsub(
      paste0("(?<![A-Za-z0-9_$])", escaped_ref, "(?![A-Za-z0-9])"),
      transformed_ref,
      placeholder_formula,
      perl = TRUE
    )
  }

  # Step 3: Replace range placeholders with transformed ranges
  for (placeholder in names(placeholder_map)) {
    range_ref <- placeholder_map[[placeholder]]
    transformed_range <- transform_range(range_ref, sheet, sheet_dims)
    placeholder_formula <- gsub(placeholder, transformed_range, placeholder_formula, fixed = TRUE)
  }

  # Step 4: Transform percentages
  placeholder_formula <- transform_percentages(placeholder_formula)

  placeholder_formula
}

#' @keywords internal
#' @noRd
transform_percentages <- function(formula) {
  matches <- gregexpr("[0-9]+\\.?[0-9]*%", formula)[[1]]
  if (matches[1] == -1) return(formula)

  positions <- as.numeric(matches)
  lengths <- attr(matches, "match.length")

  result <- formula
  for (i in length(positions):1) {
    pos <- positions[i]
    len <- lengths[i]
    if (!is_within_quotes(formula, pos)) {
      percent_str <- substr(formula, pos, pos + len - 1)
      number <- as.numeric(gsub("%", "", percent_str)) / 100
      result <- paste0(
        substr(result, 1, pos - 1),
        number,
        substr(result, pos + len, nchar(result))
      )
    }
  }
  result
}
