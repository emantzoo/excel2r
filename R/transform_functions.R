# =============================================================================
# transform_functions.R -- Excel function -> R code registry
# =============================================================================

# Transform-time state. `shadow` maps a sanitized sheet name to the columns
# that have a text copy (`<Sheet>__txt`), i.e. columns used as *IF criteria.
# Set by process_excel_file() for the duration of a transform.
.excel2r_state <- new.env(parent = emptyenv())

#' Point a translated range/cell at the text copy of its sheet when every
#' column it touches has one. Criteria must see the original text (codes such
#' as "1.a"), which the numeric sheet frames turn into 0.
#' @keywords internal
#' @noRd
to_text_shadow <- function(code) {
  shadow <- .excel2r_state$shadow
  if (is.null(shadow)) return(code)
  code <- trimws(code)
  # Sheet$COL[...]
  m <- regmatches(code, regexec("^([A-Za-z0-9_]+)\\$([A-Z]+)\\[([^]]+)\\]$", code))[[1]]
  if (length(m) == 4 && m[3] %in% shadow[[m[2]]]) {
    return(sprintf("%s__txt$%s[%s]", m[2], m[3], m[4]))
  }
  # unlist(Sheet[rows, c("B", "C")])
  m <- regmatches(code, regexec("^unlist\\(([A-Za-z0-9_]+)\\[(.+), c\\((.+)\\)\\]\\)$", code))[[1]]
  if (length(m) == 4) {
    cols <- gsub('"', "", trimws(strsplit(m[4], ",")[[1]]))
    if (length(cols) && all(cols %in% shadow[[m[2]]])) {
      return(sprintf("unlist(%s__txt[%s, c(%s)])", m[2], m[3], m[4]))
    }
  }
  code
}

# Rewrite the criteria-range and criterion arguments of a *IF/*IFS call
# (implicit intersection of criteria is applied earlier, on the formula
# text: see rewrite_implicit_intersection)
shadow_criteria_args <- function(args, ranges, crits) {
  for (j in c(ranges, crits)) {
    if (j <= length(args)) args[j] <- to_text_shadow(args[j])
  }
  args
}

#' @keywords internal
excel_function_registry <- list(

  # --- Aggregation ---
  "SUM" = function(args) {
    if (length(args) == 1) {
      sprintf("sum(%s, na.rm=TRUE)", args[1])
    } else {
      sprintf("sum(c(%s), na.rm=TRUE)", paste(args, collapse = ", "))
    }
  },

  "AVERAGE" = function(args) {
    if (length(args) == 1) {
      sprintf("mean(%s, na.rm=TRUE)", args[1])
    } else {
      sprintf("mean(c(%s), na.rm=TRUE)", paste(args, collapse = ", "))
    }
  },

  "MIN" = function(args) {
    if (length(args) == 1) {
      sprintf("min(%s, na.rm=TRUE)", args[1])
    } else {
      sprintf("min(c(%s), na.rm=TRUE)", paste(args, collapse = ", "))
    }
  },

  "MAX" = function(args) {
    if (length(args) == 1) {
      sprintf("max(%s, na.rm=TRUE)", args[1])
    } else {
      sprintf("max(c(%s), na.rm=TRUE)", paste(args, collapse = ", "))
    }
  },

  "MEDIAN" = function(args) {
    if (length(args) == 1) {
      sprintf("median(%s, na.rm=TRUE)", args[1])
    } else {
      sprintf("median(c(%s), na.rm=TRUE)", paste(args, collapse = ", "))
    }
  },

  "PRODUCT" = function(args) {
    if (length(args) == 1) {
      sprintf("prod(%s, na.rm=TRUE)", args[1])
    } else {
      sprintf("prod(c(%s), na.rm=TRUE)", paste(args, collapse = ", "))
    }
  },

  # --- Counting ---
  "COUNT" = function(args) {
    # COUNT counts numeric (non-NA) values; use per-element check, not is.numeric() on whole vector
    sprintf("sum(!is.na(%s) & suppressWarnings(!is.na(as.numeric(%s))))", args[1], args[1])
  },

  "COUNTA" = function(args) {
    sprintf("sum(!is.na(%s))", args[1])
  },

  "COUNTBLANK" = function(args) {
    sprintf("sum(is.na(%s))", args[1])
  },

  # --- Conditional ---
  "IF" = function(args) {
    if (length(args) >= 3) {
      sprintf("ifelse(%s, %s, %s)", args[1], args[2], args[3])
    } else if (length(args) == 2) {
      sprintf("ifelse(%s, %s, FALSE)", args[1], args[2])
    } else {
      sprintf("ifelse(%s, TRUE, FALSE)", args[1])
    }
  },

  "IFS" = function(args) {
    # Pairs of (condition, value)
    if (length(args) >= 2) {
      pairs <- list()
      for (i in seq(1, length(args) - 1, by = 2)) {
        cond <- args[i]
        val <- if (i + 1 <= length(args)) args[i + 1] else "NA"
        pairs[[length(pairs) + 1]] <- sprintf("%s ~ %s", cond, val)
      }
      sprintf("dplyr::case_when(%s)", paste(unlist(pairs), collapse = ", "))
    } else {
      paste0("NA  # IFS with insufficient args: ", paste(args, collapse = ", "))
    }
  },

  "IFERROR" = function(args) {
    fallback <- if (length(args) >= 2) args[2] else "NA"
    # Catch R errors, NA, NaN, and Inf (matches Excel #N/A, #DIV/0!, #VALUE!, etc.)
    sprintf("tryCatch({ .iferr_val <- (%s); if (is.na(.iferr_val) || is.nan(.iferr_val) || is.infinite(.iferr_val)) %s else .iferr_val }, error = function(e) %s)",
            args[1], fallback, fallback)
  },

  "IFNA" = function(args) {
    if (length(args) >= 2) {
      sprintf("ifelse(is.na(%s), %s, %s)", args[1], args[2], args[1])
    } else {
      args[1]
    }
  },

  # --- Conditional Aggregation (using ExcelFunctionsR) ---
  # Criteria ranges and criterion cells read the text copy of their sheet
  # (see to_text_shadow); sum/average ranges stay numeric.
  "SUMIF" = function(args) {
    args <- shadow_criteria_args(args, 1L, 2L)
    sprintf("SUMIF(%s)", paste(args, collapse = ", "))
  },

  "SUMIFS" = function(args) {
    n <- length(args)
    args <- shadow_criteria_args(args, seq(2L, n, by = 2L), seq(3L, n, by = 2L))
    sprintf("SUMIFS(%s)", paste(args, collapse = ", "))
  },

  "COUNTIF" = function(args) {
    args <- shadow_criteria_args(args, 1L, 2L)
    sprintf("COUNTIF(%s)", paste(args, collapse = ", "))
  },

  "COUNTIFS" = function(args) {
    n <- length(args)
    args <- shadow_criteria_args(args, seq(1L, n, by = 2L), seq(2L, n, by = 2L))
    sprintf("COUNTIFS(%s)", paste(args, collapse = ", "))
  },

  "AVERAGEIF" = function(args) {
    args <- shadow_criteria_args(args, 1L, 2L)
    sprintf("AVERAGEIF(%s)", paste(args, collapse = ", "))
  },

  "AVERAGEIFS" = function(args) {
    n <- length(args)
    args <- shadow_criteria_args(args, seq(2L, n, by = 2L), seq(3L, n, by = 2L))
    sprintf("AVERAGEIFS(%s)", paste(args, collapse = ", "))
  },

  # --- Lookup ---
  "VLOOKUP" = function(args) {
    exact <- if (length(args) >= 4) args[4] else "FALSE"
    sprintf("excel_vlookup(%s, %s, %s, %s)", args[1], args[2], args[3], exact)
  },

  "HLOOKUP" = function(args) {
    exact <- if (length(args) >= 4) args[4] else "FALSE"
    sprintf("excel_hlookup(%s, %s, %s, %s)", args[1], args[2], args[3], exact)
  },

  "INDEX" = function(args) {
    if (length(args) == 3) {
      sprintf("excel_index(%s, %s, %s)", args[1], args[2], args[3])
    } else if (length(args) == 2) {
      sprintf("excel_index(%s, %s)", args[1], args[2])
    } else {
      args[1]
    }
  },

  "MATCH" = function(args) {
    match_type <- if (length(args) >= 3) args[3] else "0"
    sprintf("excel_match(%s, %s, %s)", args[1], args[2], match_type)
  },

  "XLOOKUP" = function(args) {
    # XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode])
    if_not_found <- if (length(args) >= 4) args[4] else "NA"
    sprintf("excel_xlookup(%s, %s, %s, %s)", args[1], args[2], args[3], if_not_found)
  },

  # --- Math ---
  "ROUND" = function(args) {
    digits <- if (length(args) >= 2) args[2] else "0"
    sprintf("round(%s, %s)", args[1], digits)
  },

  "ROUNDUP" = function(args) {
    digits <- if (length(args) >= 2) args[2] else "0"
    sprintf("excel_roundup(%s, %s)", args[1], digits)
  },

  "ROUNDDOWN" = function(args) {
    digits <- if (length(args) >= 2) args[2] else "0"
    sprintf("excel_rounddown(%s, %s)", args[1], digits)
  },

  "INT" = function(args) {
    sprintf("as.integer(floor(%s))", args[1])
  },

  "ABS" = function(args) {
    sprintf("abs(%s)", args[1])
  },

  "SQRT" = function(args) {
    sprintf("sqrt(%s)", args[1])
  },

  "POWER" = function(args) {
    sprintf("(%s) ^ (%s)", args[1], args[2])
  },

  "LOG" = function(args) {
    if (length(args) >= 2) {
      sprintf("log(%s, base = %s)", args[1], args[2])
    } else {
      sprintf("log10(%s)", args[1])  # Excel LOG defaults to base 10
    }
  },

  "LN" = function(args) {
    sprintf("log(%s)", args[1])
  },

  "LOG10" = function(args) {
    sprintf("log10(%s)", args[1])
  },

  "MOD" = function(args) {
    sprintf("(%s) %%%% (%s)", args[1], args[2])
  },

  "SIGN" = function(args) {
    sprintf("sign(%s)", args[1])
  },

  "EXP" = function(args) {
    sprintf("exp(%s)", args[1])
  },

  "PI" = function(args) {
    "pi"
  },

  # --- Text ---
  "CONCATENATE" = function(args) {
    sprintf("paste0(%s)", paste(args, collapse = ", "))
  },

  "CONCAT" = function(args) {
    sprintf("paste0(%s)", paste(args, collapse = ", "))
  },

  "LEFT" = function(args) {
    n <- if (length(args) >= 2) args[2] else "1"
    sprintf("substr(%s, 1, %s)", args[1], n)
  },

  "RIGHT" = function(args) {
    n <- if (length(args) >= 2) args[2] else "1"
    sprintf("substr(%s, nchar(%s) - %s + 1, nchar(%s))", args[1], args[1], n, args[1])
  },

  "MID" = function(args) {
    sprintf("substr(%s, %s, %s + %s - 1)", args[1], args[2], args[2], args[3])
  },

  "LEN" = function(args) {
    sprintf("nchar(%s)", args[1])
  },

  "UPPER" = function(args) {
    sprintf("toupper(%s)", args[1])
  },

  "LOWER" = function(args) {
    sprintf("tolower(%s)", args[1])
  },

  "TRIM" = function(args) {
    sprintf("trimws(%s)", args[1])
  },

  "SUBSTITUTE" = function(args) {
    if (length(args) >= 3) {
      sprintf("gsub(%s, %s, %s, fixed = TRUE)", args[2], args[3], args[1])
    } else {
      paste0("NA  # SUBSTITUTE needs 3+ args")
    }
  },

  "TEXT" = function(args) {
    if (length(args) >= 2) {
      # Map common Excel format strings to R format specs
      fmt <- gsub('"', '', args[2])
      r_fmt <- switch(fmt,
        "0" = "%.0f",
        "0.0" = "%.1f",
        "0.00" = "%.2f",
        "0.000" = "%.3f",
        "0%" = , "0.0%" = , "0.00%" = fmt,  # percent handled below
        "#,##0" = "%.0f",
        "#,##0.00" = "%.2f",
        NULL  # unrecognized
      )
      if (grepl("%", fmt)) {
        sprintf("paste0(round(%s * 100, 1), '%%')", args[1])
      } else if (!is.null(r_fmt)) {
        sprintf("sprintf('%s', %s)", r_fmt, args[1])
      } else {
        sprintf("format(%s)", args[1])
      }
    } else {
      sprintf("format(%s)", args[1])
    }
  },

  "VALUE" = function(args) {
    sprintf("as.numeric(%s)", args[1])
  },

  # --- Logical ---
  "AND" = function(args) {
    sprintf("(%s)", paste(args, collapse = " & "))
  },

  "OR" = function(args) {
    sprintf("(%s)", paste(args, collapse = " | "))
  },

  "NOT" = function(args) {
    sprintf("!(%s)", args[1])
  },

  "TRUE" = function(args) "TRUE",
  "FALSE" = function(args) "FALSE",

  # --- Info ---
  "ISNA" = function(args) {
    sprintf("is.na(%s)", args[1])
  },

  "ISBLANK" = function(args) {
    sprintf("is.na(%s)", args[1])
  },

  "ISNUMBER" = function(args) {
    sprintf("is.numeric(%s)", args[1])
  },

  "ISTEXT" = function(args) {
    sprintf("is.character(%s)", args[1])
  },

  "ISERROR" = function(args) {
    sprintf("inherits(tryCatch(%s, error = identity), 'error')", args[1])
  },

  # --- Row/Column ---
  "ROW" = function(args) {
    if (length(args) == 0 || args[1] == "") {
      "NA  # ROW() requires context"
    } else {
      # Args are already transformed: Sheet$Col[Row] -> extract row number from [Row]
      ref <- args[1]
      m <- regmatches(ref, regexpr("\\[([0-9]+)\\]", ref))
      if (length(m) > 0 && nchar(m) > 0) {
        gsub("[\\[\\]]", "", m)
      } else {
        # Fallback: try to extract from original-style ref
        sprintf("as.numeric(gsub('[^0-9]', '', '%s'))", ref)
      }
    }
  },

  "COLUMN" = function(args) {
    if (length(args) == 0 || args[1] == "") {
      "NA  # COLUMN() requires context"
    } else {
      # Args are already transformed: Sheet$Col[Row] -> extract col letter from $Col[
      ref <- args[1]
      m <- regmatches(ref, regexpr("\\$([A-Z]+)\\[", ref))
      if (length(m) > 0 && nchar(m) > 0) {
        col_letter <- gsub("[\\$\\[]", "", m)
        sprintf("col_letter_to_index('%s')", col_letter)
      } else {
        # Fallback: try to extract from original-style ref
        sprintf("col_letter_to_index(gsub('[0-9$]', '', '%s'))", ref)
      }
    }
  }
)

#' @keywords internal
#' @noRd
get_supported_functions <- function() {
  names(excel_function_registry)
}

#' @keywords internal
#' @noRd
is_function_supported <- function(func_name) {
  toupper(func_name) %in% names(excel_function_registry)
}

#' @keywords internal
#' @noRd
transform_function_call <- function(func_name, args) {
  upper_name <- toupper(func_name)
  if (upper_name %in% names(excel_function_registry)) {
    handler <- excel_function_registry[[upper_name]]
    code <- tryCatch(
      handler(args),
      error = function(e) {
        sprintf("NA  # Error transforming %s: %s", func_name, e$message)
      }
    )
    list(status = "ok", code = code)
  } else {
    list(
      status = "unsupported",
      code = sprintf("NA  # Unsupported: %s(%s)", func_name, paste(args, collapse = ", "))
    )
  }
}

#' @keywords internal
#' @noRd
transform_all_functions <- function(formula) {
  # Single recursive pass: translate each outermost call's arguments first,
  # then the call itself. (The previous loop re-picked the smallest call each
  # iteration; translated SUMIF/SUMIFS/COUNTIF keep their upper-case names,
  # so the same call was re-translated up to 50 times and any function
  # around it -- e.g. the SUM in SUM(SUMIF(...)) -- was never translated.)
  warnings <- character(0)

  transform_expr <- function(s, depth = 0L) {
    if (depth > 50L) return(s)
    calls <- find_function_calls(s)
    if (length(calls) == 0) return(s)
    starts <- vapply(calls, function(x) x$start, numeric(1))
    ends <- vapply(calls, function(x) x$close_paren, numeric(1))
    outer <- vapply(seq_along(calls), function(k) {
      !any(starts < starts[k] & ends > ends[k])
    }, logical(1))
    calls <- calls[outer]
    calls <- calls[order(vapply(calls, function(x) x$start, numeric(1)), decreasing = TRUE)]
    for (call in calls) {
      args <- split_function_args(call$args_str)
      if (length(args)) args <- vapply(args, transform_expr, character(1), depth = depth + 1L,
                                       USE.NAMES = FALSE)
      # An empty argument -- SUM(A1,B1,) or IF(x,,y) -- is 0 in Excel
      args[!nzchar(trimws(args))] <- "0"
      transformed <- transform_function_call(call$name, args)
      if (transformed$status == "unsupported") {
        warnings <<- c(warnings, sprintf("Unsupported function: %s", call$name))
      }
      s <- paste0(substr(s, 1, call$start - 1), transformed$code,
                  substr(s, call$close_paren + 1, nchar(s)))
    }
    s
  }

  list(formula = transform_expr(formula), warnings = unique(warnings))
}
