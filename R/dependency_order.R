# =============================================================================
# dependency_order.R -- Kahn's topological sort for formula execution ordering
# =============================================================================
#
# Ordering is done over ALL formula cells of the workbook at once (one global
# graph), the way Excel's calc chain works. A sheet-first order cannot handle
# workbooks whose sheets reference each other in a loop even though the cells
# do not (Sheet A!X1 -> Sheet B!Y1 -> Sheet A!Z1), and it silently produced a
# wrong order in that case.
#
# Range dependencies are resolved against an index of formula cells per sheet
# (vectorised, cached per distinct range) instead of expanding every range
# cell by cell, so workbooks with tens of thousands of SUMIF formulas over
# long ranges are ordered in seconds.
# =============================================================================

#' @keywords internal
#' @noRd
determine_execution_order <- function(formula_data, all_sheets = NULL) {
  warnings <- character(0)
  n <- nrow(formula_data)
  sheets <- unique(formula_data$Sheet)

  if (n == 0) {
    return(list(sheet_order = character(0), cell_orders = list(),
                global_order = integer(0), warnings = warnings))
  }

  # --- Index formula cells by sheet, column, row ---
  addr <- toupper(gsub("$", "", formula_data$Cell, fixed = TRUE))
  f_col <- if (!is.null(formula_data$Col)) as.integer(formula_data$Col) else
    as.integer(vapply(gsub("[0-9]", "", addr), col_letter_to_index, numeric(1)))
  f_row <- if (!is.null(formula_data$Row)) as.integer(formula_data$Row) else
    as.integer(gsub("[^0-9]", "", addr))

  by_sheet <- split(seq_len(n), factor(formula_data$Sheet, levels = sheets))
  cell_index <- new.env(hash = TRUE, parent = emptyenv(), size = n)
  for (i in seq_len(n)) {
    assign(paste(formula_data$Sheet[i], f_col[i], f_row[i]), i, envir = cell_index)
  }

  # Excel sheet names are case-insensitive
  known <- unique(c(sheets, all_sheets))
  sheet_lookup <- setNames(known, toupper(known))

  range_cache <- new.env(hash = TRUE, parent = emptyenv())

  resolve_ref <- function(s, c1, r1, c2, r2) {
    if (!is.na(r1) && r1 == r2 && c1 == c2) {
      hit <- cell_index[[paste(s, c1, r1)]]
      return(if (is.null(hit)) integer(0) else hit)
    }
    ck <- paste(s, c1, r1, c2, r2)
    hit <- range_cache[[ck]]
    if (!is.null(hit)) return(hit)
    idx <- by_sheet[[s]]
    if (is.null(idx) || length(idx) == 0) {
      hit <- integer(0)
    } else {
      lo_c <- min(c1, c2); hi_c <- max(c1, c2)
      sel <- f_col[idx] >= lo_c & f_col[idx] <= hi_c
      if (!is.na(r1) && !is.na(r2)) {
        sel <- sel & f_row[idx] >= min(r1, r2) & f_row[idx] <= max(r1, r2)
      }
      hit <- idx[sel]
    }
    assign(ck, hit, envir = range_cache)
    hit
  }

  # --- Build the global dependency graph ---
  # Implicit intersection first: SUMIF(D:D, $F:$F, ...) in row 11 depends
  # on F11, not on all of column F (which creates false cycles)
  eval_formulas <- formulas_for_evaluation(formula_data)
  deps <- vector("list", n)
  for (i in seq_len(n)) {
    refs <- scan_references(eval_formulas[i])
    if (nrow(refs) == 0) { deps[[i]] <- integer(0); next }
    target <- ifelse(is.na(refs$sheet), formula_data$Sheet[i],
                     unname(sheet_lookup[toupper(refs$sheet)]))
    d <- integer(0)
    for (k in seq_len(nrow(refs))) {
      if (is.na(target[k])) next  # external workbook or unknown sheet
      d <- c(d, resolve_ref(target[k], refs$c1[k], refs$r1[k], refs$c2[k], refs$r2[k]))
    }
    d <- unique(d)
    deps[[i]] <- d[d != i]
  }

  # --- Kahn's algorithm on integer node ids ---
  indeg <- lengths(deps)
  rev_adj <- split(rep(seq_len(n), indeg), factor(unlist(deps), levels = seq_len(n)))
  queue <- integer(n)
  ready <- which(indeg == 0L)
  q_tail <- length(ready)
  queue[seq_len(q_tail)] <- ready
  q_head <- 1L
  while (q_head <= q_tail) {
    v <- queue[q_head]
    q_head <- q_head + 1L
    w <- rev_adj[[v]]
    if (length(w)) {
      indeg[w] <- indeg[w] - 1L
      newly <- w[indeg[w] == 0L]
      if (length(newly)) {
        queue[(q_tail + 1L):(q_tail + length(newly))] <- newly
        q_tail <- q_tail + length(newly)
      }
    }
  }
  global_order <- queue[seq_len(q_tail)]

  if (q_tail < n) {
    stuck <- setdiff(seq_len(n), global_order)
    warnings <- c(warnings, sprintf(
      "Cycle detected in cell dependencies (%d cells, e.g. %s)! Those cells use original order.",
      length(stuck),
      paste(head(paste0(formula_data$Sheet[stuck], "!", formula_data$Cell[stuck]), 5), collapse = ", ")))
    global_order <- c(global_order, stuck)
  }

  # --- Sheet order (for loading data and reporting) ---
  # Topological over sheets when possible; with a sheet-level loop (legal,
  # see header) fall back to the order in which sheets first appear in the
  # global cell order.
  sheet_deps <- setNames(vector("list", length(sheets)), sheets)
  for (i in seq_len(n)) {
    if (length(deps[[i]]) == 0) next
    ds <- unique(formula_data$Sheet[deps[[i]]])
    s <- formula_data$Sheet[i]
    sheet_deps[[s]] <- unique(c(sheet_deps[[s]], ds[ds != s]))
  }
  sheet_order <- kahns_sort(sheets, sheet_deps)
  if (is.null(sheet_order)) {
    sheet_order <- unique(formula_data$Sheet[global_order])
  }

  # Per-sheet cell orders (subsequences of the global order), kept for
  # callers that display order within a sheet
  ordered_sheet <- formula_data$Sheet[global_order]
  cell_orders <- lapply(setNames(sheets, sheets), function(s) {
    formula_data$Cell[global_order[ordered_sheet == s]]
  })

  list(sheet_order = sheet_order, cell_orders = cell_orders,
       global_order = global_order, warnings = warnings)
}

#' @keywords internal
#' @noRd
kahns_sort <- function(nodes, deps) {
  # Build reverse graph
  reverse_graph <- setNames(vector("list", length(nodes)), nodes)
  indegree <- setNames(rep(0L, length(nodes)), nodes)

  for (node in nodes) {
    node_deps <- deps[[node]]
    indegree[node] <- length(node_deps)
    for (dep in node_deps) {
      if (dep %in% nodes) {
        reverse_graph[[dep]] <- c(reverse_graph[[dep]], node)
      } else {
        # Dependency on non-formula cell (raw data) -- not a real dependency
        indegree[node] <- indegree[node] - 1L
      }
    }
  }

  # Process nodes with no dependencies first
  queue <- nodes[indegree == 0]
  sorted <- character(0)

  while (length(queue) > 0) {
    current <- queue[1]
    queue <- queue[-1]
    sorted <- c(sorted, current)
    for (dependent in reverse_graph[[current]]) {
      indegree[dependent] <- indegree[dependent] - 1L
      if (indegree[dependent] == 0) {
        queue <- c(queue, dependent)
      }
    }
  }

  if (length(sorted) != length(nodes)) {
    return(NULL)  # Cycle detected
  }

  sorted
}
