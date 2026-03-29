# =============================================================================
# dependency_order.R -- Kahn's topological sort for formula execution ordering
# =============================================================================

#' @keywords internal
#' @noRd
determine_execution_order <- function(formula_data) {
  warnings <- character(0)

  # --- Phase 1: Sheet-level ordering ---

  sheets <- unique(formula_data$Sheet)

  # Extract cross-sheet dependencies from raw formulas
  extract_cross_sheets <- function(formula) {
    matches <- regmatches(formula, gregexpr("'[^']+'!", formula))[[1]]
    if (length(matches) == 0) return(character(0))
    unique(gsub("'|!", "", matches))
  }

  # Build sheet dependency graph
  sheet_deps <- setNames(vector("list", length(sheets)), sheets)
  for (i in seq_len(nrow(formula_data))) {
    current_sheet <- formula_data$Sheet[i]
    deps <- extract_cross_sheets(formula_data$Formula[i])
    deps <- deps[deps %in% sheets & deps != current_sheet]
    sheet_deps[[current_sheet]] <- unique(c(sheet_deps[[current_sheet]], deps))
  }

  # Kahn's algorithm for sheets
  sheet_order <- kahns_sort(sheets, sheet_deps)
  if (is.null(sheet_order)) {
    warnings <- c(warnings, "Cycle detected in sheet dependencies! Using original order.")
    sheet_order <- sheets
  }

  # --- Phase 2: Cell-level ordering within each sheet ---

  cell_orders <- setNames(vector("list", length(sheets)), sheets)

  for (s in sheet_order) {
    sheet_data <- formula_data[formula_data$Sheet == s, ]
    cells <- sheet_data$Cell
    if (length(cells) == 0) next

    # Use environment for O(1) lookups
    cell_set <- new.env(hash = TRUE, parent = emptyenv())
    for (c in cells) cell_set[[c]] <- TRUE

    # Build cell dependency graph
    cell_deps <- setNames(vector("list", length(cells)), cells)

    for (i in seq_len(nrow(sheet_data))) {
      cell <- sheet_data$Cell[i]
      formula <- sheet_data$Formula[i]

      # Remove cross-sheet references (they're handled by sheet ordering)
      clean_formula <- gsub("'[^']+'![A-Z$0-9:]+", "", formula)

      dep_cells <- character(0)

      # Extract range references
      range_matches <- regmatches(
        clean_formula,
        gregexpr("([A-Z]{1,3})\\$?([0-9]+)?:([A-Z]{1,3})\\$?([0-9]+)?", clean_formula)
      )[[1]]

      for (rm in range_matches) {
        parts <- regmatches(rm, regexec("([A-Z]{1,3})\\$?([0-9]+)?:([A-Z]{1,3})\\$?([0-9]+)?", rm))[[1]]
        col1 <- parts[2]
        row1_str <- parts[3]
        col2 <- parts[4]
        row2_str <- parts[5]

        row1 <- if (row1_str == "") NA else as.numeric(row1_str)
        row2 <- if (row2_str == "") NA else as.numeric(row2_str)

        if (is.na(row1) && is.na(row2)) {
          # Whole-column reference: depends on all formula cells in those columns
          col1_idx <- col_letter_to_index(col1)
          col2_idx <- col_letter_to_index(col2)
          for (ci in col1_idx:col2_idx) {
            col_letter <- index_to_col_letter(ci)
            for (c in cells) {
              pc <- parse_cell_address(c)
              if (pc$col == col_letter && !is.null(cell_set[[c]])) {
                dep_cells <- c(dep_cells, c)
              }
            }
          }
        } else if (!is.na(row1) && !is.na(row2)) {
          # Explicit range
          col1_idx <- col_letter_to_index(col1)
          col2_idx <- col_letter_to_index(col2)
          for (ci in col1_idx:col2_idx) {
            col_letter <- index_to_col_letter(ci)
            for (r in row1:row2) {
              ref_cell <- paste0(col_letter, r)
              if (!is.null(cell_set[[ref_cell]])) {
                dep_cells <- c(dep_cells, ref_cell)
              }
            }
          }
        }
      }

      # Extract single cell references
      single_matches <- regmatches(
        clean_formula,
        gregexpr("[A-Z]{1,3}\\$?[0-9]+", clean_formula)
      )[[1]]

      for (sm in single_matches) {
        sm_clean <- gsub("\\$", "", sm)
        if (!is.null(cell_set[[sm_clean]])) {
          dep_cells <- c(dep_cells, sm_clean)
        }
      }

      # Remove self-dependency
      dep_cells <- unique(dep_cells[dep_cells != cell])
      cell_deps[[cell]] <- dep_cells
    }

    # Kahn's algorithm for cells
    cell_order <- kahns_sort(cells, cell_deps)
    if (is.null(cell_order)) {
      warnings <- c(warnings, sprintf("Cycle detected in cell dependencies for sheet '%s'! Using original order.", s))
      cell_order <- cells
    }
    cell_orders[[s]] <- cell_order
  }

  list(sheet_order = sheet_order, cell_orders = cell_orders, warnings = warnings)
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
