# =============================================================================
# Excel2R -- Convert Excel formulas to executable R code
# Shiny Web Application (package mode: inst/app/app.R)
# =============================================================================

library(excel2r)
library(shiny)
library(bslib)
library(DT)

# Increase upload limit to 50MB
options(shiny.maxRequestSize = 50 * 1024^2)

# =============================================================================
# UI
# =============================================================================
ui <- page_navbar(
  title = "Excel2R",
  theme = bs_theme(
    version = 5,
    bootswatch = "flatly",
    primary = "#2C3E50"
  ),

  # --- Tab 1: Upload ---
  nav_panel(
    title = "1. Upload",
    icon = icon("upload"),
    layout_columns(
      col_widths = c(4, 8),
      card(
        card_header("Upload Excel File"),
        fileInput("upload", NULL,
                  accept = c(".xlsx", ".xls"),
                  placeholder = "Choose .xlsx file..."),
        hr(),
        htmlOutput("upload_status")
      ),
      card(
        card_header("Summary"),
        conditionalPanel(
          condition = "output.has_results",
          layout_columns(
            col_widths = c(3, 3, 3, 3),
            value_box(
              title = "Sheets",
              value = textOutput("n_sheets"),
              showcase = icon("table"),
              theme = "primary"
            ),
            value_box(
              title = "Formulas",
              value = textOutput("n_formulas"),
              showcase = icon("calculator"),
              theme = "info"
            ),
            value_box(
              title = "OK",
              value = textOutput("n_ok"),
              showcase = icon("check"),
              theme = "success"
            ),
            value_box(
              title = "Warnings",
              value = textOutput("n_warnings"),
              showcase = icon("exclamation-triangle"),
              theme = "warning"
            )
          ),
          hr(),
          h6("Functions Detected:"),
          htmlOutput("detected_functions"),
          hr(),
          h6("Named Tables:"),
          htmlOutput("detected_tables")
        )
      )
    )
  ),

  # --- Tab 2: Review ---
  nav_panel(
    title = "2. Review",
    icon = icon("search"),
    card(
      card_header(
        class = "d-flex justify-content-between align-items-center",
        "Formula Transformation Results",
        div(
          class = "d-flex gap-2",
          selectInput("filter_sheet", "Sheet:", choices = c("All"), width = "200px"),
          selectInput("filter_status", "Status:", choices = c("All", "ok", "warning", "error"), width = "120px")
        )
      ),
      DTOutput("formula_table")
    )
  ),

  # --- Tab 3: Configure ---
  nav_panel(
    title = "3. Configure",
    icon = icon("cog"),
    layout_columns(
      col_widths = c(6, 6),
      card(
        card_header("Output Options"),
        checkboxInput("opt_trycatch", "Wrap each formula in tryCatch() for error safety", value = TRUE),
        checkboxInput("opt_comments", "Include original Excel formulas as comments", value = TRUE),
        checkboxInput("opt_named_tables", "Generate named table data frames with real column headers", value = TRUE),
        radioButtons("opt_data_source", "Data source in generated script:",
          choices = list(
            "Read from Excel file (requires .xlsx at runtime)" = "excel",
            "Read from CSV files (standalone, no Excel needed)" = "csv"
          ),
          selected = "excel"
        ),
        conditionalPanel(
          condition = "input.opt_data_source == 'excel'",
          textInput("opt_filepath", "Excel file path in generated script:",
                    value = "", placeholder = "e.g., data/my_workbook.xlsx")
        )
      ),
      card(
        card_header("Sheet Selection"),
        htmlOutput("sheet_checkboxes"),
        actionButton("select_all", "Select All", class = "btn-sm btn-outline-primary me-2"),
        actionButton("deselect_all", "Deselect All", class = "btn-sm btn-outline-secondary")
      )
    )
  ),

  # --- Tab 4: Download ---
  nav_panel(
    title = "4. Download",
    icon = icon("download"),
    layout_columns(
      col_widths = c(8, 4),
      card(
        card_header("Generated R Script Preview"),
        verbatimTextOutput("script_preview", placeholder = TRUE)
      ),
      card(
        card_header("Downloads"),
        downloadButton("download_script", "Download .R Script", class = "btn-primary w-100 mb-3"),
        downloadButton("download_report", "Download Report (.csv)", class = "btn-outline-primary w-100 mb-3"),
        hr(),
        htmlOutput("download_stats")
      )
    )
  ),

  # --- Tab 5: Verify ---
  nav_panel(
    title = "5. Verify",
    icon = icon("check-double"),
    card(
      card_header("Compare R Values vs Excel"),
      p("Run the generated script and compare computed values against Excel's cached formula results.",
        "Harmless differences (floating-point precision, Excel errors, text placeholders) are excluded."),
      actionButton("btn_verify", "Run Verification", class = "btn-primary mb-3"),
      htmlOutput("verify_summary"),
      conditionalPanel(
        condition = "output.has_verify_results",
        DTOutput("verify_table")
      )
    )
  )
)

# =============================================================================
# Server
# =============================================================================
server <- function(input, output, session) {

  # Reactive values
  rv <- reactiveValues(
    results = NULL,    # output of process_excel_file()
    file_path = NULL,
    processing = FALSE,
    verify = NULL      # output of verify_against_excel()
  )

  # --- File Upload Handler ---
  observeEvent(input$upload, {
    req(input$upload)

    rv$processing <- TRUE

    file_path <- input$upload$datapath
    rv$file_path <- file_path

    withProgress(message = "Processing Excel file...", value = 0, {
      progress_fn <- function(step, detail) {
        incProgress(1/6, detail = detail)
      }

      tryCatch({
        result <- process_excel_file(
          file_path = file_path,
          sheet_names = NULL,  # auto-detect all
          wrap_trycatch = TRUE,
          include_comments = TRUE,
          excel_path_in_script = input$upload$name,
          progress_callback = progress_fn
        )
        rv$results <- result
      }, error = function(e) {
        showNotification(paste("Error:", e$message), type = "error", duration = 10)
        rv$results <- NULL
      })
    })

    rv$processing <- FALSE

    # Update sheet filter
    if (!is.null(rv$results) && nrow(rv$results$report) > 0) {
      sheets <- unique(rv$results$report$Sheet)
      updateSelectInput(session, "filter_sheet", choices = c("All", sheets))
    }
  })

  # --- Output: has_results flag ---
  output$has_results <- reactive({ !is.null(rv$results) })
  outputOptions(output, "has_results", suspendWhenHidden = FALSE)

  # --- Summary Cards ---
  output$n_sheets <- renderText({
    req(rv$results)
    length(unique(rv$results$report$Sheet))
  })

  output$n_formulas <- renderText({
    req(rv$results)
    nrow(rv$results$report)
  })

  output$n_ok <- renderText({
    req(rv$results)
    sum(rv$results$report$Status == "ok")
  })

  output$n_warnings <- renderText({
    req(rv$results)
    sum(rv$results$report$Status != "ok")
  })

  output$upload_status <- renderUI({
    if (rv$processing) {
      tags$div(class = "text-info", icon("spinner", class = "fa-spin"), " Processing...")
    } else if (!is.null(rv$results)) {
      tags$div(class = "text-success", icon("check-circle"), " File processed successfully!")
    } else {
      tags$div(class = "text-muted", "Upload an .xlsx file to begin")
    }
  })

  output$detected_functions <- renderUI({
    req(rv$results)
    funcs <- detect_used_functions(rv$results$report[, c("Sheet", "Cell", "Row", "Col", "Formula")])
    if (length(funcs) == 0) return(tags$em("No Excel functions detected"))

    tags$div(
      lapply(names(funcs), function(fn) {
        supported <- is_function_supported(fn)
        badge_class <- if (supported) "bg-success" else "bg-danger"
        tags$span(class = paste("badge", badge_class, "me-1 mb-1"),
                  paste0(fn, " (", funcs[fn], ")"))
      })
    )
  })

  output$detected_tables <- renderUI({
    req(rv$results)
    tables <- rv$results$named_tables
    if (is.null(tables) || nrow(tables) == 0) {
      return(tags$em("No named tables detected"))
    }
    tags$div(
      lapply(seq_len(nrow(tables)), function(i) {
        t <- tables[i, ]
        n_rows <- t$data_end_row - t$data_start_row + 1
        tags$span(
          class = "badge bg-info me-1 mb-1",
          sprintf("%s (%s, %d rows)", t$table_name, t$sheet, n_rows)
        )
      })
    )
  })

  # --- Review Table ---
  output$formula_table <- renderDT({
    req(rv$results)
    df <- rv$results$report[, c("Sheet", "Cell", "Formula", "R_Code", "Status", "Sheet_Order", "Cell_Order")]

    if (input$filter_sheet != "All") {
      df <- df[df$Sheet == input$filter_sheet, ]
    }
    if (input$filter_status != "All") {
      df <- df[df$Status == input$filter_status, ]
    }

    datatable(
      df,
      options = list(
        pageLength = 25,
        scrollX = TRUE,
        columnDefs = list(
          list(width = "120px", targets = c(0, 1)),
          list(width = "300px", targets = c(2, 3))
        )
      ),
      rownames = FALSE,
      filter = "top"
    ) |>
      formatStyle("Status",
                   backgroundColor = styleEqual(
                     c("ok", "warning", "error"),
                     c("#d4edda", "#fff3cd", "#f8d7da")
                   ))
  })

  # --- Sheet Checkboxes ---
  output$sheet_checkboxes <- renderUI({
    req(rv$results)
    sheets <- unique(rv$results$report$Sheet)
    checkboxGroupInput("selected_sheets", NULL, choices = sheets, selected = sheets)
  })

  observeEvent(input$select_all, {
    req(rv$results)
    sheets <- unique(rv$results$report$Sheet)
    updateCheckboxGroupInput(session, "selected_sheets", selected = sheets)
  })

  observeEvent(input$deselect_all, {
    updateCheckboxGroupInput(session, "selected_sheets", selected = character(0))
  })

  # --- Script generation helper (shared by reactive + verify) ---
  build_script <- function(data_source = "excel") {
    req(rv$results, rv$file_path)
    selected <- if (!is.null(input$selected_sheets)) {
      input$selected_sheets
    } else {
      unique(rv$results$report$Sheet)
    }

    excel_path <- if (data_source == "excel" && nchar(input$opt_filepath) > 0) {
      input$opt_filepath
    } else {
      input$upload$name
    }

    report <- rv$results$report
    if (!is.null(selected)) {
      report <- report[report$Sheet %in% selected, ]
    }

    all_sheets <- if (!is.null(selected)) selected else readxl::excel_sheets(rv$file_path)
    sheet_dims <- detect_sheet_dimensions(rv$file_path, all_sheets)
    used_functions <- detect_used_functions(report[, c("Sheet", "Cell", "Row", "Col", "Formula")])
    exec_order <- determine_execution_order(
      report[, c("Sheet", "Cell", "Row", "Col", "Formula")]
    )

    named_tables <- NULL
    if (isTRUE(input$opt_named_tables) && !is.null(rv$results$named_tables)) {
      named_tables <- rv$results$named_tables
      named_tables <- named_tables[named_tables$sheet %in% all_sheets, , drop = FALSE]
      if (nrow(named_tables) == 0) named_tables <- NULL
    }

    script <- generate_r_script(
      excel_path = excel_path,
      formula_data = report,
      sheet_names = all_sheets,
      sheet_dims = sheet_dims,
      exec_order = exec_order,
      used_functions = used_functions,
      wrap_trycatch = input$opt_trycatch,
      include_comments = input$opt_comments,
      named_tables = named_tables,
      data_source = data_source
    )

    list(script = script, report = report, warnings = rv$results$warnings,
         data_source = data_source)
  }

  # --- Regenerate script reactively when config changes ---
  generated_script <- reactive({
    build_script(data_source = input$opt_data_source)
  })

  # --- Script Preview ---
  output$script_preview <- renderText({
    req(generated_script())
    script <- generated_script()$script
    # Show first 100 lines
    lines <- strsplit(script, "\n")[[1]]
    if (length(lines) > 100) {
      paste(c(lines[1:100], "", sprintf("# ... (%d more lines)", length(lines) - 100)), collapse = "\n")
    } else {
      script
    }
  })

  output$download_stats <- renderUI({
    req(generated_script())
    result <- generated_script()
    script <- result$script
    n_lines <- length(strsplit(script, "\n")[[1]])
    size_kb <- round(nchar(script) / 1024, 1)

    mode_text <- if (identical(result$data_source, "csv")) {
      "Standalone mode: .zip with R script + CSV data (no Excel needed)"
    } else {
      "Excel mode: R script reads from .xlsx at runtime"
    }

    tags$div(
      tags$p(icon("file-code"), sprintf(" %d lines, %.1f KB", n_lines, size_kb)),
      tags$p(tags$strong(mode_text))
    )
  })

  # --- Downloads ---
  output$download_script <- downloadHandler(
    filename = function() {
      name <- tools::file_path_sans_ext(input$upload$name)
      if (identical(input$opt_data_source, "csv")) {
        paste0(name, "_standalone.zip")
      } else {
        paste0(name, "_generated.R")
      }
    },
    content = function(file) {
      result <- generated_script()
      if (identical(result$data_source, "csv")) {
        tmp_dir <- tempfile()
        project_dir <- file.path(tmp_dir, "excel2r_output")
        data_dir <- file.path(project_dir, "data")
        dir.create(data_dir, showWarnings = FALSE, recursive = TRUE)
        on.exit(unlink(tmp_dir, recursive = TRUE), add = TRUE)

        sheets <- unique(result$report$Sheet)
        export_sheet_csvs(rv$file_path, sheets, data_dir,
                          formula_data = result$report)

        writeLines(result$script, file.path(project_dir, "generated_script.R"))

        writeLines(c(
          "# Excel2R Standalone Output",
          "",
          "## Contents",
          "- generated_script.R - R script that computes all Excel formulas",
          "- data/ - raw cell values from each sheet (tidy format: row, col, value)",
          "",
          "## Usage",
          "1. Edit CSV files in data/ to change input values",
          "2. Run: source('generated_script.R')",
          "3. Results are in R data frames, same structure as original Excel sheets",
          "",
          "## No dependencies required",
          "The script uses only base R. No packages to install."
        ), file.path(project_dir, "README.txt"))

        old_wd <- setwd(tmp_dir)
        on.exit(setwd(old_wd), add = TRUE)
        zip(file, "excel2r_output", flags = "-r")
      } else {
        writeLines(result$script, file)
      }
    }
  )

  output$download_report <- downloadHandler(
    filename = function() {
      name <- tools::file_path_sans_ext(input$upload$name)
      paste0(name, "_report.csv")
    },
    content = function(file) {
      write.csv(generated_script()$report, file, row.names = FALSE)
    }
  )

  # --- Verify Tab ---
  observeEvent(input$btn_verify, {
    req(rv$results, rv$file_path)

    rv$verify <- NULL

    withProgress(message = "Verifying R values vs Excel...", value = 0, {
      incProgress(0.1, detail = "Running generated script...")

      tryCatch({
        # Always use Excel-mode script for verification
        # (CSV files don't exist on disk -- only created at download)
        excel_version <- build_script(data_source = "excel")
        result <- verify_against_excel(
          file_path = rv$file_path,
          report = excel_version$report,
          script_text = excel_version$script
        )
        rv$verify <- result
        incProgress(0.9, detail = "Done!")
      }, error = function(e) {
        showNotification(paste("Verification error:", e$message),
                         type = "error", duration = 10)
      })
    })
  })

  output$has_verify_results <- reactive({ !is.null(rv$verify) })
  outputOptions(output, "has_verify_results", suspendWhenHidden = FALSE)

  output$verify_summary <- renderUI({
    req(rv$verify)
    s <- rv$verify$summary

    if (!is.null(s$error)) {
      return(tags$div(class = "alert alert-danger", s$error))
    }

    match_pct <- if (s$total > 0) round(100 * (s$matches + s$fp_precision + s$na_mismatches) / s$total, 1) else 0
    alert_class <- if (s$value_mismatches == 0) "alert-success" else "alert-warning"

    tags$div(
      class = paste("alert", alert_class),
      tags$strong(sprintf("%.1f%% match", match_pct)),
      sprintf(" (%d / %d formulas)", s$matches, s$total),
      tags$br(),
      if (s$value_mismatches > 0)
        tags$span(class = "text-danger",
                  sprintf("%d value mismatch(es)", s$value_mismatches)),
      if (s$fp_precision > 0)
        tags$span(class = "text-muted ms-2",
                  sprintf("| %d minor precision diffs", s$fp_precision)),
      if (s$na_mismatches > 0)
        tags$span(class = "text-muted ms-2",
                  sprintf("| %d harmless NA/error diffs", s$na_mismatches))
    )
  })

  output$verify_table <- renderDT({
    req(rv$verify)
    df <- rv$verify$mismatches
    if (nrow(df) == 0) return(NULL)
    datatable(
      df,
      options = list(pageLength = 25, scrollX = TRUE),
      rownames = FALSE
    )
  })
}

# =============================================================================
# Run
# =============================================================================
shinyApp(ui = ui, server = server)
