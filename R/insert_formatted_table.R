#' Insert a formatted table from one Word file into another
#'
#' @description
#' Copies a table from `source_docx_path` and inserts it into `template_path`
#' at the paragraph containing `placeholder_text`. The placeholder paragraph is
#' removed, and the resulting document is written to `output_path`.
#'
#' @param source_docx_path Path to the source `.docx` containing the table.
#' @param template_path Path to the template `.docx` containing placeholder text.
#' @param output_path Path to write the output `.docx`.
#' @param placeholder_text Placeholder text to replace.
#' @param table_index 1-based index of the table in `source_docx_path`.
#'   Default is `1`.
#' @param debug Debug.
#'
#' @return Invisibly returns `output_path`.
#' @export
#'
#' @examples \dontrun{
#' insert_formatted_table(
#'   source_docx_path = "source.docx",
#'   template_path = "template.docx",
#'   output_path = "final.docx",
#'   placeholder_text = "{rpfy}:table1.csv",
#'   table_index = 1
#' )
#' }
insert_formatted_table <- function(
  source_docx_path,
  template_path,
  output_path,
  placeholder_text,
  table_index = 1,
  debug = FALSE
) {
  log4r::debug(.le$logger, "Starting insert_formatted_table function")
  tictoc::tic()

  if (debug) {
    log4r::debug(.le$logger, "Debug mode enabled")
    browser()
  }

  if (!is.character(source_docx_path) || length(source_docx_path) != 1) {
    stop("source_docx_path must be a single file path")
  }
  if (!is.character(template_path) || length(template_path) != 1) {
    stop("template_path must be a single file path")
  }
  if (!is.character(output_path) || length(output_path) != 1) {
    stop("output_path must be a single file path")
  }
  if (!is.character(placeholder_text) || length(placeholder_text) != 1) {
    stop("placeholder_text must be a single string")
  }
  if (!nzchar(trimws(placeholder_text))) {
    stop("placeholder_text cannot be empty")
  }
  if (!is.numeric(table_index) || length(table_index) != 1 ||
      is.na(table_index) || table_index < 1 || table_index %% 1 != 0) {
    stop("table_index must be a positive integer (1-based)")
  }

  if (!file.exists(source_docx_path)) {
    stop(paste("The source document does not exist:", source_docx_path))
  }
  if (tools::file_ext(source_docx_path) != "docx") {
    stop("source_docx_path must point to a .docx file")
  }

  validate_input_args(template_path, output_path)

  script <- system.file(
    "scripts/insert_formatted_table.py",
    package = "azreportifyr"
  )
  if (!nzchar(script) || !file.exists(script)) {
    stop("insert_formatted_table.py script not found in installed package")
  }

  args <- c(
    "run", script,
    "-s", source_docx_path,
    "-t", template_path,
    "-o", output_path,
    "-p", placeholder_text,
    "--table-index", as.character(as.integer(table_index))
  )

  paths <- get_venv_uv_paths()

  log4r::debug(.le$logger, "Running formatted table insertion script")
  tryCatch(
    {
      processx::run(
        command = paths$uv,
        args = args,
        env = c("current", VIRTUAL_ENV = paths$venv),
        error_on_status = TRUE
      )
    },
    error = function(e) {
      log4r::error(
        .le$logger,
        paste0("insert_formatted_table script failed. Status: ", e$status)
      )
      log4r::error(
        .le$logger,
        paste0("insert_formatted_table script failed. Stderr: ", e$stderr)
      )
      stop(paste(
        "insert_formatted_table script failed. Status:",
        e$status,
        "Stderr:",
        e$stderr
      ))
    }
  )

  log4r::info(.le$logger, paste0("Formatted table inserted to: ", output_path))
  tictoc::toc()
  log4r::debug(.le$logger, "Exiting insert_formatted_table function")

  invisible(output_path)
}
