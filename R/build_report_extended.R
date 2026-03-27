#' Build report (extended) with YAML input
#'
#' @description Minimal library entry point for the DOCX/YAML review-loop flow.
#'   This initial version validates file inputs and returns a status message so
#'   integrators can verify function wiring.
#' @param docx_in The file path to the input `.docx` file.
#' @param yaml_in The file path to the input `.yml` or `.yaml` file.
#' @param docx_out Optional file path to an output `.docx` file. Reserved for
#'   future extended rendering behavior.
#'
#' @return A character scalar status message.
#' @export
build_report_extended <- function(docx_in, yaml_in, docx_out = NULL) {
  if (!is.character(docx_in) || length(docx_in) != 1 || is.na(docx_in) || !nzchar(docx_in)) {
    stop("docx_in must be a non-empty character scalar", call. = FALSE)
  }
  if (!is.character(yaml_in) || length(yaml_in) != 1 || is.na(yaml_in) || !nzchar(yaml_in)) {
    stop("yaml_in must be a non-empty character scalar", call. = FALSE)
  }
  if (!file.exists(docx_in)) {
    stop("docx_in does not exist: ", docx_in, call. = FALSE)
  }
  if (!file.exists(yaml_in)) {
    stop("yaml_in does not exist: ", yaml_in, call. = FALSE)
  }
  if (!is.null(docx_out) &&
      (!is.character(docx_out) || length(docx_out) != 1 || is.na(docx_out) || !nzchar(docx_out))) {
    stop("docx_out must be NULL or a non-empty character scalar", call. = FALSE)
  }

  status <- paste0(
    "build_report_extended initialized successfully (docx_in=",
    basename(docx_in),
    ", yaml_in=",
    basename(yaml_in),
    ")"
  )
  log4r::info(.le$logger, status)
  status
}
