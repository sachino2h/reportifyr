#' Extract regex pattern matches from a Word document
#'
#' @description Reads an input `.docx` file and extracts matches for one or more
#'   user-supplied regular expression patterns from the document text.
#' @param docx_path The file path to the input `.docx` file.
#' @param patterns A character vector of regular expression patterns to match.
#'
#' @return If `patterns` has length 1, returns a character vector of unique
#'   matches. If `patterns` has length greater than 1, returns a named list of
#'   unique matches for each pattern.
#' @export
#'
#' @examples \dontrun{
#' extract_magic_strings(
#'   docx_path = "report-draft.docx",
#'   patterns = c("<<[^<>]+>>", "\\\\{rpfy\\\\}:[^[:space:]]+")
#' )
#' }
extract_magic_strings <- function(docx_path, patterns) {
  tictoc::tic()
  log4r::debug(.le$logger, "Starting extract_magic_strings function")

  validate_input_args(docx_path, NULL)

  if (!is.character(patterns) || length(patterns) < 1 || any(is.na(patterns))) {
    stop("patterns must be a non-missing character vector")
  }
  if (any(!nzchar(trimws(patterns)))) {
    stop("patterns cannot contain empty strings")
  }

  log4r::info(.le$logger, paste0("Reading input document: ", docx_path))
  document <- officer::read_docx(docx_path)
  doc_summary <- officer::docx_summary(document)

  text_values <- doc_summary$text
  text_values <- text_values[!is.na(text_values) & nzchar(text_values)]

  extract_pattern <- function(pattern) {
    matches <- regmatches(
      text_values,
      gregexpr(pattern, text_values, perl = TRUE)
    )
    unique(unlist(matches, use.names = FALSE))
  }

  results <- lapply(patterns, extract_pattern)
  result_names <- names(patterns)
  if (is.null(result_names) || any(!nzchar(result_names))) {
    result_names <- make.unique(patterns)
  }
  names(results) <- result_names

  log4r::debug(.le$logger, "Exiting extract_magic_strings function")
  tictoc::toc()

  if (length(results) == 1) {
    return(results[[1]])
  }

  results
}
