#' Replace a magic string in a Microsoft Word file with text
#'
#' @description Reads an input `.docx` file and replaces all occurrences of
#'   `magic_string` with `replacement_text`, then writes the updated file.
#' @param docx_in The file path to the input `.docx` file.
#' @param docx_out The file path to the output `.docx` file to save to.
#' @param magic_string Magic string to replace (for example, `"{rpfy}:note.txt"`).
#' @param replacement_text Text to insert in place of `magic_string`.
#'
#' @return ()
#' @export
#'
#' @examples \dontrun{
#' replace_magic_string_text(
#'   docx_in = "report-draft.docx",
#'   docx_out = "report-updated.docx",
#'   magic_string = "{rpfy}:custom.txt",
#'   replacement_text = "Custom replacement text"
#' )
#' }
replace_magic_string_text <- function(
  docx_in,
  docx_out,
  magic_string,
  replacement_text
) {
  tictoc::tic()
  log4r::debug(.le$logger, "Starting replace_magic_string_text function")
  validate_input_args(docx_in, docx_out)

  if (!is.character(magic_string) || length(magic_string) != 1 || is.na(magic_string)) {
    stop("magic_string must be a single non-missing string")
  }
  if (!nzchar(trimws(magic_string))) {
    stop("magic_string cannot be empty")
  }
  if (
    !is.character(replacement_text) ||
      length(replacement_text) != 1 ||
      is.na(replacement_text)
  ) {
    stop("replacement_text must be a single non-missing string")
  }

  log4r::info(.le$logger, paste0("Reading input document: ", docx_in))
  document <- officer::read_docx(docx_in)

  document <- officer::body_replace_all_text(
    x = document,
    old_value = magic_string,
    new_value = replacement_text,
    only_at_cursor = FALSE,
    warn = FALSE
  )

  print(document, target = docx_out)
  log4r::info(.le$logger, paste0("Saved updated document to: ", docx_out))
  log4r::debug(.le$logger, "Exiting replace_magic_string_text function")
  tictoc::toc()
}
