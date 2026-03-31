#' Build report (extended) with YAML input
#'
#' @description Reads a `.docx` template plus YAML values and builds an
#'   extended report version with inline text replacement and image block
#'   rendering. Hidden block markers are inserted to support reverse sync.
#'   Extended tag matching is used (`{{...}}`, `<<text:...>>`, `<<image:...>>`,
#'   `<<table:...>>`, and expanded block tags). For table insertion, `<<table:...>>`
#'   is internally mapped to temporary `{rpfy}:file.ext` placeholders only to reuse
#'   existing `add_tables()` behavior; these temporary placeholders are removed
#'   from the final output document.
#' @param docx_in The file path to the input `.docx` file.
#' @param docx_out Optional file path to an output `.docx` file. Reserved for
#'   future extended rendering behavior.
#' @param figures_path The file path to the figures and associated metadata directory.
#' @param tables_path The file path to the tables and associated metadata directory.
#' @param yaml_in The file path to the input `.yml` or `.yaml` file.
#' @param standard_footnotes_yaml Optional file path to a standard footnotes YAML.
#' @param config_yaml Optional file path to the config YAML.
#' @param add_footnotes A boolean indicating whether footnotes should be added.
#' @param docx_footnote A boolean indicating whether DOCX artifact footnotes are enabled.
#' @param include_object_path A boolean indicating whether object paths are included.
#' @param footnotes_fail_on_missing_metadata A boolean controlling strict metadata handling.
#' @param docx_table_style Optional styling object for DOCX table artifacts.
#' @param block_paragraph_styles Optional paragraph style mapping for extended
#'   block title/footnote text. Supply a named list with any of:
#'   `table_title`, `table_footnote`, `image_title`, `image_footnote`.
#' @param version Optional version label/number used to create versioned output folders
#'   and update YAML metadata.
#' @param versions_root Optional root directory for version folders. If `NULL`,
#'   defaults to `<dirname(docx_out)>/versions`.
#'
#' @return A character scalar status message.
#' @export
build_report_extended <- function(
  docx_in,
  docx_out = NULL,
  figures_path,
  tables_path,
  yaml_in,
  standard_footnotes_yaml = NULL,
  config_yaml = NULL,
  add_footnotes = TRUE,
  docx_footnote = FALSE,
  include_object_path = FALSE,
  footnotes_fail_on_missing_metadata = TRUE,
  docx_table_style = NULL,
  block_paragraph_styles = NULL,
  version = NULL,
  versions_root = NULL
) {
  log4r::debug(.le$logger, "Starting build_report_extended function")

  if (is.null(config_yaml)) {
    config_yaml <- system.file("extdata", "config.yaml", package = "azreportifyr")
    log4r::info(.le$logger, paste0("using built-in config.yaml: ", config_yaml))
  }

  validate_build_report_extended_args(
    docx_in = docx_in,
    docx_out = docx_out,
    figures_path = figures_path,
    tables_path = tables_path,
    yaml_in = yaml_in,
    version = version,
    versions_root = versions_root
  )

  base_output_path <- resolve_build_report_extended_output(docx_in, docx_out)
  output_path <- resolve_versioned_output_path(
    output_path = base_output_path,
    version = version,
    versions_root = versions_root
  )
  log4r::info(.le$logger, paste0("Output document path set: ", output_path))
  intermediate_yaml_docx <- gsub(".docx", "-intyaml.docx", output_path)
  normalized_block_styles <- normalize_block_paragraph_styles(block_paragraph_styles)

  script <- system.file("scripts/add_yaml_blocks.py", package = "azreportifyr")
  args <- c(
    "run",
    script,
    "-i",
    docx_in,
    "-o",
    intermediate_yaml_docx,
    "-y",
    yaml_in,
    "-a",
    figures_path,
    "-t",
    tables_path,
    "-c",
    config_yaml
  )
  if (!is.null(normalized_block_styles)) {
    args <- c(
      args,
      "--block-style-json",
      jsonlite::toJSON(normalized_block_styles, auto_unbox = TRUE)
    )
  }

  paths <- get_venv_uv_paths()
  result <- tryCatch(
    {
      run_build_report_extended_script(paths = paths, args = args)
    },
    error = function(e) {
      log4r::error(.le$logger, paste0("build_report_extended script failed: ", e$message))
      if (!is.null(e$stderr)) {
        log4r::error(.le$logger, paste0("stderr: ", e$stderr))
      }
      stop(
        "build_report_extended stopped: Failed to render YAML blocks into DOCX.",
        call. = FALSE
      )
    }
  )

  status <- paste0(
    "build_report_extended completed (docx_out=",
    output_path,
    ")"
  )
  if (!is.null(result$stdout) && nzchar(result$stdout)) {
    log4r::info(.le$logger, paste0("script stdout: ", result$stdout))
  }

  table_magic_strings <- extract_extended_table_magic_strings(yaml_in)
  if (length(table_magic_strings) > 0) {
    add_tables(
      docx_in = intermediate_yaml_docx,
      docx_out = output_path,
      tables_path = tables_path,
      config_yaml = config_yaml,
      docx_footnote = docx_footnote,
      docx_table_style = docx_table_style
    )
    detected_magic_strings <- extract_rpfy_magic_strings_from_doc(output_path)
    remove_magic_strings_from_doc(
      docx_in = output_path,
      docx_out = output_path,
      magic_strings = unique(c(table_magic_strings, detected_magic_strings))
    )
    unlink(intermediate_yaml_docx)
  } else {
    file.copy(intermediate_yaml_docx, output_path, overwrite = TRUE)
    unlink(intermediate_yaml_docx)
  }

  update_yaml_version_metadata(
    yaml_in = yaml_in,
    version = version,
    output_docx = output_path
  )

  log4r::info(.le$logger, status)
  log4r::debug(.le$logger, "Exiting build_report_extended function")
  status
}
