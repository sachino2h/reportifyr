validate_build_report_extended_args <- function(
  docx_in,
  docx_out,
  figures_path,
  tables_path,
  yaml_in,
  version = NULL,
  versions_root = NULL
) {
  validate_input_args(docx_in, docx_out)

  if (!is.character(yaml_in) || length(yaml_in) != 1 || is.na(yaml_in) || !nzchar(yaml_in)) {
    stop("yaml_in must be a non-empty character scalar", call. = FALSE)
  }
  if (!file.exists(yaml_in)) {
    stop("yaml_in does not exist: ", yaml_in, call. = FALSE)
  }
  if (!(tools::file_ext(yaml_in) %in% c("yml", "yaml"))) {
    stop("yaml_in must point to a .yml or .yaml file", call. = FALSE)
  }
  if (!is.character(figures_path) || length(figures_path) != 1 || is.na(figures_path) || !nzchar(figures_path)) {
    stop("figures_path must be a non-empty character scalar", call. = FALSE)
  }
  if (!is.character(tables_path) || length(tables_path) != 1 || is.na(tables_path) || !nzchar(tables_path)) {
    stop("tables_path must be a non-empty character scalar", call. = FALSE)
  }
  if (!dir.exists(figures_path)) {
    stop("figures_path does not exist: ", figures_path, call. = FALSE)
  }
  if (!is.null(version)) {
    if ((is.character(version) && (length(version) != 1 || is.na(version) || !nzchar(trimws(version)))) ||
        (is.numeric(version) && (length(version) != 1 || is.na(version))) ||
        (!is.character(version) && !is.numeric(version))) {
      stop("version must be NULL, a non-empty string, or a single number", call. = FALSE)
    }
  }
  if (!is.null(versions_root) &&
      (!is.character(versions_root) || length(versions_root) != 1 || is.na(versions_root) || !nzchar(versions_root))) {
    stop("versions_root must be NULL or a non-empty character scalar", call. = FALSE)
  }
}

resolve_build_report_extended_output <- function(docx_in, docx_out) {
  if (!is.null(docx_out)) {
    return(docx_out)
  }
  make_doc_dirs(docx_in = docx_in)$doc_draft
}

normalize_block_paragraph_styles <- function(block_paragraph_styles) {
  if (is.null(block_paragraph_styles)) {
    return(NULL)
  }

  if (!is.list(block_paragraph_styles) || is.data.frame(block_paragraph_styles)) {
    stop("block_paragraph_styles must be NULL or a named list", call. = FALSE)
  }

  if (length(block_paragraph_styles) == 0) {
    return(NULL)
  }

  style_names <- names(block_paragraph_styles)
  if (is.null(style_names) || any(!nzchar(style_names))) {
    stop("block_paragraph_styles must be a named list", call. = FALSE)
  }

  allowed_fields <- c(
    "table_title",
    "table_footnote",
    "image_title",
    "image_footnote"
  )
  unknown_fields <- setdiff(style_names, allowed_fields)
  if (length(unknown_fields) > 0) {
    stop(
      paste0(
        "Unknown block_paragraph_styles field(s): ",
        paste(unknown_fields, collapse = ", ")
      ),
      call. = FALSE
    )
  }

  out <- list()
  for (nm in intersect(allowed_fields, style_names)) {
    val <- block_paragraph_styles[[nm]]
    if (!is.character(val) || length(val) != 1 || is.na(val) || !nzchar(trimws(val))) {
      stop(
        paste0("block_paragraph_styles$", nm, " must be a single non-empty string"),
        call. = FALSE
      )
    }
    out[[nm]] <- trimws(val)
  }

  if (length(out) == 0) {
    return(NULL)
  }

  out
}

normalize_version_label <- function(version) {
  if (is.null(version)) {
    return(NULL)
  }
  if (is.numeric(version)) {
    return(sprintf("v%03d", as.integer(version)))
  }
  v <- trimws(as.character(version))
  if (grepl("^[0-9]+$", v)) {
    return(sprintf("v%03d", as.integer(v)))
  }
  if (grepl("^v[0-9]+$", tolower(v))) {
    return(tolower(v))
  }
  v
}

resolve_versioned_output_path <- function(output_path, version = NULL, versions_root = NULL) {
  version_label <- normalize_version_label(version)
  if (is.null(version_label)) {
    return(output_path)
  }

  root <- if (is.null(versions_root)) {
    file.path(dirname(output_path), "versions")
  } else {
    versions_root
  }
  version_dir <- file.path(root, version_label)
  dir.create(version_dir, recursive = TRUE, showWarnings = FALSE)
  file.path(version_dir, basename(output_path))
}

update_yaml_version_metadata <- function(yaml_in, version = NULL, output_docx = NULL) {
  version_label <- normalize_version_label(version)
  if (is.null(version_label)) {
    return(invisible(NULL))
  }

  text <- paste(readLines(yaml_in, warn = FALSE, encoding = "UTF-8"), collapse = "\n")
  eol <- if (grepl("\r\n", text, fixed = TRUE)) "\r\n" else "\n"
  has_trailing_newline <- grepl(paste0(eol, "$"), text)
  lines <- strsplit(text, "\r\n|\n", perl = TRUE)[[1]]
  if (length(lines) == 0) {
    lines <- character()
  }

  version_number <- NA_integer_
  if (is.numeric(version)) {
    version_number <- as.integer(version)
  } else {
    digits <- gsub("[^0-9]", "", as.character(version))
    if (nzchar(digits)) {
      version_number <- as.integer(digits)
    }
  }

  key_values <- list(
    current_version = paste0("'", version_label, "'")
  )
  if (!is.na(version_number)) {
    key_values$version_number <- as.character(version_number)
  }

  metadata_idx <- which(grepl("^metadata:[[:space:]]*$", lines))[1]
  if (is.na(metadata_idx)) {
    if (length(lines) > 0 && nzchar(lines[length(lines)])) {
      lines <- c(lines, "")
    }
    lines <- c(lines, "metadata:")
    for (nm in names(key_values)) {
      lines <- c(lines, paste0("  ", nm, ": ", key_values[[nm]]))
    }
  } else {
    end_idx <- length(lines) + 1L
    if (metadata_idx < length(lines)) {
      for (i in seq(metadata_idx + 1L, length(lines))) {
        if (grepl("^[^[:space:]#][^:]*:[[:space:]]*", lines[[i]])) {
          end_idx <- i
          break
        }
      }
    }
    block_start <- metadata_idx + 1L
    block_end <- end_idx - 1L

    for (nm in names(key_values)) {
      pat <- paste0("^\\s{2}['\"]?", nm, "['\"]?\\s*:")
      idx <- if (block_start <= block_end) {
        which(grepl(pat, lines[block_start:block_end]))[1]
      } else {
        NA_integer_
      }
      new_line <- paste0("  ", nm, ": ", key_values[[nm]])
      if (!is.na(idx)) {
        lines[block_start + idx - 1L] <- new_line
      } else {
        insert_at <- end_idx - 1L
        if (insert_at <= 0L) {
          lines <- c(new_line, lines)
        } else if (insert_at >= length(lines)) {
          lines <- c(lines, new_line)
        } else {
          lines <- append(lines, new_line, after = insert_at)
        }
        end_idx <- end_idx + 1L
        block_end <- block_end + 1L
      }
    }
  }

  out <- paste(lines, collapse = eol)
  if (has_trailing_newline || !nzchar(out)) {
    out <- paste0(out, eol)
  }
  writeLines(out, yaml_in, useBytes = TRUE)
  invisible(NULL)
}

run_build_report_extended_script <- function(paths, args) {
  processx::run(
    command = paths$uv,
    args = args,
    env = c("current", VIRTUAL_ENV = paths$venv),
    error_on_status = TRUE
  )
}

extract_extended_table_magic_strings <- function(yaml_in) {
  yml <- yaml::read_yaml(yaml_in)
  blocks <- yml$blocks
  if (is.null(blocks) || length(blocks) == 0) {
    return(character())
  }

  out <- character()
  for (nm in names(blocks)) {
    block <- blocks[[nm]]
    block_type <- tolower(trimws(as.character(block$type %||% "")))
    if (nzchar(block_type) && !(block_type %in% c("table", "block"))) {
      next
    }
    files <- block$files
    if (is.null(files)) {
      next
    }
    if (is.character(files)) {
      file_entries <- files
    } else if (is.list(files)) {
      file_entries <- unlist(files, use.names = FALSE)
    } else {
      next
    }
    if (length(file_entries) == 0) {
      next
    }
    chosen <- NULL
    for (entry in file_entries) {
      file_name <- trimws(as.character(entry))
      file_name <- sub("<.*$", "", file_name)
      file_name <- basename(file_name)
      if (!nzchar(file_name)) {
        next
      }
      ext <- tolower(tools::file_ext(file_name))
      if (ext %in% c("csv", "rds", "docx")) {
        chosen <- file_name
        break
      }
    }
    if (!is.null(chosen)) {
      out <- c(out, paste0("{rpfy}:", chosen))
    }
  }

  unique(out)
}

`%||%` <- function(x, y) {
  if (is.null(x)) y else x
}

escape_regex_literal <- function(x) {
  gsub("([][{}()+*^$|\\\\?.])", "\\\\\\1", x)
}

remove_magic_strings_from_doc <- function(docx_in, docx_out, magic_strings) {
  if (length(magic_strings) == 0) {
    if (normalizePath(docx_in, winslash = "/", mustWork = FALSE) !=
        normalizePath(docx_out, winslash = "/", mustWork = FALSE)) {
      file.copy(docx_in, docx_out, overwrite = TRUE)
    }
    return(invisible(NULL))
  }

  tmp_out <- if (normalizePath(docx_in, winslash = "/", mustWork = FALSE) ==
    normalizePath(docx_out, winslash = "/", mustWork = FALSE)) {
    tempfile(fileext = ".docx")
  } else {
    docx_out
  }

  document <- officer::read_docx(docx_in)
  for (magic in unique(magic_strings)) {
    if (!is.character(magic) || length(magic) != 1 || !nzchar(magic)) {
      next
    }
    escaped_magic <- escape_regex_literal(magic)
    document <- officer::body_replace_all_text(
      x = document,
      old_value = escaped_magic,
      new_value = "",
      only_at_cursor = FALSE,
      warn = FALSE
    )
  }

  print(document, target = tmp_out)
  if (!identical(tmp_out, docx_out)) {
    file.copy(tmp_out, docx_out, overwrite = TRUE)
    unlink(tmp_out)
  }
  invisible(NULL)
}

extract_rpfy_magic_strings_from_doc <- function(docx_in) {
  if (!file.exists(docx_in)) {
    return(character())
  }
  doc <- officer::read_docx(docx_in)
  summary <- officer::docx_summary(doc)
  if (is.null(summary$text) || length(summary$text) == 0) {
    return(character())
  }
  pattern <- "\\{rpfy\\}:.*?\\.[A-Za-z0-9]+"
  out <- unlist(
    lapply(summary$text, function(x) {
      if (is.na(x) || !nzchar(x)) return(character())
      regmatches(x, gregexpr(pattern, x, perl = TRUE))[[1]]
    }),
    use.names = FALSE
  )
  out <- out[out != "-1"]
  unique(out)
}
