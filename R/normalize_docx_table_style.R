normalize_docx_table_style <- function(docx_table_style, table_file = NULL) {
  if (is.null(docx_table_style)) {
    return(NULL)
  }

  if (!is.list(docx_table_style) || is.data.frame(docx_table_style)) {
    stop("docx_table_style must be NULL or a named list")
  }

  if (length(docx_table_style) == 0) {
    return(NULL)
  }

  style_names <- names(docx_table_style)
  if (is.null(style_names) || any(!nzchar(style_names))) {
    stop("docx_table_style must be a named list")
  }

  alias_map <- c(
    table_header = "header_paragraph_style",
    table_split_row = "split_row_paragraph_style",
    table_first_column = "first_column_paragraph_style",
    table_normal_cell = "body_paragraph_style",
    table_footnote = "footnote_paragraph_style"
  )

  alias_keys <- intersect(style_names, names(alias_map))
  if (length(alias_keys) > 0) {
    for (alias_key in alias_keys) {
      target_key <- unname(alias_map[[alias_key]])
      if (target_key %in% style_names) {
        stop(
          paste0(
            "docx_table_style cannot include both '",
            alias_key,
            "' and '",
            target_key,
            "'"
          )
        )
      }
      names(docx_table_style)[names(docx_table_style) == alias_key] <- target_key
      style_names[style_names == alias_key] <- target_key
    }
  }

  direct_keys <- c(
    "header_fill",
    "header_bold",
    "font_family",
    "font_size",
    "header_rows",
    "extract_table_footnotes",
    "header_paragraph_style",
    "split_row_paragraph_style",
    "first_column_paragraph_style",
    "body_paragraph_style",
    "footnote_paragraph_style"
  )

  if (any(style_names %in% direct_keys)) {
    style <- validate_docx_table_style_list(
      docx_table_style,
      "docx_table_style"
    )
    if (is.null(style$header_rows)) {
      style$header_rows <- 1L
    }
    return(style)
  }

  allowed_top_level <- c("default", "per_table")
  unknown_top_level <- setdiff(style_names, allowed_top_level)
  if (length(unknown_top_level) > 0) {
    stop(
      paste0(
        "Unknown docx_table_style field(s): ",
        paste(unknown_top_level, collapse = ", ")
      )
    )
  }

  default_style <- docx_table_style$default
  if (is.null(default_style)) {
    default_style <- list()
  }
  if (!is.list(default_style) || is.data.frame(default_style)) {
    stop("docx_table_style$default must be a named list")
  }
  default_style <- validate_docx_table_style_list(
    default_style,
    "docx_table_style$default"
  )

  per_table <- docx_table_style$per_table
  if (is.null(per_table)) {
    per_table <- list()
  }
  if (!is.list(per_table) || is.data.frame(per_table)) {
    stop("docx_table_style$per_table must be a named list")
  }
  if (length(per_table) > 0 && (is.null(names(per_table)) ||
      any(!nzchar(names(per_table))))) {
    stop("docx_table_style$per_table must be a named list")
  }

  style <- default_style
  if (!is.null(table_file) && length(per_table) > 0) {
    table_key_candidates <- c(basename(table_file), table_file)
    matching_key <- table_key_candidates[table_key_candidates %in% names(per_table)]

    if (length(matching_key) > 0) {
      table_style <- per_table[[matching_key[[1]]]]
      if (!is.list(table_style) || is.data.frame(table_style)) {
        stop(
          paste0(
            "docx_table_style$per_table[['",
            matching_key[[1]],
            "']] must be a named list"
          )
        )
      }

      table_style <- validate_docx_table_style_list(
        table_style,
        paste0("docx_table_style$per_table[['", matching_key[[1]], "']]")
      )
      style <- utils::modifyList(style, table_style)
    }
  }

  if (length(style) == 0) {
    return(NULL)
  }

  if (is.null(style$header_rows)) {
    style$header_rows <- 1L
  }

  style
}


validate_docx_table_style_list <- function(style, arg_name) {
  if (length(style) == 0) {
    return(list())
  }

  if (is.null(names(style)) || any(!nzchar(names(style)))) {
    stop(paste0(arg_name, " must be a named list"))
  }

  alias_map <- c(
    table_header = "header_paragraph_style",
    table_split_row = "split_row_paragraph_style",
    table_first_column = "first_column_paragraph_style",
    table_normal_cell = "body_paragraph_style",
    table_footnote = "footnote_paragraph_style"
  )

  style_names <- names(style)
  aliased <- intersect(style_names, names(alias_map))
  if (length(aliased) > 0) {
    for (alias_key in aliased) {
      target_key <- unname(alias_map[[alias_key]])
      if (target_key %in% style_names) {
        stop(
          paste0(
            arg_name,
            " cannot include both '",
            alias_key,
            "' and '",
            target_key,
            "'"
          )
        )
      }
      names(style)[names(style) == alias_key] <- target_key
      style_names[style_names == alias_key] <- target_key
    }
  }

  allowed_fields <- c(
    "header_fill",
    "header_bold",
    "font_family",
    "font_size",
    "header_rows",
    "extract_table_footnotes",
    "header_paragraph_style",
    "split_row_paragraph_style",
    "first_column_paragraph_style",
    "body_paragraph_style",
    "footnote_paragraph_style"
  )

  unknown_fields <- setdiff(names(style), allowed_fields)
  if (length(unknown_fields) > 0) {
    stop(
      paste0(
        "Unknown ",
        arg_name,
        " field(s): ",
        paste(unknown_fields, collapse = ", ")
      )
    )
  }

  normalized_style <- style

  if ("header_fill" %in% names(normalized_style)) {
    fill_value <- normalized_style$header_fill
    if (!is.character(fill_value) || length(fill_value) != 1 ||
        is.na(fill_value) || !nzchar(trimws(fill_value))) {
      stop(paste0(arg_name, "$header_fill must be a single color value"))
    }

    fill_value <- trimws(fill_value)
    if (grepl("^#?[0-9A-Fa-f]{6}$", fill_value)) {
      fill_value <- paste0("#", sub("^#", "", fill_value))
    } else if (grepl("^#?[0-9A-Fa-f]{3}$", fill_value)) {
      short_hex <- sub("^#", "", fill_value)
      short_hex_chars <- strsplit(short_hex, "", fixed = TRUE)[[1]]
      fill_value <- paste0("#", paste0(short_hex_chars, short_hex_chars, collapse = ""))
    }

    rgb_value <- tryCatch(
      grDevices::col2rgb(fill_value),
      error = function(...) NULL
    )
    if (is.null(rgb_value)) {
      stop(
        paste0(
          arg_name,
          "$header_fill must be a valid color name or hex value"
        )
      )
    }

    normalized_style$header_fill <- toupper(
      grDevices::rgb(
        rgb_value[1, 1],
        rgb_value[2, 1],
        rgb_value[3, 1],
        maxColorValue = 255
      )
    )
    normalized_style$header_fill <- sub("^#", "", normalized_style$header_fill)
  }

  if ("header_bold" %in% names(normalized_style)) {
    if (!is.logical(normalized_style$header_bold) ||
        length(normalized_style$header_bold) != 1 ||
        is.na(normalized_style$header_bold)) {
      stop(paste0(arg_name, "$header_bold must be TRUE or FALSE"))
    }
  }

  if ("font_family" %in% names(normalized_style)) {
    if (!is.character(normalized_style$font_family) ||
        length(normalized_style$font_family) != 1 ||
        is.na(normalized_style$font_family) ||
        !nzchar(trimws(normalized_style$font_family))) {
      stop(paste0(arg_name, "$font_family must be a single non-empty string"))
    }
  }

  if ("font_size" %in% names(normalized_style)) {
    if (!is.numeric(normalized_style$font_size) ||
        length(normalized_style$font_size) != 1 ||
        is.na(normalized_style$font_size) ||
        normalized_style$font_size <= 0) {
      stop(paste0(arg_name, "$font_size must be a positive number"))
    }

    normalized_style$font_size <- as.numeric(normalized_style$font_size)
  }

  if ("header_rows" %in% names(normalized_style)) {
    if (!is.numeric(normalized_style$header_rows) ||
        length(normalized_style$header_rows) != 1 ||
        is.na(normalized_style$header_rows) ||
        normalized_style$header_rows < 0 ||
        normalized_style$header_rows %% 1 != 0) {
      stop(paste0(arg_name, "$header_rows must be a non-negative integer"))
    }

    normalized_style$header_rows <- as.integer(normalized_style$header_rows)
  }

  if ("extract_table_footnotes" %in% names(normalized_style)) {
    if (!is.logical(normalized_style$extract_table_footnotes) ||
        length(normalized_style$extract_table_footnotes) != 1 ||
        is.na(normalized_style$extract_table_footnotes)) {
      stop(paste0(arg_name, "$extract_table_footnotes must be TRUE or FALSE"))
    }
  }

  paragraph_style_fields <- c(
    "header_paragraph_style",
    "split_row_paragraph_style",
    "first_column_paragraph_style",
    "body_paragraph_style",
    "footnote_paragraph_style"
  )
  for (fld in paragraph_style_fields) {
    if (fld %in% names(normalized_style)) {
      value <- normalized_style[[fld]]
      if (!is.character(value) || length(value) != 1 ||
          is.na(value) || !nzchar(trimws(value))) {
        stop(paste0(arg_name, "$", fld, " must be a single non-empty string"))
      }
      normalized_style[[fld]] <- trimws(value)
    }
  }

  normalized_style
}
