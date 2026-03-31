test_that("build_report_extended returns wiring success message", {
  docx_in <- tempfile(fileext = ".docx")
  yaml_in <- tempfile(fileext = ".yaml")
  figures_path <- tempfile("figures-")
  tables_path <- tempfile("tables-")
  file.create(docx_in)
  writeLines("inline: {}", con = yaml_in)
  dir.create(figures_path)

  mockery::stub(build_report_extended, "resolve_build_report_extended_output", function(...) "out.docx")
  mockery::stub(build_report_extended, "resolve_versioned_output_path", function(...) "out.docx")
  mockery::stub(build_report_extended, "get_venv_uv_paths", function(...) {
    list(uv = "uv", venv = "venv")
  })
  mockery::stub(build_report_extended, "run_build_report_extended_script", function(...) {
    list(status = 0L, stdout = "ok", stderr = "")
  })
  mockery::stub(build_report_extended, "extract_extended_table_magic_strings", function(...) character())
  mockery::stub(build_report_extended, "extract_rpfy_magic_strings_from_doc", function(...) character())
  mockery::stub(build_report_extended, "update_yaml_version_metadata", function(...) invisible(NULL))
  mockery::stub(build_report_extended, "file.copy", function(...) TRUE)
  mockery::stub(build_report_extended, "unlink", function(...) 0L)

  out <- build_report_extended(
    docx_in = docx_in,
    docx_out = "out.docx",
    figures_path = figures_path,
    tables_path = tables_path,
    yaml_in = yaml_in
  )

  expect_type(out, "character")
  expect_length(out, 1)
  expect_match(out, "build_report_extended completed", fixed = TRUE)
  expect_match(out, "out.docx", fixed = TRUE)
})

test_that("build_report_extended validates missing files", {
  expect_error(
    build_report_extended(
      docx_in = "missing.docx",
      figures_path = "figures",
      tables_path = "tables",
      yaml_in = "missing.yaml"
    ),
    "docx_in does not exist",
    fixed = TRUE
  )
})

test_that("build_report_extended validates missing figures_path directory", {
  docx_in <- tempfile(fileext = ".docx")
  yaml_in <- tempfile(fileext = ".yaml")
  file.create(docx_in)
  writeLines("inline: {}", con = yaml_in)

  expect_error(
    build_report_extended(
      docx_in = docx_in,
      figures_path = tempfile("missing-figures-"),
      tables_path = "tables",
      yaml_in = yaml_in
    ),
    "figures_path does not exist",
    fixed = TRUE
  )
})

test_that("build_report_extended runs table pipeline when table blocks exist", {
  docx_in <- tempfile(fileext = ".docx")
  yaml_in <- tempfile(fileext = ".yaml")
  figures_path <- tempfile("figures-")
  tables_path <- tempfile("tables-")
  file.create(docx_in)
  writeLines("blocks: {}", con = yaml_in)
  dir.create(figures_path)

  calls <- new.env(parent = emptyenv())
  calls$add_tables <- 0L
  calls$remove_magic <- 0L
  calls$extract_doc_magic <- 0L

  mockery::stub(build_report_extended, "resolve_build_report_extended_output", function(...) "out.docx")
  mockery::stub(build_report_extended, "resolve_versioned_output_path", function(...) "out.docx")
  mockery::stub(build_report_extended, "get_venv_uv_paths", function(...) {
    list(uv = "uv", venv = "venv")
  })
  mockery::stub(build_report_extended, "run_build_report_extended_script", function(...) {
    list(status = 0L, stdout = "ok", stderr = "")
  })
  mockery::stub(build_report_extended, "extract_extended_table_magic_strings", function(...) "{rpfy}:table1.csv")
  mockery::stub(build_report_extended, "extract_rpfy_magic_strings_from_doc", function(...) {
    calls$extract_doc_magic <- calls$extract_doc_magic + 1L
    "{rpfy}:table1.csv"
  })
  mockery::stub(build_report_extended, "add_tables", function(...) {
    calls$add_tables <- calls$add_tables + 1L
    invisible(NULL)
  })
  mockery::stub(build_report_extended, "remove_magic_strings_from_doc", function(...) {
    calls$remove_magic <- calls$remove_magic + 1L
    invisible(NULL)
  })
  mockery::stub(build_report_extended, "update_yaml_version_metadata", function(...) invisible(NULL))
  mockery::stub(build_report_extended, "unlink", function(...) 0L)

  out <- build_report_extended(
    docx_in = docx_in,
    docx_out = "out.docx",
    figures_path = figures_path,
    tables_path = tables_path,
    yaml_in = yaml_in
  )

  expect_match(out, "build_report_extended completed", fixed = TRUE)
  expect_equal(calls$add_tables, 1L)
  expect_equal(calls$remove_magic, 1L)
  expect_equal(calls$extract_doc_magic, 1L)
})

test_that("build_report_extended uses versioned output and updates yaml metadata", {
  docx_in <- tempfile(fileext = ".docx")
  yaml_in <- tempfile(fileext = ".yaml")
  figures_path <- tempfile("figures-")
  tables_path <- tempfile("tables-")
  file.create(docx_in)
  writeLines("inline: {}", con = yaml_in)
  dir.create(figures_path)

  calls <- new.env(parent = emptyenv())
  calls$resolved_output <- NULL
  calls$metadata_version <- NULL

  mockery::stub(build_report_extended, "resolve_build_report_extended_output", function(...) "base-out.docx")
  mockery::stub(build_report_extended, "resolve_versioned_output_path", function(...) {
    calls$resolved_output <- "versions/v005/base-out.docx"
    calls$resolved_output
  })
  mockery::stub(build_report_extended, "get_venv_uv_paths", function(...) {
    list(uv = "uv", venv = "venv")
  })
  mockery::stub(build_report_extended, "run_build_report_extended_script", function(...) {
    list(status = 0L, stdout = "ok", stderr = "")
  })
  mockery::stub(build_report_extended, "extract_extended_table_magic_strings", function(...) character())
  mockery::stub(build_report_extended, "extract_rpfy_magic_strings_from_doc", function(...) character())
  mockery::stub(build_report_extended, "file.copy", function(...) TRUE)
  mockery::stub(build_report_extended, "unlink", function(...) 0L)
  mockery::stub(build_report_extended, "update_yaml_version_metadata", function(yaml_in, version, output_docx) {
    calls$metadata_version <- version
    calls$resolved_output <- output_docx
    invisible(NULL)
  })

  out <- build_report_extended(
    docx_in = docx_in,
    figures_path = figures_path,
    tables_path = tables_path,
    yaml_in = yaml_in,
    version = 5
  )

  expect_match(out, "versions/v005/base-out.docx", fixed = TRUE)
  expect_equal(calls$metadata_version, 5)
  expect_equal(calls$resolved_output, "versions/v005/base-out.docx")
})

test_that("build_report_extended passes block paragraph styles to yaml renderer", {
  docx_in <- tempfile(fileext = ".docx")
  yaml_in <- tempfile(fileext = ".yaml")
  figures_path <- tempfile("figures-")
  tables_path <- tempfile("tables-")
  file.create(docx_in)
  writeLines("inline: {}", con = yaml_in)
  dir.create(figures_path)

  calls <- new.env(parent = emptyenv())
  calls$args <- NULL

  mockery::stub(build_report_extended, "resolve_build_report_extended_output", function(...) "out.docx")
  mockery::stub(build_report_extended, "resolve_versioned_output_path", function(...) "out.docx")
  mockery::stub(build_report_extended, "get_venv_uv_paths", function(...) {
    list(uv = "uv", venv = "venv")
  })
  mockery::stub(build_report_extended, "run_build_report_extended_script", function(paths, args) {
    calls$args <- args
    list(status = 0L, stdout = "ok", stderr = "")
  })
  mockery::stub(build_report_extended, "extract_extended_table_magic_strings", function(...) character())
  mockery::stub(build_report_extended, "extract_rpfy_magic_strings_from_doc", function(...) character())
  mockery::stub(build_report_extended, "update_yaml_version_metadata", function(...) invisible(NULL))
  mockery::stub(build_report_extended, "file.copy", function(...) TRUE)
  mockery::stub(build_report_extended, "unlink", function(...) 0L)

  out <- build_report_extended(
    docx_in = docx_in,
    docx_out = "out.docx",
    figures_path = figures_path,
    tables_path = tables_path,
    yaml_in = yaml_in,
    block_paragraph_styles = list(
      table_title = "Table Title",
      table_footnote = "Table Footnote Info",
      image_title = "Figure Title",
      image_footnote = "Figure Footnote Info"
    )
  )

  expect_match(out, "build_report_extended completed", fixed = TRUE)
  style_arg_index <- which(calls$args == "--block-style-json")
  expect_equal(length(style_arg_index), 1L)
  style_json <- calls$args[[style_arg_index + 1L]]
  style <- jsonlite::fromJSON(style_json, simplifyVector = TRUE)
  expect_equal(style$table_title, "Table Title")
  expect_equal(style$table_footnote, "Table Footnote Info")
  expect_equal(style$image_title, "Figure Title")
  expect_equal(style$image_footnote, "Figure Footnote Info")
})

test_that("build_report_extended validates unknown block paragraph style fields", {
  docx_in <- tempfile(fileext = ".docx")
  yaml_in <- tempfile(fileext = ".yaml")
  figures_path <- tempfile("figures-")
  file.create(docx_in)
  writeLines("inline: {}", con = yaml_in)
  dir.create(figures_path)

  expect_error(
    build_report_extended(
      docx_in = docx_in,
      figures_path = figures_path,
      tables_path = "tables",
      yaml_in = yaml_in,
      block_paragraph_styles = list(title = "Table Title")
    ),
    "Unknown block_paragraph_styles field",
    fixed = TRUE
  )
})
