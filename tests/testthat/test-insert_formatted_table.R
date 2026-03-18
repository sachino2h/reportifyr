test_that("insert_formatted_table validates source path", {
  template <- withr::local_tempfile(fileext = ".docx")
  file.create(template)
  output <- withr::local_tempfile(fileext = ".docx")

  expect_error(
    insert_formatted_table(
      source_docx_path = "missing.docx",
      template_path = template,
      output_path = output,
      placeholder_text = "{rpfy}:table.csv"
    ),
    "The source document does not exist"
  )
})

test_that("insert_formatted_table validates source extension", {
  source <- withr::local_tempfile(fileext = ".txt")
  file.create(source)
  template <- withr::local_tempfile(fileext = ".docx")
  file.create(template)
  output <- withr::local_tempfile(fileext = ".docx")

  expect_error(
    insert_formatted_table(
      source_docx_path = source,
      template_path = template,
      output_path = output,
      placeholder_text = "{rpfy}:table.csv"
    ),
    "source_docx_path must point to a .docx file"
  )
})

test_that("insert_formatted_table validates placeholder_text", {
  source <- withr::local_tempfile(fileext = ".docx")
  file.create(source)
  template <- withr::local_tempfile(fileext = ".docx")
  file.create(template)
  output <- withr::local_tempfile(fileext = ".docx")

  expect_error(
    insert_formatted_table(
      source_docx_path = source,
      template_path = template,
      output_path = output,
      placeholder_text = " "
    ),
    "placeholder_text cannot be empty"
  )
})

test_that("insert_formatted_table validates table_index", {
  source <- withr::local_tempfile(fileext = ".docx")
  file.create(source)
  template <- withr::local_tempfile(fileext = ".docx")
  file.create(template)
  output <- withr::local_tempfile(fileext = ".docx")

  expect_error(
    insert_formatted_table(
      source_docx_path = source,
      template_path = template,
      output_path = output,
      placeholder_text = "{rpfy}:table.csv",
      table_index = 0
    ),
    "table_index must be a positive integer"
  )
})

test_that("insert_formatted_table validates include_footnote", {
  source <- withr::local_tempfile(fileext = ".docx")
  file.create(source)
  template <- withr::local_tempfile(fileext = ".docx")
  file.create(template)
  output <- withr::local_tempfile(fileext = ".docx")

  expect_error(
    insert_formatted_table(
      source_docx_path = source,
      template_path = template,
      output_path = output,
      placeholder_text = "{rpfy}:table.csv",
      include_footnote = "yes"
    ),
    "include_footnote must be TRUE or FALSE"
  )
})

test_that("insert_formatted_table passes style JSON to the script", {
  source <- withr::local_tempfile(fileext = ".docx")
  file.create(source)
  template <- withr::local_tempfile(fileext = ".docx")
  file.create(template)
  output <- withr::local_tempfile(fileext = ".docx")
  calls <- new.env(parent = emptyenv())
  calls$args <- NULL

  mockery::stub(insert_formatted_table, "validate_input_args", function(...) NULL)
  mockery::stub(insert_formatted_table, "system.file", function(...) "script.py")
  mockery::stub(insert_formatted_table, "file.exists", function(...) TRUE)
  mockery::stub(
    insert_formatted_table,
    "get_venv_uv_paths",
    function(...) list(uv = "uv", venv = "venv")
  )
  mockery::stub(insert_formatted_table, "processx::run", function(args, ...) {
    calls$args <- args
    list(status = 0, stdout = "", stderr = "")
  })

  expect_silent(
    insert_formatted_table(
      source_docx_path = source,
      template_path = template,
      output_path = output,
      placeholder_text = "{rpfy}:table.csv",
      docx_table_style = list(
        header_fill = "#D9EAF7",
        header_bold = TRUE,
        font_family = "Times New Roman",
        font_size = 10
      )
    )
  )

  style_arg_index <- match("--style-json", calls$args)
  expect_false(is.na(style_arg_index))

  style_json <- calls$args[[style_arg_index + 1]]
  style <- jsonlite::fromJSON(style_json, simplifyVector = TRUE)

  expect_equal(style$header_fill, "D9EAF7")
  expect_true(style$header_bold)
  expect_equal(style$font_family, "Times New Roman")
  expect_equal(style$font_size, 10)
  expect_equal(style$header_rows, 1)
})
