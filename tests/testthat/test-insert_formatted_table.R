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
