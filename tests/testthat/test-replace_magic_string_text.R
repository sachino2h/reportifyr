test_that("replace_magic_string_text replaces magic string with text", {
  skip_if_not_installed("officer")

  docx_in <- withr::local_tempfile(fileext = ".docx")
  docx_out <- withr::local_tempfile(fileext = ".docx")

  document <- officer::read_docx()
  document <- officer::body_add_par(document, "Header text")
  document <- officer::body_add_par(document, "{rpfy}:custom.txt")
  print(document, target = docx_in)

  expect_silent(
    replace_magic_string_text(
      docx_in = docx_in,
      docx_out = docx_out,
      magic_string = "{rpfy}:custom.txt",
      replacement_text = "Custom replacement text"
    )
  )

  out_doc <- officer::read_docx(docx_out)
  out_summary <- officer::docx_summary(out_doc)
  out_text <- out_summary$text

  expect_true(any(grepl("Custom replacement text", out_text, fixed = TRUE)))
  expect_false(any(grepl("{rpfy}:custom.txt", out_text, fixed = TRUE)))
})

test_that("replace_magic_string_text validates magic_string", {
  docx_in <- withr::local_tempfile(fileext = ".docx")
  docx_out <- withr::local_tempfile(fileext = ".docx")
  file.create(docx_in)

  expect_error(
    replace_magic_string_text(
      docx_in = docx_in,
      docx_out = docx_out,
      magic_string = " ",
      replacement_text = "x"
    ),
    "magic_string cannot be empty"
  )
})
