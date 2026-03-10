test_that("add_tables routes .docx tables through insert_formatted_table", {
  calls <- new.env(parent = emptyenv())
  calls$insert <- 0L
  calls$process <- 0L
  calls$include_footnote <- NULL
  calls$alt_text_input <- NULL
  calls$alt_text_output <- NULL

  mockery::stub(add_tables, "validate_input_args", function(...) NULL)
  mockery::stub(add_tables, "validate_docx", function(...) NULL)
  mockery::stub(add_tables, "keep_caption_next", function(...) NULL)
  mockery::stub(add_tables, "officer::read_docx", function(...) structure(list(), class = "rdocx"))
  mockery::stub(add_tables, "officer::docx_summary", function(...) {
    data.frame(text = "{rpfy}:table1.docx", stringsAsFactors = FALSE)
  })
  mockery::stub(add_tables, "file.exists", function(...) TRUE)
  mockery::stub(add_tables, "process_table_file", function(...) {
    calls$process <- calls$process + 1L
    structure(list(), class = "rdocx")
  })
  mockery::stub(add_tables, "insert_formatted_table", function(
    source_docx_path,
    template_path,
    output_path,
    placeholder_text,
    include_footnote = FALSE,
    ...
  ) {
    calls$insert <- calls$insert + 1L
    calls$include_footnote <- include_footnote
    invisible(NULL)
  })
  mockery::stub(add_tables, "print", function(...) invisible(NULL))
  mockery::stub(add_tables, "add_tables_alt_text", function(docx_in, docx_out, ...) {
    calls$alt_text_input <- docx_in
    calls$alt_text_output <- docx_out
    invisible(NULL)
  })
  mockery::stub(add_tables, "unlink", function(...) TRUE)

  expect_silent(
    add_tables(
      docx_in = "input.docx",
      docx_out = "output.docx",
      tables_path = "tables",
      docx_footnote = TRUE
    )
  )

  expect_equal(calls$insert, 1L)
  expect_equal(calls$process, 0L)
  expect_true(calls$include_footnote)
  expect_true(grepl("-intdocxtab-1\\.docx$", calls$alt_text_input))
  expect_equal(calls$alt_text_output, "output.docx")
})
