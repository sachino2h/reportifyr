test_that("build_report passes docx_table_style to add_tables", {
  calls <- new.env(parent = emptyenv())
  calls$table_style <- NULL

  mockery::stub(build_report, "validate_input_args", function(...) NULL)
  mockery::stub(build_report, "validate_docx", function(...) NULL)
  mockery::stub(build_report, "remove_tables_figures_footnotes", function(...) NULL)
  mockery::stub(build_report, "add_plots", function(...) NULL)
  mockery::stub(build_report, "dir.exists", function(...) TRUE)
  mockery::stub(build_report, "list.files", function(...) character())
  mockery::stub(build_report, "make_doc_dirs", function(...) {
    list(
      doc_clean = "input-clean.docx",
      doc_tables = "input-tabs.docx",
      doc_tabs_figs = "input-tabsfigs.docx",
      doc_draft = "input-draft.docx"
    )
  })
  mockery::stub(build_report, "add_tables", function(..., docx_table_style = NULL) {
    calls$table_style <- docx_table_style
    invisible(NULL)
  })

  expect_silent(
    build_report(
      docx_in = "input.docx",
      docx_out = "output.docx",
      figures_path = "figures",
      tables_path = "tables",
      add_footnotes = FALSE,
      docx_table_style = list(
        font_family = "Times New Roman",
        font_size = 10
      )
    )
  )

  expect_equal(calls$table_style$font_family, "Times New Roman")
  expect_equal(calls$table_style$font_size, 10)
})
