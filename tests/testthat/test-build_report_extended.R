test_that("build_report_extended returns wiring success message", {
  docx_in <- tempfile(fileext = ".docx")
  yaml_in <- tempfile(fileext = ".yaml")
  file.create(docx_in)
  writeLines("inline: {}", con = yaml_in)

  out <- build_report_extended(docx_in = docx_in, yaml_in = yaml_in)

  expect_type(out, "character")
  expect_length(out, 1)
  expect_match(out, "build_report_extended initialized successfully")
  expect_match(out, basename(docx_in), fixed = TRUE)
  expect_match(out, basename(yaml_in), fixed = TRUE)
})

test_that("build_report_extended validates missing files", {
  expect_error(
    build_report_extended(docx_in = "missing.docx", yaml_in = "missing.yaml"),
    "docx_in does not exist",
    fixed = TRUE
  )
})
