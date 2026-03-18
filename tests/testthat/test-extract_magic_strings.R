test_that("extract_magic_strings returns unique matches for a single pattern", {
  skip_if_not_installed("officer")

  docx_in <- withr::local_tempfile(fileext = ".docx")

  document <- officer::read_docx()
  document <- officer::body_add_par(document, "Hello <<name>>")
  document <- officer::body_add_par(document, "Repeat <<name>> and <<id>>")
  print(document, target = docx_in)

  result <- extract_magic_strings(
    docx_path = docx_in,
    patterns = "<<[^<>]+>>"
  )

  expect_equal(result, c("<<name>>", "<<id>>"))
})

test_that("extract_magic_strings returns matches for multiple patterns", {
  skip_if_not_installed("officer")

  docx_in <- withr::local_tempfile(fileext = ".docx")

  document <- officer::read_docx()
  document <- officer::body_add_par(document, "Hello <<name>>")
  document <- officer::body_add_par(document, "{rpfy}:table1.docx")
  document <- officer::body_add_par(document, "Again <<name>>")
  print(document, target = docx_in)

  result <- extract_magic_strings(
    docx_path = docx_in,
    patterns = c(
      chevrons = "<<[^<>]+>>",
      rpfy = "\\{rpfy\\}:[^[:space:]]+"
    )
  )

  expect_type(result, "list")
  expect_named(result, c("chevrons", "rpfy"))
  expect_equal(result$chevrons, "<<name>>")
  expect_equal(result$rpfy, "{rpfy}:table1.docx")
})

test_that("extract_magic_strings validates patterns", {
  docx_in <- withr::local_tempfile(fileext = ".docx")
  file.create(docx_in)

  expect_error(
    extract_magic_strings(
      docx_path = docx_in,
      patterns = c("<<[^<>]+>>", " ")
    ),
    "patterns cannot contain empty strings"
  )
})
