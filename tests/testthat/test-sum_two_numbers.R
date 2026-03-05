test_that("sum_two_numbers returns the sum", {
  expect_equal(sum_two_numbers(2, 3), 5)
  expect_equal(sum_two_numbers(-1, 1), 0)
  expect_equal(sum_two_numbers(0.5, 0.25), 0.75)
})
