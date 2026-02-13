test_that("extract_comments() returns expected structure from fixture", {
  path <- test_path("fixtures", "test_review.docx")
  result <- extract_comments(path)

  expect_s3_class(result, "tbl_df")
  expect_named(result, c(
    "comment_id", "author", "date", "comment_text",
    "commented_text", "paragraph_context"
  ))
  expect_equal(nrow(result), 1)
})

test_that("extract_comments() extracts correct content", {
  path <- test_path("fixtures", "test_review.docx")
  result <- extract_comments(path)

  expect_equal(result$author, "Max Mustermann")
  expect_match(result$comment_text, "secondary endpoints")
  expect_match(result$commented_text, "overall survival")
})

test_that("extract_comments() includes paragraph context", {
  path <- test_path("fixtures", "test_review.docx")
  result <- extract_comments(path)

  expect_match(result$paragraph_context, "primary endpoint")
  expect_match(result$paragraph_context, "overall survival")
})

test_that("extract_comments() returns empty tibble for doc without comments", {
  path <- test_path("fixtures", "test_no_review.docx")
  result <- extract_comments(path)

  expect_s3_class(result, "tbl_df")
  expect_equal(nrow(result), 0)
  expect_named(result, c(
    "comment_id", "author", "date", "comment_text",
    "commented_text", "paragraph_context"
  ))
})

test_that("extract_comments() errors on non-existent file", {
  expect_error(
    extract_comments("nonexistent.docx"),
    class = "rlang_error"
  )
})

test_that("extract_comments() errors on non-docx file", {
  tmp <- withr::local_tempfile(fileext = ".txt")
  writeLines("not a docx", tmp)
  expect_error(
    extract_comments(tmp),
    class = "rlang_error"
  )
})
