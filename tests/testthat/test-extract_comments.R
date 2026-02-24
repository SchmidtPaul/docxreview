test_that("extract_comments() returns expected structure from fixture", {
  path <- test_path("fixtures", "test_review.docx")
  result <- extract_comments(path)

  expect_s3_class(result, "tbl_df")
  expect_named(result, c(
    "comment_id", "author", "date", "comment_text",
    "commented_text", "paragraph_context", "parent_comment_id",
    "page", "section"
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
    "commented_text", "paragraph_context", "parent_comment_id",
    "page", "section"
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

# -- Multi-paragraph comments ---------------------------------------------------

test_that("extract_comments() handles single-paragraph comment in multi-par fixture", {
  path <- test_path("fixtures", "test_multipar_comment.docx")
  result <- extract_comments(path)

  # Comment 0 is a single-paragraph comment (regression)
  row <- result[result$comment_id == "0", ]
  expect_equal(nrow(row), 1)
  expect_match(row$commented_text, "study")
  expect_match(row$paragraph_context, "enrolled patients")
})

test_that("extract_comments() spans 2 paragraphs", {
  path <- test_path("fixtures", "test_multipar_comment.docx")
  result <- extract_comments(path)

  row <- result[result$comment_id == "1", ]
  expect_equal(nrow(row), 1)
  # commented_text must contain text from both paragraphs

  expect_match(row$commented_text, "patients from multiple sites")
  expect_match(row$commented_text, "Enrollment was completed")
  # paragraph_context must span both paragraphs
  expect_match(row$paragraph_context, "enrolled patients")
  expect_match(row$paragraph_context, "December 2023")
})

test_that("extract_comments() spans 3 paragraphs", {
  path <- test_path("fixtures", "test_multipar_comment.docx")
  result <- extract_comments(path)

  row <- result[result$comment_id == "2", ]
  expect_equal(nrow(row), 1)
  # commented_text must contain text from all 3 paragraphs
  expect_match(row$commented_text, "all randomized subjects")
  expect_match(row$commented_text, "Sensitivity analyses")
  expect_match(row$commented_text, "findings were consistent")
  # paragraph_context must span all 3 paragraphs
  expect_match(row$paragraph_context, "primary analysis")
  expect_match(row$paragraph_context, "Sensitivity analyses")
  expect_match(row$paragraph_context, "across subgroups")
})

test_that("extract_comments() returns correct count for multi-par fixture", {
  path <- test_path("fixtures", "test_multipar_comment.docx")
  result <- extract_comments(path)
  expect_equal(nrow(result), 3)
})

# -- Reply comments ------------------------------------------------------------

test_that("extract_comments() includes parent_comment_id column", {
  path <- test_path("fixtures", "test_review.docx")
  result <- extract_comments(path)
  expect_true("parent_comment_id" %in% names(result))
})

test_that("extract_comments() returns NA parent for DOCX without commentsExtended", {
  path <- test_path("fixtures", "test_review.docx")
  result <- extract_comments(path)
  expect_true(all(is.na(result$parent_comment_id)))
})

test_that("extract_comments() resolves reply threading", {
  path <- test_path("fixtures", "test_comment_replies.docx")
  result <- extract_comments(path)

  expect_equal(nrow(result), 3)
  # Comment 0: top-level
  expect_true(is.na(result$parent_comment_id[result$comment_id == "0"]))
  # Comment 1: reply to comment 0
  expect_equal(result$parent_comment_id[result$comment_id == "1"], "0")
  # Comment 2: top-level
  expect_true(is.na(result$parent_comment_id[result$comment_id == "2"]))
})

test_that("extract_comments() extracts correct reply content", {
  path <- test_path("fixtures", "test_comment_replies.docx")
  result <- extract_comments(path)

  reply <- result[result$comment_id == "1", ]
  expect_equal(reply$author, "Anna Beispiel")
  expect_match(reply$comment_text, "more detail in section 3")
})

test_that("empty_comments_tibble() includes parent_comment_id", {
  tb <- docxreview:::empty_comments_tibble()
  expect_true("parent_comment_id" %in% names(tb))
  expect_equal(nrow(tb), 0)
})

# -- commentRangeEnd directly under w:body (no ancestor::w:p) ------------------

test_that("extract_comments() handles commentRangeEnd under w:body", {

  path <- test_path("fixtures", "test_comment_range_end_body.docx")
  result <- extract_comments(path)

  expect_equal(nrow(result), 1)
  expect_equal(result$author, "Max Mustermann")
  expect_match(result$comment_text, "Check this section")
})

test_that("commentRangeEnd under w:body: commented_text is extracted", {
  path <- test_path("fixtures", "test_comment_range_end_body.docx")
  result <- extract_comments(path)

  # Should still extract the text from the paragraph with commentRangeStart
  expect_match(result$commented_text, "overall survival")
})

test_that("commentRangeEnd under w:body: paragraph_context is not NA", {
  path <- test_path("fixtures", "test_comment_range_end_body.docx")
  result <- extract_comments(path)

  expect_false(is.na(result$paragraph_context))
  expect_match(result$paragraph_context, "overall survival")
})

# -- page/section extraction ---------------------------------------------------

test_that("extract_comments() returns page number from page break fixture", {
  path <- test_path("fixtures", "test_page_breaks.docx")
  result <- extract_comments(path)

  # Comment is on paragraph 2 (before first page break) → page 1
  expect_equal(result$page, 1L)
})

test_that("extract_comments() returns section from heading fixture", {
  path <- test_path("fixtures", "test_page_breaks.docx")
  result <- extract_comments(path)

  # Comment is on paragraph 2, which follows the "Introduction" heading
  expect_equal(result$section, "Introduction")
})
