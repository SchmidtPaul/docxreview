test_that("extract_tracked_changes() returns expected structure", {
  path <- test_path("fixtures", "test_review.docx")
  result <- extract_tracked_changes(path)

  expect_s3_class(result, "tbl_df")
  expect_named(result, c(
    "change_id", "type", "author", "date",
    "changed_text", "paragraph_context"
  ))
  expect_equal(nrow(result), 2)
})

test_that("extract_tracked_changes() finds deletion", {
  path <- test_path("fixtures", "test_review.docx")
  result <- extract_tracked_changes(path)

  deletions <- result[result$type == "deletion", ]
  expect_equal(nrow(deletions), 1)
  expect_match(deletions$changed_text, "approximately")
  expect_equal(deletions$author, "Max Mustermann")
})

test_that("extract_tracked_changes() finds insertion", {
  path <- test_path("fixtures", "test_review.docx")
  result <- extract_tracked_changes(path)

  insertions <- result[result$type == "insertion", ]
  expect_equal(nrow(insertions), 1)
  expect_match(insertions$changed_text, "statistically")
  expect_equal(insertions$author, "Max Mustermann")
})

test_that("extract_tracked_changes() includes paragraph context with markup", {
  path <- test_path("fixtures", "test_review.docx")
  result <- extract_tracked_changes(path)

  deletions <- result[result$type == "deletion", ]
  expect_match(deletions$paragraph_context, "~~approximately")

  insertions <- result[result$type == "insertion", ]
  expect_match(insertions$paragraph_context, "\\*\\*statistically")
})

test_that("extract_tracked_changes() returns empty tibble for clean doc", {
  path <- test_path("fixtures", "test_no_review.docx")
  result <- extract_tracked_changes(path)

  expect_s3_class(result, "tbl_df")
  expect_equal(nrow(result), 0)
  expect_named(result, c(
    "change_id", "type", "author", "date",
    "changed_text", "paragraph_context"
  ))
})
