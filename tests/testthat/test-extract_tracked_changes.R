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

# -- Field instructions fixture ------------------------------------------------

test_that("extract_tracked_changes() filters field instruction nodes", {
  path <- test_path("fixtures", "test_field_instructions.docx")
  result <- extract_tracked_changes(path)

  # 7 raw nodes → 3 after filtering (2 field instr + 1 whitespace + 1 dedup)
  expect_equal(nrow(result), 3)
  expect_true("approximately" %in% result$changed_text)
  expect_true("significant" %in% result$changed_text)
  expect_true("extra" %in% result$changed_text)
})

test_that("extract_tracked_changes() removes whitespace-only changes", {
  path <- test_path("fixtures", "test_field_instructions.docx")
  result <- extract_tracked_changes(path)

  # No row should have whitespace-only changed_text
  expect_true(all(nchar(trimws(result$changed_text)) > 0))
})

test_that("extract_tracked_changes() collapses consecutive duplicates", {
  path <- test_path("fixtures", "test_field_instructions.docx")
  result <- extract_tracked_changes(path)

  # "extra" appears only once despite 2 identical w:del nodes
  expect_equal(sum(result$changed_text == "extra"), 1)
})

test_that("extract_tracked_changes() assigns sequential change_id after filter", {
  path <- test_path("fixtures", "test_field_instructions.docx")
  result <- extract_tracked_changes(path)

  expect_equal(result$change_id, 1:3)
})
