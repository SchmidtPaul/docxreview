test_that("extract_tracked_changes() returns expected structure", {
  path <- test_path("fixtures", "test_review.docx")
  result <- extract_tracked_changes(path)

  expect_s3_class(result, "tbl_df")
  expect_named(result, c(
    "change_id", "type", "author", "date",
    "changed_text", "paragraph_context",
    "page", "section"
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
    "changed_text", "paragraph_context",
    "page", "section"
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

# -- Moved text fixture --------------------------------------------------------

test_that("extract_tracked_changes() finds move_from and move_to", {
  path <- test_path("fixtures", "test_moved_text.docx")
  result <- extract_tracked_changes(path)

  expect_equal(nrow(result), 4)
  expect_true("move_from" %in% result$type)
  expect_true("move_to" %in% result$type)
})

test_that("extract_tracked_changes() extracts correct move_from data", {
  path <- test_path("fixtures", "test_moved_text.docx")
  result <- extract_tracked_changes(path)

  mf <- result[result$type == "move_from", ]
  expect_equal(nrow(mf), 1)
  expect_equal(mf$author, "Anna Beispiel")
  expect_match(mf$changed_text, "The conclusion")
  expect_match(mf$paragraph_context, "~~The conclusion~~")
})

test_that("extract_tracked_changes() extracts correct move_to data", {
  path <- test_path("fixtures", "test_moved_text.docx")
  result <- extract_tracked_changes(path)

  mt <- result[result$type == "move_to", ]
  expect_equal(nrow(mt), 1)
  expect_equal(mt$author, "Anna Beispiel")
  expect_match(mt$changed_text, "The conclusion")
  expect_match(mt$paragraph_context, "\\*\\*The conclusion\\*\\*")
})

test_that("extract_tracked_changes() still finds ins/del alongside moves", {
  path <- test_path("fixtures", "test_moved_text.docx")
  result <- extract_tracked_changes(path)

  expect_equal(nrow(result[result$type == "deletion", ]), 1)
  expect_equal(nrow(result[result$type == "insertion", ]), 1)
})

# -- Page and section extraction -----------------------------------------------

test_that("extract_tracked_changes() returns page number from page break fixture", {
  path <- test_path("fixtures", "test_page_breaks.docx")
  result <- extract_tracked_changes(path)

  # The deletion is in paragraph 5, after 1 page break (paragraph 3) → page 2
  expect_equal(nrow(result), 1)
  expect_equal(result$type, "deletion")
  expect_equal(result$page, 2L)
})

test_that("extract_tracked_changes() returns section from heading fixture", {
  path <- test_path("fixtures", "test_page_breaks.docx")
  result <- extract_tracked_changes(path)

  # The deletion is in paragraph 5, which follows "Methods" heading (paragraph 4)
  expect_equal(result$section, "Methods")
})

test_that("extract_tracked_changes() returns NA page for doc without page breaks", {
  path <- test_path("fixtures", "test_review.docx")
  result <- extract_tracked_changes(path)

  expect_true(all(is.na(result$page)))
})

test_that("extract_tracked_changes() returns NA section for doc without headings", {
  path <- test_path("fixtures", "test_review.docx")
  result <- extract_tracked_changes(path)

  expect_true(all(is.na(result$section)))
})
