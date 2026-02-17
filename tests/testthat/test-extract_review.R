test_that("extract_review() returns markdown string", {
  path <- test_path("fixtures", "test_review.docx")
  result <- capture.output(md <- extract_review(path))

  expect_type(md, "character")
  expect_true(length(md) > 0)
})

test_that("extract_review() markdown contains expected sections", {
  path <- test_path("fixtures", "test_review.docx")
  md <- capture.output(extract_review(path))
  md_text <- paste(md, collapse = "\n")

  expect_match(md_text, "# Review Feedback:")
  expect_match(md_text, "## Comments \\(1\\)")
  expect_match(md_text, "## Tracked Changes \\(2\\)")
})

test_that("extract_review() markdown contains comment details", {
  path <- test_path("fixtures", "test_review.docx")
  md <- capture.output(extract_review(path))
  md_text <- paste(md, collapse = "\n")

  expect_match(md_text, "Max Mustermann")
  expect_match(md_text, "overall survival")
  expect_match(md_text, "secondary endpoints")
})

test_that("extract_review() markdown contains tracked change details", {
  path <- test_path("fixtures", "test_review.docx")
  md <- capture.output(extract_review(path))
  md_text <- paste(md, collapse = "\n")

  expect_match(md_text, "\\[Deletion\\]")
  expect_match(md_text, "\\[Insertion\\]")
  expect_match(md_text, "approximately")
  expect_match(md_text, "statistically")
})

test_that("extract_review() writes to file when output is specified", {
  path <- test_path("fixtures", "test_review.docx")
  out_file <- withr::local_tempfile(fileext = ".md")

  capture.output(extract_review(path, output_file = out_file))

  expect_true(file.exists(out_file))
  content <- readLines(out_file)
  expect_true(length(content) > 0)
  expect_match(content[1], "# Review Feedback:")
})

test_that("extract_review() handles doc without feedback", {
  path <- test_path("fixtures", "test_no_review.docx")
  md <- capture.output(extract_review(path))
  md_text <- paste(md, collapse = "\n")

  expect_match(md_text, "Comments \\(0\\)")
  expect_match(md_text, "Tracked Changes \\(0\\)")
  expect_match(md_text, "No comments found")
  expect_match(md_text, "No tracked changes found")
})

# -- Moved text in markdown ----------------------------------------------------

test_that("extract_review() markdown contains move labels", {
  path <- test_path("fixtures", "test_moved_text.docx")
  md <- capture.output(extract_review(path))
  md_text <- paste(md, collapse = "\n")

  expect_match(md_text, "\\[Moved from here\\]")
  expect_match(md_text, "\\[Moved to here\\]")
  expect_match(md_text, "Moved text")
})

# -- group_by_author -----------------------------------------------------------

test_that("extract_review(group_by_author = TRUE) groups by author", {
  path <- test_path("fixtures", "test_review_rich.docx")
  md <- capture.output(extract_review(path, group_by_author = TRUE))
  md_text <- paste(md, collapse = "\n")

  expect_match(md_text, "## Max Mustermann")
  expect_match(md_text, "## Anna Beispiel")
})

test_that("extract_review(group_by_author = TRUE) shows correct counts", {
  path <- test_path("fixtures", "test_review_rich.docx")
  md <- capture.output(extract_review(path, group_by_author = TRUE))
  md_text <- paste(md, collapse = "\n")

  # Max: 2 comments, 4 changes; Anna: 1 comment, 2 changes
  expect_match(md_text, "### Comments \\(2\\)")
  expect_match(md_text, "### Comments \\(1\\)")
  expect_match(md_text, "### Tracked Changes \\(4\\)")
  expect_match(md_text, "### Tracked Changes \\(2\\)")
})

test_that("extract_review(group_by_author = FALSE) is unchanged (regression)", {
  path <- test_path("fixtures", "test_review_rich.docx")
  md <- capture.output(extract_review(path, group_by_author = FALSE))
  md_text <- paste(md, collapse = "\n")

  # Flat layout uses ## Comments / ## Tracked Changes, not author headings
  expect_match(md_text, "## Comments \\(3\\)")
  expect_match(md_text, "## Tracked Changes \\(6\\)")
  expect_no_match(md_text, "## Max Mustermann")
})

# -- Reply comments in markdown ------------------------------------------------

test_that("extract_review() shows replies indented under parent", {
  path <- test_path("fixtures", "test_comment_replies.docx")
  md <- capture.output(extract_review(path))
  md_text <- paste(md, collapse = "\n")

  expect_match(md_text, "Reply by Anna Beispiel")
  expect_match(md_text, "more detail in section 3")
})

test_that("extract_review() shows only top-level comments as numbered items", {
  path <- test_path("fixtures", "test_comment_replies.docx")
  md <- capture.output(extract_review(path))
  md_text <- paste(md, collapse = "\n")

  # 3 total comments, but only 2 top-level
  expect_match(md_text, "## Comments \\(3\\)")
  expect_match(md_text, "### 1\\.")
  expect_match(md_text, "### 2\\.")
  # No ### 3. because reply is inlined
  expect_no_match(md_text, "### 3\\.")
})
