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

# -- anonymize -----------------------------------------------------------------

test_that("anonymize = TRUE replaces real names with aliases", {
  path <- test_path("fixtures", "test_review.docx")
  capture.output(md <- extract_review(path, anonymize = TRUE))
  md_text <- paste(md, collapse = "")

  expect_false(grepl("Max Mustermann", md_text))
  expect_true(grepl("Reviewer 1", md_text))
})

test_that("anonymize = TRUE prints mapping to console when no output_file", {
  path <- test_path("fixtures", "test_review.docx")
  out <- capture.output(extract_review(path, anonymize = TRUE))
  out_text <- paste(out, collapse = "\n")

  expect_match(out_text, "Reviewer Mapping")
  expect_match(out_text, "Reviewer 1")
  expect_match(out_text, "Max Mustermann")
})

test_that("anonymize = TRUE writes sidecar file when output_file is set", {
  path <- test_path("fixtures", "test_review.docx")
  out_file <- withr::local_tempfile(fileext = ".md")

  capture.output(extract_review(path, output_file = out_file, anonymize = TRUE))

  map_file <- paste0(tools::file_path_sans_ext(out_file), "-reviewers.md")
  expect_true(file.exists(map_file))

  map_content <- paste(readLines(map_file), collapse = "\n")
  expect_match(map_content, "# Reviewer Mapping")
  expect_match(map_content, "Reviewer 1")
  expect_match(map_content, "Max Mustermann")
  # Alias in column 1, real name in column 2
  expect_match(map_content, "\\| Reviewer 1 \\| Max Mustermann \\|")
})

test_that("anonymize = FALSE (default) leaves names unchanged", {
  path <- test_path("fixtures", "test_review.docx")
  capture.output(md <- extract_review(path, anonymize = FALSE))
  md_text <- paste(md, collapse = "")

  expect_true(grepl("Max Mustermann", md_text))
  expect_false(grepl("Reviewer 1", md_text))
})

test_that("anonymize_authors() replaces name mention in comment_text", {
  comments <- tibble::tibble(
    comment_id = "1", author = "Anna Beispiel",
    date = NA_character_, comment_text = "Hey @Anna Beispiel, see above.",
    commented_text = "some text", paragraph_context = "some text",
    parent_comment_id = NA_character_, page = NA_integer_, section = NA_character_
  )
  changes <- tibble::tibble(
    change_id = character(), author = character(), date = character(),
    type = character(), changed_text = character(),
    paragraph_context = character(), page = integer(), section = character()
  )
  result <- docxreview:::anonymize_authors(comments, changes)

  expect_equal(result$comments$author, "Reviewer 1")
  expect_false(grepl("Anna Beispiel", result$comments$comment_text))
  expect_true(grepl("Reviewer 1", result$comments$comment_text))
})

test_that("anonymize_authors() replaces sub-name split on ' - '", {
  comments <- tibble::tibble(
    comment_id = "1", author = "Paul Schmidt - BioMath GmbH",
    date = NA_character_, comment_text = "@Paul Schmidt please check this.",
    commented_text = "", paragraph_context = "",
    parent_comment_id = NA_character_, page = NA_integer_, section = NA_character_
  )
  changes <- tibble::tibble(
    change_id = character(), author = character(), date = character(),
    type = character(), changed_text = character(),
    paragraph_context = character(), page = integer(), section = character()
  )
  result <- docxreview:::anonymize_authors(comments, changes)

  expect_false(grepl("Paul Schmidt", result$comments$comment_text))
  expect_true(grepl("Reviewer 1", result$comments$comment_text))
})

test_that("name_variants() splits on ' - ' and keeps 2+-word parts", {
  variants <- docxreview:::name_variants("Paul Schmidt - BioMath GmbH")
  expect_true("Paul Schmidt - BioMath GmbH" %in% variants)
  expect_true("Paul Schmidt" %in% variants)
  expect_true("BioMath GmbH" %in% variants)

  # Single-word parts should not be included
  variants_simple <- docxreview:::name_variants("Paul Schmidt")
  expect_equal(variants_simple, "Paul Schmidt")
})
