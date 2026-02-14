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
