# -- Rich fixture: extract_comments() ----------------------------------------

test_that("extract_comments() finds 3 comments from 2 authors (rich fixture)", {
  path <- test_path("fixtures", "test_review_rich.docx")
  result <- extract_comments(path)

  expect_equal(nrow(result), 3)
  expect_equal(sort(unique(result$author)), c("Anna Beispiel", "Max Mustermann"))
})

test_that("extract_comments() handles Unicode in comment text", {
  path <- test_path("fixtures", "test_review_rich.docx")
  result <- extract_comments(path)

  anna <- result[result$author == "Anna Beispiel", ]
  expect_match(anna$comment_text, "pr\u00fcfen")
})

test_that("extract_comments() paragraph context includes deleted text", {
  path <- test_path("fixtures", "test_review_rich.docx")
  result <- extract_comments(path)

  # Comment 10 is on "200 patients" in paragraph that also has deletion "approximately"
  row10 <- result[result$comment_id == "10", ]
  expect_match(row10$paragraph_context, "approximately")
  expect_match(row10$paragraph_context, "200 patients")
})

# -- Rich fixture: extract_tracked_changes() ---------------------------------

test_that("extract_tracked_changes() finds 6 changes (4 del, 2 ins) (rich fixture)", {
  path <- test_path("fixtures", "test_review_rich.docx")
  result <- extract_tracked_changes(path)

  expect_equal(nrow(result), 6)
  expect_equal(sum(result$type == "deletion"), 4)
  expect_equal(sum(result$type == "insertion"), 2)
})

test_that("extract_tracked_changes() handles consecutive changes in same paragraph", {
  path <- test_path("fixtures", "test_review_rich.docx")
  result <- extract_tracked_changes(path)

  # P4: deletion "old " and insertion "new " in same paragraph
  del_old <- result[result$changed_text == "old ", ]
  ins_new <- result[result$changed_text == "new ", ]

  expect_equal(nrow(del_old), 1)
  expect_equal(nrow(ins_new), 1)
  expect_equal(del_old$type, "deletion")
  expect_equal(ins_new$type, "insertion")
})

test_that("extract_tracked_changes() returns NA date when attribute is missing", {
  path <- test_path("fixtures", "test_review_rich.docx")
  result <- extract_tracked_changes(path)

  # P5: deletion "preliminary " has no w:date attribute
  prelim <- result[result$changed_text == "preliminary ", ]
  expect_true(is.na(prelim$date))
})

test_that("sub() marks only first occurrence of duplicated text", {
  path <- test_path("fixtures", "test_review_rich.docx")
  result <- extract_tracked_changes(path)

  # P6: "data" deleted but "data" also appears in "data quality"
  data_change <- result[result$changed_text == "data", ]
  expect_match(data_change$paragraph_context, "~~data~~")
  # Second "data" (in "data quality") should NOT be marked
  expect_match(data_change$paragraph_context, "data quality")
  # Only one set of strikethrough markers
  expect_equal(
    lengths(regmatches(
      data_change$paragraph_context,
      gregexpr("~~data~~", data_change$paragraph_context)
    )),
    1
  )
})

# -- Rich fixture: extract_review() -----------------------------------------

test_that("extract_review() markdown has correct counts (rich fixture)", {
  path <- test_path("fixtures", "test_review_rich.docx")
  md <- extract_review(path, output = tempfile(fileext = ".md"))

  expect_match(md, "Comments \\(3\\)", all = FALSE)
  expect_match(md, "Tracked Changes \\(6\\)", all = FALSE)
})

test_that("extract_review() markdown shows 'unknown date' for missing date", {
  path <- test_path("fixtures", "test_review_rich.docx")
  md <- extract_review(path, output = tempfile(fileext = ".md"))

  expect_match(md, "unknown date", all = FALSE)
})

# -- format_date() edge cases ------------------------------------------------

test_that("format_date() handles NA variants", {
  expect_equal(docxreview:::format_date(NA), "unknown date")
  expect_equal(docxreview:::format_date(NA_character_), "unknown date")
  expect_equal(docxreview:::format_date(""), "unknown date")
})

test_that("format_date() handles NULL", {
  expect_equal(docxreview:::format_date(NULL), "unknown date")
})

test_that("format_date() parses ISO 8601 to date-only", {
  expect_equal(
    docxreview:::format_date("2024-01-15T10:30:00Z"),
    "2024-01-15"
  )
})

test_that("format_date() passes through unparseable strings", {
  expect_equal(docxreview:::format_date("not-a-date"), "not-a-date")
})

test_that("format_date() passes through plain date unchanged", {
  expect_equal(docxreview:::format_date("2024-01-15"), "2024-01-15")
})

# -- format_context_with_markup() edge cases ---------------------------------

test_that("format_context_with_markup() returns NA paragraph as-is", {
  expect_true(is.na(
    docxreview:::format_context_with_markup(NA_character_, "text", "deletion")
  ))
})

test_that("format_context_with_markup() returns paragraph when changed_text is NA", {
  expect_equal(
    docxreview:::format_context_with_markup("some paragraph", NA_character_, "deletion"),
    "some paragraph"
  )
})

test_that("format_context_with_markup() returns paragraph when changed_text is empty", {
  expect_equal(
    docxreview:::format_context_with_markup("some paragraph", "", "deletion"),
    "some paragraph"
  )
})

test_that("format_context_with_markup() wraps deletion in ~~", {
  expect_equal(
    docxreview:::format_context_with_markup("remove old word", "old", "deletion"),
    "remove ~~old~~ word"
  )
})

test_that("format_context_with_markup() wraps insertion in **", {
  expect_equal(
    docxreview:::format_context_with_markup("add new word", "new", "insertion"),
    "add **new** word"
  )
})

test_that("format_context_with_markup() marks only first match for duplicates", {
  result <- docxreview:::format_context_with_markup(
    "data showed data quality", "data", "deletion"
  )
  expect_equal(result, "~~data~~ showed data quality")
})

# -- Output path validation --------------------------------------------------

test_that("extract_review() errors when output directory does not exist", {
  path <- test_path("fixtures", "test_review_rich.docx")
  bad_output <- file.path(tempdir(), "nonexistent_subdir_12345", "out.md")

  expect_error(
    extract_review(path, output = bad_output),
    "does not exist",
    class = "rlang_error"
  )
})

test_that("extract_review() writes file to valid temp path", {
  path <- test_path("fixtures", "test_review_rich.docx")
  out <- tempfile(fileext = ".md")

  extract_review(path, output = out)
  expect_true(file.exists(out))
})
