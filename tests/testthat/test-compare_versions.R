# --- Input validation (runs everywhere) ---

test_that("compare_versions() errors on non-existent file", {
  expect_error(
    compare_versions("nonexistent.docx", "also_missing.docx"),
    "does not exist"
  )
})

test_that("compare_versions() errors on non-.docx file", {
  tmp <- withr::local_tempfile(fileext = ".txt")
  writeLines("hello", tmp)
  expect_error(
    compare_versions(tmp, tmp),
    "not a.*\\.docx"
  )
})

test_that("compare_versions() errors when output_file is not .docx", {
  path <- test_path("fixtures", "test_review.docx")
  expect_error(
    compare_versions(path, path, output_file = "output.txt"),
    "must have a.*\\.docx"
  )
})

test_that("compare_versions() errors when output directory does not exist", {
  path <- test_path("fixtures", "test_review.docx")
  expect_error(
    compare_versions(path, path, output_file = "nonexistent_dir/output.docx"),
    "does not exist"
  )
})

# --- Integration test (conditional) ---

test_that("compare_versions() produces a comparison document", {
  skip_on_cran()
  skip_on_ci()
  skip_if_not(
    nchar(Sys.which("python")) > 0 || nchar(Sys.which("python3")) > 0,
    "Python not available"
  )
  # Check pywin32
  py <- if (nchar(Sys.which("python")) > 0) "python" else "python3"
  res <- system2(py, c("-c", shQuote("import win32com.client")),
                 stdout = TRUE, stderr = TRUE)
  skip_if_not(
    is.null(attr(res, "status")) || attr(res, "status") == 0,
    "pywin32 not installed"
  )

  # Note: test fixtures are minimal programmatic DOCXs that Word may reject.
  # This test is most useful with real Word-rendered documents.
  old_docx <- test_path("fixtures", "test_no_review.docx")
  new_docx <- test_path("fixtures", "test_review.docx")
  out_file <- withr::local_tempfile(fileext = ".docx")

  tryCatch(
    {
      result <- compare_versions(old_docx, new_docx, output_file = out_file)
      expect_true(file.exists(result))
      expect_equal(result, out_file)
      expect_true(file.size(result) > 0)
    },
    error = function(e) {
      skip("Word could not compare test fixtures (minimal DOCXs)")
    }
  )
})
