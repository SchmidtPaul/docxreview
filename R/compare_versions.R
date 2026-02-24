#' Compare Two Versions of a Word Document
#'
#' Uses Microsoft Word's built-in document comparison to produce a new DOCX
#' with tracked changes showing all differences between an old and a new
#' version. This is useful after revising a Quarto/R Markdown source based on
#' reviewer feedback — the comparison document lets the reviewer see exactly
#' what changed without manually diffing.
#'
#' @param old_docx Path to the original (older) `.docx` file.
#' @param new_docx Path to the revised (newer) `.docx` file.
#' @param output_file Path for the comparison output `.docx`. If `NULL`
#'   (default), a file named `compared_<old>-vs-<new>.docx` is created in the
#'   directory of `new_docx`.
#'
#' @return The path to the comparison document (invisibly).
#'
#' @details
#' ## Prerequisites
#'
#' This function requires **all three** of the following:
#'
#' - **Microsoft Word** installed (Windows only, uses COM automation)
#' - **Python** available on the system PATH
#' - **pywin32** Python package installed (`pip install pywin32`)
#'
#' The function calls Python via [system2()] to automate Word's
#' `CompareDocuments` feature. It is not available on macOS or Linux.
#'
#' **Note:** This function opens Microsoft Word in the background via COM
#' automation. If Word is already running with unsaved documents, those
#' documents are not affected. Word is only closed if no other documents
#' remain open after the comparison.
#'
#' @export
#' @examples
#' \dontrun{
#' # Auto-generated output path
#' compare_versions("report_v1.docx", "report_v2.docx")
#'
#' # Explicit output path
#' compare_versions("report_v1.docx", "report_v2.docx", "comparison.docx")
#' }
compare_versions <- function(old_docx, new_docx, output_file = NULL) {
  check_docx_path(old_docx)
  check_docx_path(new_docx)

  python_bin <- check_python()
  check_pywin32(python_bin)

  # Resolve to absolute paths (Word COM requires them)
  old_docx <- normalizePath(old_docx, winslash = "/")
  new_docx <- normalizePath(new_docx, winslash = "/")

  if (is.null(output_file)) {
    old_base <- tools::file_path_sans_ext(basename(old_docx))
    new_base <- tools::file_path_sans_ext(basename(new_docx))
    output_file <- file.path(
      dirname(new_docx),
      paste0("compared_", old_base, "-vs-", new_base, ".docx")
    )
  } else {
    if (!grepl("\\.docx$", output_file, ignore.case = TRUE)) {
      cli::cli_abort("{.arg output_file} must have a {.file .docx} extension.")
    }
    out_dir <- dirname(output_file)
    if (!dir.exists(out_dir)) {
      cli::cli_abort("Directory {.file {out_dir}} does not exist.")
    }
    out_base <- basename(output_file)
    output_file <- file.path(normalizePath(out_dir, winslash = "/"), out_base)
  }

  # Build and run temporary Python script
  py_script <- tempfile(fileext = ".py")
  on.exit(unlink(py_script), add = TRUE)

  writeLines(build_compare_script(old_docx, new_docx, output_file), py_script)

  result <- system2(
    python_bin,
    args = shQuote(py_script),
    stdout = TRUE,
    stderr = TRUE,
    timeout = 120  # 2-minute timeout for Word COM
  )

  status <- attr(result, "status")
  if (!is.null(status) && status != 0) {
    # Sanitize non-UTF-8 bytes from Python/Word error output
    output_msg <- iconv(
      paste(result, collapse = "\n"),
      from = "", to = "UTF-8", sub = "?"
    )
    cli::cli_abort(c(
      "Word document comparison failed.",
      "i" = "Python output: {output_msg}"
    ))
  }

  if (!file.exists(output_file)) {
    cli::cli_abort("Comparison file was not created at {.file {output_file}}.")
  }

  cli::cli_alert_success(
    "Comparison saved to {.file {output_file}}"
  )
  invisible(output_file)
}


#' Check that Python is available
#' @return The Python executable path.
#' @noRd
check_python <- function() {
  python <- Sys.which("python")
  if (nchar(python) == 0) {
    python <- Sys.which("python3")
  }
  if (nchar(python) == 0) {
    cli::cli_abort(c(
      "Python not found on the system PATH.",
      "i" = "Install Python from {.url https://www.python.org/downloads/}."
    ))
  }
  unname(python)
}

#' Check that pywin32 is installed
#' @param python_bin Path to the Python executable.
#' @noRd
check_pywin32 <- function(python_bin) {
  result <- system2(
    python_bin,
    args = c("-c", shQuote("import win32com.client")),
    stdout = TRUE,
    stderr = TRUE
  )
  status <- attr(result, "status")
  if (!is.null(status) && status != 0) {
    cli::cli_abort(c(
      "Python package {.pkg pywin32} is not installed.",
      "i" = "Install it with: {.code pip install pywin32}"
    ))
  }
  invisible(TRUE)
}

#' Build the Python script for Word comparison
#' @param old_path,new_path,out_path Absolute paths (forward slashes).
#' @return Character vector of Python script lines.
#' @noRd
build_compare_script <- function(old_path, new_path, out_path) {
  c(
    "import sys",
    "import win32com.client",
    "",
    "try:",
    "    word = win32com.client.Dispatch('Word.Application')",
    "    word.Visible = False",
    paste0("    old_doc = word.Documents.Open(r'",
           gsub("/", "\\\\", old_path), "')"),
    paste0("    new_doc = word.Documents.Open(r'",
           gsub("/", "\\\\", new_path), "')"),
    "    compared = word.Application.CompareDocuments(old_doc, new_doc)",
    paste0("    compared.SaveAs2(r'",
           gsub("/", "\\\\", out_path), "')"),
    "    compared.Close()",
    "    new_doc.Close()",
    "    old_doc.Close()",
    "    # Only quit Word if we launched it (no other documents open)",
    "    if word.Documents.Count == 0:",
    "        word.Quit()",
    "except Exception as e:",
    "    try:",
    "        # Clean up opened documents without quitting Word",
    "        for doc in list(word.Documents):",
    "            doc.Close(0)  # wdDoNotSaveChanges",
    "        if word.Documents.Count == 0:",
    "            word.Quit()",
    "    except Exception:",
    "        pass",
    "    print(str(e), file=sys.stderr)",
    "    sys.exit(1)"
  )
}
