# Internal helper functions for XML extraction and input validation

#' Validate docx path
#' @noRd
check_docx_path <- function(path) {
  if (!file.exists(path)) {
    cli::cli_abort("File {.file {path}} does not exist.")
  }
  if (!grepl("\\.docx$", path, ignore.case = TRUE)) {
    cli::cli_abort("File {.file {path}} is not a {.file .docx} file.")
  }
  invisible(path)
}

#' Empty tibble with the comments schema
#' @noRd
empty_comments_tibble <- function() {
  tibble::tibble(
    comment_id = character(),
    author = character(),
    date = character(),
    comment_text = character(),
    commented_text = character(),
    paragraph_context = character()
  )
}

#' Word XML namespace URI
#' @noRd
ns_w <- "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

#' Read and parse DOCX XML parts
#'
#' Replaces officer dependency by directly unzipping the DOCX and reading the
#' XML parts with xml2.
#'
#' @param docx_path Path to a .docx file (already validated by check_docx_path).
#' @return A list with elements:
#'   - `body_xml`: xml_document for word/document.xml
#'   - `comments_xml`: xml_document for word/comments.xml, or NULL if absent
#' @noRd
read_docx_xml <- function(docx_path) {
  tmpdir <- tempfile(pattern = "docxreview_")
  on.exit(unlink(tmpdir, recursive = TRUE), add = TRUE)
  utils::unzip(docx_path, exdir = tmpdir)

  body_path <- file.path(tmpdir, "word", "document.xml")
  if (!file.exists(body_path)) {
    cli::cli_abort("Invalid DOCX: {.file word/document.xml} not found.")
  }

  comments_path <- file.path(tmpdir, "word", "comments.xml")

  list(
    body_xml = xml2::read_xml(body_path),
    comments_xml = if (file.exists(comments_path)) {
      xml2::read_xml(comments_path)
    } else {
      NULL
    }
  )
}

#' Extract full text from a w:p (paragraph) node
#'
#' Collects all w:t and w:delText nodes within the paragraph.
#' @param p_node An xml_node representing a w:p element.
#' @return A single character string with the paragraph text.
#' @noRd
get_paragraph_text <- function(p_node) {
  text_nodes <- xml2::xml_find_all(
    p_node,
    ".//w:t | .//w:delText",
    ns = c(w = ns_w)
  )
  paste(xml2::xml_text(text_nodes), collapse = "")
}

#' Navigate from a change node to its parent paragraph and extract text
#'
#' @param change_node An xml_node (w:ins or w:del element).
#' @return A single character string with the paragraph context.
#' @noRd
get_paragraph_context <- function(change_node) {
  parent_p <- xml2::xml_find_first(
    change_node,
    "ancestor::w:p",
    ns = c(w = ns_w)
  )
  if (inherits(parent_p, "xml_missing")) {
    return(NA_character_)
  }
  get_paragraph_text(parent_p)
}

#' Mark a change within the paragraph context
#'
#' Wraps the changed text with markdown markup:
#' - Deletions: ~~deleted text~~
#' - Insertions: **inserted text**
#'
#' @param paragraph_text Full paragraph text.
#' @param changed_text The text that was inserted or deleted.
#' @param type Either "insertion" or "deletion".
#' @return The paragraph text with the change marked up.
#' @noRd
format_context_with_markup <- function(paragraph_text, changed_text, type) {
  if (is.na(paragraph_text) || is.na(changed_text) || nchar(changed_text) == 0) {
    return(paragraph_text)
  }
  markup <- switch(type,
    deletion = paste0("~~", changed_text, "~~"),
    insertion = paste0("**", changed_text, "**"),
    changed_text
  )
  # sub() replaces only the first occurrence — if changed_text appears multiple

  # times in the paragraph, only the first match is marked (see test-edge_cases.R)
  sub(changed_text, markup, paragraph_text, fixed = TRUE)
}
