#' Extract comments from a Word document
#'
#' Reads a .docx file and extracts all comments including the commented
#' (highlighted) text, the comment text, author, date, and the full paragraph
#' context in which the comment is anchored.
#'
#' @param docx_path Path to a .docx file.
#' @return A [tibble::tibble] with columns:
#'   - `comment_id`: Character. The comment ID from the document.
#'   - `author`: Character. Name of the comment author.
#'   - `date`: Character. ISO 8601 timestamp of the comment.
#'   - `comment_text`: Character. The comment body text.
#'   - `commented_text`: Character. The text range that was highlighted/selected,
#'     or an empty string if no text range was selected.
#'   - `paragraph_context`: Character. Full text of the paragraph(s) containing
#'     the commented range.
#' @export
#' @examples
#' # Extract comments from a reviewed document
#' \dontrun{
#' comments <- extract_comments("report_reviewed.docx")
#' comments
#' }
extract_comments <- function(docx_path) {
  check_docx_path(docx_path)

  doc <- officer::read_docx(docx_path)
  raw <- officer::docx_comments(doc)

  if (nrow(raw) == 0) {
    cli::cli_inform("No comments found in {.file {docx_path}}.")
    return(empty_comments_tibble())
  }

  doc_xml <- officer::docx_body_xml(doc)

  # For each comment, find the paragraph(s) containing the comment range
  paragraph_contexts <- vapply(raw$comment_id, function(id) {
    # Find all commentRangeStart nodes with this ID
    range_starts <- xml2::xml_find_all(
      doc_xml,
      paste0("//w:commentRangeStart[@w:id='", id, "']"),
      ns = c(w = ns_w)
    )
    if (length(range_starts) == 0) return(NA_character_)

    # Get parent paragraph(s) — usually one, but could span multiple
    parent_ps <- xml2::xml_find_all(
      range_starts[[1]],
      "ancestor::w:p",
      ns = c(w = ns_w)
    )
    if (length(parent_ps) == 0) return(NA_character_)

    paste(vapply(parent_ps, get_paragraph_text, character(1)), collapse = " ")
  }, character(1))

  # Collapse list columns from officer::docx_comments() into character
  comment_text <- vapply(raw$text, function(x) {
    paste(x, collapse = " ")
  }, character(1))
  commented_text <- vapply(raw$commented_text, function(x) {
    paste(x, collapse = "")
  }, character(1))

  tibble::tibble(
    comment_id = raw$comment_id,
    author = raw$author,
    date = raw$date,
    comment_text = comment_text,
    commented_text = commented_text,
    paragraph_context = paragraph_contexts
  )
}
