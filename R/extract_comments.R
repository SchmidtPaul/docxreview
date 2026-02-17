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
  extract_comments_impl(read_docx_xml(docx_path))
}

#' Internal implementation for comment extraction
#' @param parts List from read_docx_xml() with body_xml and comments_xml.
#' @return A tibble with comment data.
#' @noRd
extract_comments_impl <- function(parts) {
  comments_xml <- parts$comments_xml
  body_xml <- parts$body_xml
  ns <- c(w = ns_w)

  if (is.null(comments_xml)) {
    return(empty_comments_tibble())
  }

  comment_nodes <- xml2::xml_find_all(comments_xml, "//w:comment", ns = ns)

  if (length(comment_nodes) == 0) {
    return(empty_comments_tibble())
  }

  # Extract comment metadata + text from comments.xml
  ids <- xml2::xml_attr(comment_nodes, "id")
  authors <- xml2::xml_attr(comment_nodes, "author")
  dates <- xml2::xml_attr(comment_nodes, "date")

  comment_texts <- vapply(comment_nodes, function(node) {
    text_nodes <- xml2::xml_find_all(node, ".//w:t", ns = ns)
    paste(xml2::xml_text(text_nodes), collapse = " ")
  }, character(1))

  # Extract commented_text and paragraph_context from body XML
  commented_texts <- vapply(ids, function(id) {
    extract_commented_text(body_xml, id, ns)
  }, character(1))

  paragraph_contexts <- vapply(ids, function(id) {
    extract_comment_paragraph_context(body_xml, id, ns)
  }, character(1))

  tibble::tibble(
    comment_id = ids,
    author = authors,
    date = dates,
    comment_text = comment_texts,
    commented_text = commented_texts,
    paragraph_context = paragraph_contexts
  )
}

#' Extract the text range selected by a comment
#'
#' Finds all w:r (run) nodes between commentRangeStart and commentRangeEnd
#' within the same paragraph and collects their text.
#'
#' @param body_xml xml_document of the document body.
#' @param comment_id Character. The comment ID to look up.
#' @param ns Named character vector of XML namespaces.
#' @return Character string of the commented text, or "" if not found.
#' @noRd
extract_commented_text <- function(body_xml, comment_id, ns) {
  range_start <- xml2::xml_find_first(
    body_xml,
    paste0("//w:commentRangeStart[@w:id='", comment_id, "']"),
    ns = ns
  )
  if (inherits(range_start, "xml_missing")) return("")

  range_end <- xml2::xml_find_first(
    body_xml,
    paste0("//w:commentRangeEnd[@w:id='", comment_id, "']"),
    ns = ns
  )
  if (inherits(range_end, "xml_missing")) return("")

  # Get the parent paragraph of the range start
  parent_p <- xml2::xml_find_first(range_start, "ancestor::w:p", ns = ns)
  if (inherits(parent_p, "xml_missing")) return("")

  # Strategy: walk sibling nodes between commentRangeStart and commentRangeEnd
  # within the paragraph, collecting text from w:r nodes
  children <- xml2::xml_children(parent_p)
  collecting <- FALSE
  texts <- character()

  for (child in children) {
    child_name <- xml2::xml_name(child)

    # Check if this is our commentRangeStart
    if (child_name == "commentRangeStart" &&
        xml2::xml_attr(child, "id") == comment_id) {
      collecting <- TRUE
      next
    }

    # Check if this is our commentRangeEnd
    if (child_name == "commentRangeEnd" &&
        xml2::xml_attr(child, "id") == comment_id) {
      break
    }

    # Collect text from runs between start and end
    if (collecting && child_name == "r") {
      t_nodes <- xml2::xml_find_all(child, ".//w:t", ns = ns)
      texts <- c(texts, xml2::xml_text(t_nodes))
    }
  }

  paste(texts, collapse = "")
}

#' Extract paragraph context for a comment
#'
#' @param body_xml xml_document of the document body.
#' @param comment_id Character. The comment ID to look up.
#' @param ns Named character vector of XML namespaces.
#' @return Character string of the full paragraph text, or NA_character_.
#' @noRd
extract_comment_paragraph_context <- function(body_xml, comment_id, ns) {
  range_start <- xml2::xml_find_first(
    body_xml,
    paste0("//w:commentRangeStart[@w:id='", comment_id, "']"),
    ns = ns
  )
  if (inherits(range_start, "xml_missing")) return(NA_character_)

  parent_p <- xml2::xml_find_first(range_start, "ancestor::w:p", ns = ns)
  if (inherits(parent_p, "xml_missing")) return(NA_character_)

  get_paragraph_text(parent_p)
}
