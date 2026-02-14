# Internal helper functions for XML extraction

#' Word XML namespace URI
#' @noRd
ns_w <- "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

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
