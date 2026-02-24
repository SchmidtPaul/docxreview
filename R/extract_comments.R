#' Extract comments from a Word document
#'
#' Reads a .docx file and extracts all comments including the commented
#' (highlighted) text, the comment text, author, date, and the full paragraph
#' context in which the comment is anchored. Reply threading is resolved
#' from `commentsExtended.xml` when available.
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
#'   - `parent_comment_id`: Character. The ID of the parent comment for replies,
#'     or `NA` for top-level comments.
#'   - `page`: Integer. Page number based on `w:lastRenderedPageBreak` elements,
#'     or `NA` if the document has no rendered page break info.
#'   - `section`: Character. Text of the nearest preceding heading, or `NA` if
#'     the document has no headings or the comment precedes the first heading.
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

  # Resolve reply threading from commentsExtended.xml
  parent_ids <- resolve_comment_threading(
    comments_xml, parts$comments_ext_xml, ids
  )

  # Precompute page/section lookups (once per document)
  all_p <- xml2::xml_find_all(body_xml, "//w:p", ns = ns)
  p_paths <- xml2::xml_path(all_p)
  page_break_indices <- precompute_page_break_indices(all_p, ns)
  heading_style_ids <- build_heading_style_ids(parts$styles_xml, ns)
  heading_indices <- precompute_heading_indices(all_p, heading_style_ids, ns)

  # Per-comment: find anchor paragraph (commentRangeStart → ancestor::w:p)
  pages <- vapply(ids, function(id) {
    range_start <- xml2::xml_find_first(
      body_xml,
      paste0("//w:commentRangeStart[@w:id='", id, "']"),
      ns = ns
    )
    if (inherits(range_start, "xml_missing")) return(NA_integer_)
    p_start <- xml2::xml_find_first(range_start, "ancestor::w:p", ns = ns)
    if (inherits(p_start, "xml_missing")) return(NA_integer_)
    p_idx <- match(xml2::xml_path(p_start), p_paths)
    if (is.na(p_idx)) return(NA_integer_)
    get_page_number(p_start, all_p, p_idx, page_break_indices)
  }, integer(1), USE.NAMES = FALSE)

  sections <- vapply(ids, function(id) {
    range_start <- xml2::xml_find_first(
      body_xml,
      paste0("//w:commentRangeStart[@w:id='", id, "']"),
      ns = ns
    )
    if (inherits(range_start, "xml_missing")) return(NA_character_)
    p_start <- xml2::xml_find_first(range_start, "ancestor::w:p", ns = ns)
    if (inherits(p_start, "xml_missing")) return(NA_character_)
    p_idx <- match(xml2::xml_path(p_start), p_paths)
    if (is.na(p_idx)) return(NA_character_)
    get_nearest_heading(p_idx, all_p, heading_indices)
  }, character(1), USE.NAMES = FALSE)

  tibble::tibble(
    comment_id = ids,
    author = authors,
    date = dates,
    comment_text = comment_texts,
    commented_text = commented_texts,
    paragraph_context = paragraph_contexts,
    parent_comment_id = parent_ids,
    page = pages,
    section = sections
  )
}

#' Resolve comment threading from commentsExtended.xml
#'
#' Maps paraId attributes between comments.xml and commentsExtended.xml
#' to determine parent-child relationships.
#'
#' @param comments_xml xml_document of comments.xml.
#' @param comments_ext_xml xml_document of commentsExtended.xml, or NULL.
#' @param comment_ids Character vector of comment IDs.
#' @return Character vector of parent comment IDs (NA for top-level).
#' @noRd
resolve_comment_threading <- function(comments_xml, comments_ext_xml,
                                       comment_ids) {
  if (is.null(comments_ext_xml)) {
    return(rep(NA_character_, length(comment_ids)))
  }

  ns <- c(w = ns_w)

  # Build map: paraId → comment_id from comments.xml
  # Each w:comment has w:p children; the last w:p carries the paraId
  comment_nodes <- xml2::xml_find_all(comments_xml, "//w:comment", ns = ns)

  para_to_comment <- character()
  for (node in comment_nodes) {
    cid <- xml2::xml_attr(node, "id")
    p_nodes <- xml2::xml_find_all(node, ".//w:p", ns = ns)
    if (length(p_nodes) > 0) {
      last_p <- p_nodes[[length(p_nodes)]]
      # xml2 strips namespace prefix from attributes
      para_id <- xml2::xml_attr(last_p, "paraId")
      if (!is.na(para_id)) {
        para_to_comment[para_id] <- cid
      }
    }
  }

  if (length(para_to_comment) == 0) {
    return(rep(NA_character_, length(comment_ids)))
  }

  # Build map: paraId → paraIdParent from commentsExtended.xml
  # local-name() XPath handles namespace variations (w15, w16cid, etc.)
  ext_nodes <- xml2::xml_find_all(
    comments_ext_xml, "//*[local-name()='commentEx']"
  )

  para_to_parent_para <- character()
  for (node in ext_nodes) {
    pid <- xml2::xml_attr(node, "paraId")
    parent_pid <- xml2::xml_attr(node, "paraIdParent")
    if (!is.na(pid) && !is.na(parent_pid)) {
      para_to_parent_para[pid] <- parent_pid
    }
  }

  # For each comment_id: find its paraId, look up paraIdParent, map back
  vapply(comment_ids, function(cid) {
    my_para <- names(para_to_comment)[para_to_comment == cid]
    if (length(my_para) == 0) return(NA_character_)

    parent_para <- para_to_parent_para[my_para[1]]
    if (is.na(parent_para)) return(NA_character_)

    parent_cid <- para_to_comment[parent_para]
    if (is.na(parent_cid)) NA_character_ else unname(parent_cid)
  }, character(1), USE.NAMES = FALSE)
}

#' Extract the text range selected by a comment
#'
#' Finds all w:r (run) nodes between commentRangeStart and commentRangeEnd,
#' handling comments that span across multiple paragraphs.
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

  p_start <- xml2::xml_find_first(range_start, "ancestor::w:p", ns = ns)
  p_end <- xml2::xml_find_first(range_end, "ancestor::w:p", ns = ns)
  if (inherits(p_start, "xml_missing")) return("")

  # commentRangeEnd may sit directly under w:body (no ancestor w:p)
  if (inherits(p_end, "xml_missing")) {
    p_end <- xml2::xml_find_first(range_end, "preceding-sibling::w:p[1]",
                                   ns = ns)
    if (inherits(p_end, "xml_missing")) return("")
  }

  # Same paragraph: fast path
  if (xml2::xml_path(p_start) == xml2::xml_path(p_end)) {
    return(collect_runs_in_paragraph(p_start, comment_id, ns,
                                     from_start = TRUE, to_end = TRUE))
  }

  # Multi-paragraph: collect from start-p, middle paragraphs, end-p
  all_p <- xml2::xml_find_all(body_xml, "//w:p", ns = ns)
  p_paths <- xml2::xml_path(all_p)
  idx_start <- match(xml2::xml_path(p_start), p_paths)
  idx_end <- match(xml2::xml_path(p_end), p_paths)

  # Guard against match() returning NA (p_start/p_end not found in all_p)
  if (is.na(idx_start) || is.na(idx_end)) return("")

  texts <- character()
  # Start paragraph: runs after commentRangeStart
  texts <- c(texts, collect_runs_in_paragraph(p_start, comment_id, ns,
                                               from_start = TRUE, to_end = FALSE))
  # Middle paragraphs: all text
  if (idx_end - idx_start > 1) {
    for (idx in seq(idx_start + 1, idx_end - 1)) {
      texts <- c(texts, get_paragraph_text(all_p[[idx]]))
    }
  }
  # End paragraph: runs before commentRangeEnd
  texts <- c(texts, collect_runs_in_paragraph(p_end, comment_id, ns,
                                               from_start = FALSE, to_end = TRUE))

  paste(texts[nchar(texts) > 0], collapse = " ")
}

#' Collect run text within a paragraph relative to comment markers
#'
#' @param p_node An xml_node representing a w:p element.
#' @param comment_id The comment ID.
#' @param ns XML namespaces.
#' @param from_start If TRUE, start collecting after commentRangeStart.
#' @param to_end If TRUE, stop collecting at commentRangeEnd.
#' @return Character string of collected text.
#' @noRd
collect_runs_in_paragraph <- function(p_node, comment_id, ns,
                                       from_start, to_end) {
  children <- xml2::xml_children(p_node)
  # If from_start is FALSE, we start collecting immediately (middle/end para)
  collecting <- !from_start
  texts <- character()

  for (child in children) {
    child_name <- xml2::xml_name(child)

    if (child_name == "commentRangeStart" &&
        xml2::xml_attr(child, "id") == comment_id) {
      collecting <- TRUE
      next
    }

    if (child_name == "commentRangeEnd" &&
        xml2::xml_attr(child, "id") == comment_id) {
      break
    }

    if (collecting && child_name == "r") {
      t_nodes <- xml2::xml_find_all(child, ".//w:t", ns = ns)
      texts <- c(texts, xml2::xml_text(t_nodes))
    }
  }

  paste(texts, collapse = "")
}

#' Extract paragraph context for a comment
#'
#' Returns text from all paragraphs that the comment range spans.
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

  range_end <- xml2::xml_find_first(
    body_xml,
    paste0("//w:commentRangeEnd[@w:id='", comment_id, "']"),
    ns = ns
  )

  p_start <- xml2::xml_find_first(range_start, "ancestor::w:p", ns = ns)
  if (inherits(p_start, "xml_missing")) return(NA_character_)

  # If no range_end or same paragraph, return single paragraph

  if (inherits(range_end, "xml_missing")) {
    return(get_paragraph_text(p_start))
  }
  p_end <- xml2::xml_find_first(range_end, "ancestor::w:p", ns = ns)
  # commentRangeEnd may sit directly under w:body (no ancestor w:p)
  if (inherits(p_end, "xml_missing")) {
    p_end <- xml2::xml_find_first(range_end, "preceding-sibling::w:p[1]",
                                   ns = ns)
    if (inherits(p_end, "xml_missing")) return(get_paragraph_text(p_start))
  }
  if (xml2::xml_path(p_start) == xml2::xml_path(p_end)) {
    return(get_paragraph_text(p_start))
  }

  # Multi-paragraph: collect text from all paragraphs in range
  all_p <- xml2::xml_find_all(body_xml, "//w:p", ns = ns)
  p_paths <- xml2::xml_path(all_p)
  idx_start <- match(xml2::xml_path(p_start), p_paths)
  idx_end <- match(xml2::xml_path(p_end), p_paths)

  # Guard against match() returning NA
  if (is.na(idx_start) || is.na(idx_end)) return(get_paragraph_text(p_start))

  texts <- vapply(seq(idx_start, idx_end), function(i) {
    get_paragraph_text(all_p[[i]])
  }, character(1))

  paste(texts[nchar(texts) > 0], collapse = " ")
}
