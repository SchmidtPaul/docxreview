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
    paragraph_context = character(),
    parent_comment_id = character(),
    page = integer(),
    section = character()
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
  comments_ext_path <- file.path(tmpdir, "word", "commentsExtended.xml")

  styles_path <- file.path(tmpdir, "word", "styles.xml")

  list(
    body_xml = xml2::read_xml(body_path),
    comments_xml = if (file.exists(comments_path)) {
      xml2::read_xml(comments_path)
    } else {
      NULL
    },
    comments_ext_xml = if (file.exists(comments_ext_path)) {
      xml2::read_xml(comments_ext_path)
    } else {
      NULL
    },
    styles_xml = if (file.exists(styles_path)) {
      xml2::read_xml(styles_path)
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
    move_from = paste0("~~", changed_text, "~~"),
    move_to = paste0("**", changed_text, "**"),
    changed_text
  )
  # sub() replaces only the first occurrence — if changed_text appears multiple

  # times in the paragraph, only the first match is marked (see test-edge_cases.R)
  sub(changed_text, markup, paragraph_text, fixed = TRUE)
}

#' Count the page number of a node based on w:lastRenderedPageBreak
#'
#' @param p_node An xml_node representing the paragraph containing the item.
#' @param all_p xml_nodeset of all w:p in the body (precomputed).
#' @param p_idx Integer index of p_node in all_p.
#' @param page_break_p_indices Integer vector of paragraph indices that contain
#'   a page break (precomputed). NULL if no breaks exist.
#' @return An integer page number, or NA_integer_ if no breaks in the document.
#' @noRd
get_page_number <- function(p_node, all_p, p_idx, page_break_p_indices) {
  if (is.null(page_break_p_indices)) return(NA_integer_)
  # Page = 1 + number of page breaks in paragraphs before (and including) this one
  1L + sum(page_break_p_indices <= p_idx)
}

#' Build a mapping of heading style IDs from styles.xml
#'
#' @param styles_xml xml_document of styles.xml, or NULL.
#' @param ns Named character vector of XML namespaces.
#' @return Character vector of style IDs that are headings, or NULL.
#' @noRd
build_heading_style_ids <- function(styles_xml, ns) {
  if (is.null(styles_xml)) return(NULL)

  style_nodes <- xml2::xml_find_all(
    styles_xml,
    "//w:style[@w:type='paragraph']",
    ns = ns
  )
  if (length(style_nodes) == 0) return(NULL)

  ids <- character()
  for (node in style_nodes) {
    outline <- xml2::xml_find_first(node, ".//w:pPr/w:outlineLvl", ns = ns)
    if (!inherits(outline, "xml_missing")) {
      lvl <- as.integer(xml2::xml_attr(outline, "val"))
      if (!is.na(lvl) && lvl <= 8L) {
        ids <- c(ids, xml2::xml_attr(node, "styleId"))
      }
    }
  }
  if (length(ids) == 0) NULL else ids
}

#' Find the nearest preceding heading for a paragraph
#'
#' @param p_idx Integer index of the target paragraph in all_p.
#' @param all_p xml_nodeset of all w:p in the body (precomputed).
#' @param heading_indices Named integer vector where values are indices of
#'   heading paragraphs and names are their text (precomputed). NULL if none.
#' @return The heading text, or NA_character_.
#' @noRd
get_nearest_heading <- function(p_idx, all_p, heading_indices) {
  if (is.null(heading_indices) || length(heading_indices) == 0) {
    return(NA_character_)
  }
  # Find headings that come before (or at) p_idx
  preceding <- heading_indices[heading_indices <= p_idx]
  if (length(preceding) == 0) return(NA_character_)
  # Return the text of the closest heading (last in the preceding list)
  names(preceding)[length(preceding)]
}

#' Precompute heading paragraph indices and their text
#'
#' @param all_p xml_nodeset of all w:p in the body.
#' @param heading_style_ids Character vector of heading style IDs, or NULL.
#' @param ns Named character vector of XML namespaces.
#' @return Named integer vector (names = heading text, values = indices), or NULL.
#' @noRd
precompute_heading_indices <- function(all_p, heading_style_ids, ns) {
  if (is.null(heading_style_ids) || length(all_p) == 0) return(NULL)

  # Build XPath to match any heading style
  style_conditions <- paste0("@w:val='", heading_style_ids, "'", collapse = " or ")
  heading_xpath <- paste0(".//w:pPr/w:pStyle[", style_conditions, "]")

  indices <- integer()
  texts <- character()

  for (i in seq_along(all_p)) {
    style_match <- xml2::xml_find_first(all_p[[i]], heading_xpath, ns = ns)
    if (!inherits(style_match, "xml_missing")) {
      txt <- get_paragraph_text(all_p[[i]])
      if (nchar(txt) > 0) {
        indices <- c(indices, i)
        texts <- c(texts, txt)
      }
    }
  }

  if (length(indices) == 0) return(NULL)
  stats::setNames(indices, texts)
}

#' Precompute paragraph indices that contain a page break
#'
#' @param all_p xml_nodeset of all w:p in the body.
#' @param ns Named character vector of XML namespaces.
#' @return Integer vector of paragraph indices with page breaks, or NULL.
#' @noRd
precompute_page_break_indices <- function(all_p, ns) {
  if (length(all_p) == 0) return(NULL)

  indices <- integer()
  for (i in seq_along(all_p)) {
    breaks <- xml2::xml_find_first(
      all_p[[i]], ".//w:lastRenderedPageBreak", ns = ns
    )
    if (!inherits(breaks, "xml_missing")) {
      indices <- c(indices, i)
    }
  }

  if (length(indices) == 0) NULL else indices
}
