#' Extract tracked changes from a Word document
#'
#' Reads a .docx file and extracts all tracked changes (insertions, deletions,
#' and moved text) including the changed text, author, date, and the full
#' paragraph context. Field instructions (cross-references, TOC entries,
#' page numbers) and whitespace-only changes are automatically filtered out.
#' Consecutive duplicate changes (same type, author, text, and context) are
#' collapsed into a single entry.
#'
#' @param docx_path Path to a .docx file.
#' @return A [tibble::tibble] with columns:
#'   - `change_id`: Integer. Sequential ID for each change.
#'   - `type`: Character. One of `"insertion"`, `"deletion"`, `"move_from"`,
#'     or `"move_to"`.
#'   - `author`: Character. Name of the person who made the change.
#'   - `date`: Character. ISO 8601 timestamp of the change.
#'   - `changed_text`: Character. The text that was inserted, deleted, or moved.
#'   - `paragraph_context`: Character. Full text of the containing paragraph,
#'     with the change marked up (~~deletion~~ or **insertion**).
#' @export
#' @examples
#' # Extract tracked changes from a reviewed document
#' \dontrun{
#' changes <- extract_tracked_changes("report_reviewed.docx")
#' changes
#' }
extract_tracked_changes <- function(docx_path) {
  check_docx_path(docx_path)
  extract_tracked_changes_impl(read_docx_xml(docx_path))
}

#' Internal implementation for tracked changes extraction
#' @param parts List from read_docx_xml() with body_xml and comments_xml.
#' @return A tibble with tracked changes data.
#' @noRd
extract_tracked_changes_impl <- function(parts) {
  body_xml <- parts$body_xml
  ns <- c(w = ns_w)

  # Combined XPath preserves document order (includes moved text)
  change_nodes <- xml2::xml_find_all(
    body_xml,
    "//w:ins | //w:del | //w:moveFrom | //w:moveTo",
    ns = ns
  )

  if (length(change_nodes) == 0) {
    return(empty_changes_tibble())
  }

  result <- extract_change_nodes(change_nodes, ns = ns)

  # Filter out field instructions and whitespace-only changes
  result <- result[nchar(trimws(result$changed_text)) > 0, , drop = FALSE]

  if (nrow(result) == 0) {
    return(empty_changes_tibble())
  }

  # Collapse consecutive duplicates (same type + author + changed_text + context)
  result <- collapse_consecutive_duplicates(result)

  result$change_id <- seq_len(nrow(result))
  result[, c("change_id", "type", "author", "date",
             "changed_text", "paragraph_context")]
}

#' Extract data from a set of change nodes (w:ins or w:del)
#' @param nodes xml_nodeset of w:ins and/or w:del elements.
#' @param ns Named character vector of XML namespaces.
#' @return A tibble with change data.
#' @noRd
extract_change_nodes <- function(nodes, ns) {
  if (length(nodes) == 0) return(empty_changes_tibble())

  # Determine type from node name
  type_map <- c(
    ins = "insertion", del = "deletion",
    moveFrom = "move_from", moveTo = "move_to"
  )
  types <- unname(type_map[xml2::xml_name(nodes)])
  # Text xpath: ins/moveTo use w:t, del/moveFrom use w:delText
  text_xpaths <- ifelse(types %in% c("insertion", "move_to"),
                         ".//w:t", ".//w:delText")

  authors <- xml2::xml_attr(nodes, "author")
  dates <- xml2::xml_attr(nodes, "date")

  changed_texts <- vapply(seq_along(nodes), function(i) {
    text_nodes <- xml2::xml_find_all(nodes[[i]], text_xpaths[i], ns = ns)
    paste(xml2::xml_text(text_nodes), collapse = "")
  }, character(1))

  paragraph_contexts <- vapply(seq_along(nodes), function(i) {
    ctx <- get_paragraph_context(nodes[[i]])
    format_context_with_markup(ctx, changed_texts[i], types[i])
  }, character(1))

  tibble::tibble(
    change_id = NA_integer_,
    type = types,
    author = authors,
    date = dates,
    changed_text = changed_texts,
    paragraph_context = paragraph_contexts
  )
}

#' Collapse consecutive duplicate rows
#'
#' When Word splits a single change across multiple w:r runs, it produces
#' consecutive rows with identical type, author, changed_text, and
#' paragraph_context. This function collapses them into one row each.
#'
#' @param df A tibble from extract_change_nodes().
#' @return The tibble with consecutive duplicates removed.
#' @noRd
collapse_consecutive_duplicates <- function(df) {
  if (nrow(df) <= 1) return(df)

  keep <- rep(TRUE, nrow(df))
  for (i in seq(2, nrow(df))) {
    if (df$type[i] == df$type[i - 1] &&
        identical(df$author[i], df$author[i - 1]) &&
        df$changed_text[i] == df$changed_text[i - 1] &&
        identical(df$paragraph_context[i], df$paragraph_context[i - 1])) {
      keep[i] <- FALSE
    }
  }
  df[keep, , drop = FALSE]
}

#' Empty tibble with the tracked changes schema
#' @noRd
empty_changes_tibble <- function() {
  tibble::tibble(
    change_id = integer(),
    type = character(),
    author = character(),
    date = character(),
    changed_text = character(),
    paragraph_context = character()
  )
}
