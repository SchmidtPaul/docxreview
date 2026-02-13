#' Extract tracked changes from a Word document
#'
#' Reads a .docx file and extracts all tracked changes (insertions and
#' deletions) including the changed text, author, date, and the full
#' paragraph context.
#'
#' @param docx_path Path to a .docx file.
#' @return A [tibble::tibble] with columns:
#'   - `change_id`: Integer. Sequential ID for each change.
#'   - `type`: Character. Either `"insertion"` or `"deletion"`.
#'   - `author`: Character. Name of the person who made the change.
#'   - `date`: Character. ISO 8601 timestamp of the change.
#'   - `changed_text`: Character. The text that was inserted or deleted.
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

  doc <- officer::read_docx(docx_path)
  body_xml <- officer::docx_body_xml(doc)

  ns <- c(w = ns_w)

  # Combined XPath preserves document order (separate queries + rbind would not)
  change_nodes <- xml2::xml_find_all(body_xml, "//w:ins | //w:del", ns = ns)

  if (length(change_nodes) == 0) {
    cli::cli_inform("No tracked changes found in {.file {docx_path}}.")
    return(empty_changes_tibble())
  }

  result <- extract_change_nodes(change_nodes, ns = ns)
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

  # Determine type from node name (preserves document order from combined XPath)
  types <- ifelse(xml2::xml_name(nodes) == "ins", "insertion", "deletion")
  # Text xpath differs: w:ins contains w:r/w:t, w:del contains w:r/w:delText
  text_xpaths <- ifelse(types == "insertion", ".//w:t", ".//w:delText")

  # Attributes are not namespace-prefixed in OOXML (just "author", not "w:author")
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
