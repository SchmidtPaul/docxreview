#' Extract all review feedback from a Word document
#'
#' Extracts both comments and tracked changes from a reviewed .docx file and
#' returns a structured Markdown summary. This is the main entry point for
#' processing reviewer feedback.
#'
#' @param docx_path Path to a .docx file.
#' @param output_file Optional file path to write the Markdown output to. If
#'   `NULL` (default), the Markdown string is returned invisibly and printed to
#'   the console.
#' @return Invisibly returns the Markdown string. If `output_file` is specified,
#'   the Markdown is also written to that file.
#' @export
#' @examples
#' # Print review feedback to console
#' \dontrun{
#' extract_review("report_reviewed.docx")
#'
#' # Save to file
#' extract_review("report_reviewed.docx", output_file = "feedback.md")
#' }
extract_review <- function(docx_path, output_file = NULL) {
  check_docx_path(docx_path)

  comments <- extract_comments(docx_path)
  changes <- extract_tracked_changes(docx_path)

  md <- format_review_markdown(comments, changes, basename(docx_path))

  if (!is.null(output_file)) {
    output_dir <- dirname(output_file)
    if (!dir.exists(output_dir)) {
      cli::cli_abort("Directory {.file {output_dir}} does not exist.")
    }
    writeLines(md, output_file)
    cli::cli_inform("Review feedback written to {.file {output_file}}.")
  } else {
    cat(md, sep = "\n")
  }

  invisible(md)
}

#' Format review data as Markdown
#' @param comments Tibble from extract_comments().
#' @param changes Tibble from extract_tracked_changes().
#' @param filename The document filename for the header.
#' @return Character vector of Markdown lines.
#' @noRd
format_review_markdown <- function(comments, changes, filename) {
  lines <- character()
  lines <- c(lines, paste0("# Review Feedback: ", filename), "")

  # Comments section
  lines <- c(lines, paste0("## Comments (", nrow(comments), ")"), "")

  if (nrow(comments) > 0) {
    for (i in seq_len(nrow(comments))) {
      row <- comments[i, ]
      date_fmt <- format_date(row$date)
      lines <- c(lines,
        paste0("### ", i, ". ", row$author, " (", date_fmt, ")"),
        paste0('> **Commented text:** "', row$commented_text, '"'),
        paste0('> **Paragraph:** "', row$paragraph_context, '"'),
        "",
        paste0("**Comment:** ", row$comment_text),
        "",
        "---",
        ""
      )
    }
  } else {
    lines <- c(lines, "*No comments found.*", "")
  }

  # Tracked changes section
  lines <- c(lines, paste0("## Tracked Changes (", nrow(changes), ")"), "")

  if (nrow(changes) > 0) {
    for (i in seq_len(nrow(changes))) {
      row <- changes[i, ]
      date_fmt <- format_date(row$date)
      type_label <- if (row$type == "deletion") "Deletion" else "Insertion"
      text_label <- if (row$type == "deletion") "Deleted" else "Inserted"

      lines <- c(lines,
        paste0("### ", i, ". [", type_label, "] ", row$author,
               " (", date_fmt, ")"),
        paste0('> **', text_label, ':** "', row$changed_text, '"'),
        paste0('> **Paragraph:** "', row$paragraph_context, '"'),
        "",
        "---",
        ""
      )
    }
  } else {
    lines <- c(lines, "*No tracked changes found.*", "")
  }

  lines
}

#' Format ISO 8601 date to a shorter display format
#' @noRd
format_date <- function(date_string) {
  if (is.null(date_string) || is.na(date_string) || nchar(date_string) == 0) {
    return("unknown date")
  }
  # Try to parse and format; fall back to raw string
  parsed <- tryCatch(
    as.Date(date_string),
    error = function(e) NULL
  )
  if (!is.null(parsed) && !is.na(parsed)) {
    format(parsed, "%Y-%m-%d")
  } else {
    date_string
  }
}
