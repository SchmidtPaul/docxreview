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
#' @param group_by_author If `TRUE`, the Markdown output is grouped by author
#'   with each author as a top-level section. Defaults to `FALSE`.
#' @return Invisibly returns the Markdown string. If `output_file` is specified,
#'   the Markdown is also written to that file.
#' @export
#' @examples
#' # Print review feedback to console
#' \dontrun{
#' extract_review("report_reviewed.docx")
#'
#' # Group feedback by author
#' extract_review("report_reviewed.docx", group_by_author = TRUE)
#'
#' # Save to file
#' extract_review("report_reviewed.docx", output_file = "feedback.md")
#' }
extract_review <- function(docx_path, output_file = NULL,
                           group_by_author = FALSE) {
  check_docx_path(docx_path)

  # Parse DOCX once, pass pre-parsed XML to both extractors
  parts <- read_docx_xml(docx_path)
  comments <- extract_comments_impl(parts)
  changes <- extract_tracked_changes_impl(parts)

  if (group_by_author) {
    md <- format_review_markdown_by_author(comments, changes,
                                           basename(docx_path))
  } else {
    md <- format_review_markdown(comments, changes, basename(docx_path))
  }

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

# -- Per-item formatters (shared by both layout modes) -------------------------

#' Format a single comment as Markdown lines
#' @noRd
format_comment_md <- function(row, index) {
  date_fmt <- format_date(row$date)
  c(
    paste0("### ", index, ". ", row$author, " (", date_fmt, ")"),
    paste0('> **Commented text:** "', row$commented_text, '"'),
    paste0('> **Paragraph:** "', row$paragraph_context, '"'),
    "",
    paste0("**Comment:** ", row$comment_text),
    "",
    "---",
    ""
  )
}

#' Format a single tracked change as Markdown lines
#' @noRd
format_change_md <- function(row, index) {
  date_fmt <- format_date(row$date)
  type_labels <- c(
    deletion = "Deletion", insertion = "Insertion",
    move_from = "Moved from here", move_to = "Moved to here"
  )
  text_labels <- c(
    deletion = "Deleted", insertion = "Inserted",
    move_from = "Moved text", move_to = "Moved text"
  )
  type_label <- type_labels[row$type]
  text_label <- text_labels[row$type]

  c(
    paste0("### ", index, ". [", type_label, "] ", row$author,
           " (", date_fmt, ")"),
    paste0('> **', text_label, ':** "', row$changed_text, '"'),
    paste0('> **Paragraph:** "', row$paragraph_context, '"'),
    "",
    "---",
    ""
  )
}

# -- Flat layout (default) ----------------------------------------------------

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
    lines <- c(lines, format_comments_with_threading(comments))
  } else {
    lines <- c(lines, "*No comments found.*", "")
  }

  # Tracked changes section
  lines <- c(lines, paste0("## Tracked Changes (", nrow(changes), ")"), "")

  if (nrow(changes) > 0) {
    for (i in seq_len(nrow(changes))) {
      lines <- c(lines, format_change_md(changes[i, ], i))
    }
  } else {
    lines <- c(lines, "*No tracked changes found.*", "")
  }

  lines
}

# -- Grouped-by-author layout -------------------------------------------------

#' Format review data as Markdown grouped by author
#' @noRd
format_review_markdown_by_author <- function(comments, changes, filename) {
  lines <- c(paste0("# Review Feedback: ", filename), "")

  # Collect all unique authors (preserving document order)
  all_authors <- unique(c(comments$author, changes$author))
  all_authors[is.na(all_authors)] <- "Unknown Author"

  for (author in all_authors) {
    lines <- c(lines, paste0("## ", author), "")

    # Comments for this author
    auth_comments <- comments[
      !is.na(comments$author) & comments$author == author, , drop = FALSE
    ]
    if (author == "Unknown Author") {
      auth_comments <- rbind(
        auth_comments,
        comments[is.na(comments$author), , drop = FALSE]
      )
    }
    lines <- c(lines, paste0("### Comments (", nrow(auth_comments), ")"), "")

    if (nrow(auth_comments) > 0) {
      counter <- 0L
      for (i in seq_len(nrow(auth_comments))) {
        row <- auth_comments[i, ]
        # Skip replies (they're shown under their parent)
        if ("parent_comment_id" %in% names(row) &&
            !is.na(row$parent_comment_id)) next
        counter <- counter + 1L
        date_fmt <- format_date(row$date)
        lines <- c(lines,
          paste0("#### ", counter, ". (", date_fmt, ")"),
          paste0('> **Commented text:** "', row$commented_text, '"'),
          paste0('> **Paragraph:** "', row$paragraph_context, '"'),
          "",
          paste0("**Comment:** ", row$comment_text),
          "",
          ""
        )
        # Append replies from any author
        if ("parent_comment_id" %in% names(comments)) {
          replies <- comments[
            !is.na(comments$parent_comment_id) &
              comments$parent_comment_id == row$comment_id, , drop = FALSE
          ]
          for (r in seq_len(nrow(replies))) {
            rr <- replies[r, ]
            lines <- c(lines,
              paste0("> **Reply by ", rr$author, " (", format_date(rr$date),
                     "):** ", rr$comment_text),
              ""
            )
          }
        }
        lines <- c(lines, "---", "")
      }
    } else {
      lines <- c(lines, "*No comments.*", "")
    }

    # Changes for this author
    auth_changes <- changes[
      !is.na(changes$author) & changes$author == author, , drop = FALSE
    ]
    if (author == "Unknown Author") {
      auth_changes <- rbind(
        auth_changes,
        changes[is.na(changes$author), , drop = FALSE]
      )
    }
    lines <- c(lines, paste0("### Tracked Changes (", nrow(auth_changes), ")"),
               "")

    if (nrow(auth_changes) > 0) {
      type_labels <- c(
        deletion = "Deletion", insertion = "Insertion",
        move_from = "Moved from here", move_to = "Moved to here"
      )
      text_labels <- c(
        deletion = "Deleted", insertion = "Inserted",
        move_from = "Moved text", move_to = "Moved text"
      )

      for (i in seq_len(nrow(auth_changes))) {
        row <- auth_changes[i, ]
        date_fmt <- format_date(row$date)
        type_label <- type_labels[row$type]
        text_label <- text_labels[row$type]

        lines <- c(lines,
          paste0("#### ", i, ". [", type_label, "] (", date_fmt, ")"),
          paste0('> **', text_label, ':** "', row$changed_text, '"'),
          paste0('> **Paragraph:** "', row$paragraph_context, '"'),
          "",
          "---",
          ""
        )
      }
    } else {
      lines <- c(lines, "*No tracked changes.*", "")
    }
  }

  lines
}

# -- Comment threading in markdown ----------------------------------------------

#' Format comments with reply threading
#'
#' Top-level comments are numbered. Replies appear indented below their parent.
#' @param comments Tibble from extract_comments() (with parent_comment_id).
#' @return Character vector of Markdown lines.
#' @noRd
format_comments_with_threading <- function(comments) {
  has_threading <- "parent_comment_id" %in% names(comments) &&
    any(!is.na(comments$parent_comment_id))

  if (!has_threading) {
    # No threading — flat list
    lines <- character()
    for (i in seq_len(nrow(comments))) {
      lines <- c(lines, format_comment_md(comments[i, ], i))
    }
    return(lines)
  }

  # Threaded: show top-level comments with their replies
  top_level <- comments[is.na(comments$parent_comment_id), , drop = FALSE]
  lines <- character()
  counter <- 0L

  for (i in seq_len(nrow(top_level))) {
    counter <- counter + 1L
    row <- top_level[i, ]
    date_fmt <- format_date(row$date)
    lines <- c(lines,
      paste0("### ", counter, ". ", row$author, " (", date_fmt, ")"),
      paste0('> **Commented text:** "', row$commented_text, '"'),
      paste0('> **Paragraph:** "', row$paragraph_context, '"'),
      "",
      paste0("**Comment:** ", row$comment_text),
      ""
    )

    # Find and append replies
    replies <- comments[
      !is.na(comments$parent_comment_id) &
        comments$parent_comment_id == row$comment_id, , drop = FALSE
    ]
    for (j in seq_len(nrow(replies))) {
      rrow <- replies[j, ]
      lines <- c(lines,
        paste0("> **Reply by ", rrow$author, " (", format_date(rrow$date),
               "):** ", rrow$comment_text),
        ""
      )
    }

    lines <- c(lines, "---", "")
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
