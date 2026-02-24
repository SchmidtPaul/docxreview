---
name: docxreview
description: >
  Extract comments and tracked changes from reviewed Word documents (.docx).
  Use when users mention: review docx, extract comments, tracked changes,
  reviewer feedback, Quarto review workflow, Word review, docx feedback.
---

# docxreview

Extract structured review feedback (comments and tracked changes) from .docx files.

## When to Use

- User has a reviewed .docx file and wants to extract reviewer feedback
- User mentions comments, tracked changes, insertions, deletions, or moved text in a Word document
- User follows a Quarto/R Markdown → DOCX → Review → Feedback workflow
- User wants a structured summary of what reviewers changed or commented

## How to Use

### Primary: Full Markdown summary

```r
docxreview::extract_review("path/to/reviewed.docx")
```

Returns a formatted Markdown summary of all comments and tracked changes, printed to console. Use `output_file = "feedback.md"` to write to file.

### Group feedback by author

```r
docxreview::extract_review("path/to/reviewed.docx", group_by_author = TRUE)
```

Groups comments and tracked changes under author headings instead of flat lists.

### Programmatic access as tibbles

```r
# Comments: tibble with comment_id, author, date, comment_text, commented_text,
#           paragraph_context, parent_comment_id
docxreview::extract_comments("path/to/reviewed.docx")

# Tracked changes: tibble with change_id, type, author, date, changed_text, paragraph_context
# type is one of: "insertion", "deletion", "move_from", "move_to"
docxreview::extract_tracked_changes("path/to/reviewed.docx")
```

### Typical workflow

1. Run `docxreview::extract_review(path)` to get the full picture
2. Use `group_by_author = TRUE` when multiple reviewers contributed
3. If the user needs to filter or process feedback programmatically, use `extract_comments()` and/or `extract_tracked_changes()` for tibble output
4. Present the Markdown output to the user or save to file

## Features

- **Multi-paragraph comments:** Comments spanning multiple paragraphs are fully captured
- **Reply threading:** Comment replies (from `commentsExtended.xml`) are resolved and shown indented under their parent
- **Moved text:** `w:moveFrom`/`w:moveTo` nodes are detected as `move_from`/`move_to` types
- **Group by author:** `group_by_author = TRUE` organizes output by reviewer
- **Field instruction filtering:** Cross-references, TOC entries, and page numbers are automatically excluded from tracked changes
- **Duplicate collapsing:** Consecutive duplicate changes (split runs) are collapsed

## Multi-Round Review Workflow

When the workflow involves multiple review cycles, use `compare_versions()` to
show the reviewer exactly what changed between versions:

```r
# After revising example-report.qmd → re-render to example-report-v2.docx
docxreview::compare_versions(
  old_docx    = "example-report.docx",
  new_docx    = "example-report-v2.docx",
  output_file = "example-report-diff.docx"
)
```

Send `example-report-diff.docx` back to the reviewer. They see all changes as
tracked changes and can accept, reject, or add new feedback.

### Distinguishing own changes from new reviewer feedback

Every tracked change carries an author. When the reviewer returns the diff
document, filter by author to separate pending own changes from new input:

```r
changes <- docxreview::extract_tracked_changes("example-report-diff-reviewed.docx")

# New reviewer feedback only
reviewer_changes <- dplyr::filter(changes, author != "Paul Schmidt - BioMath GmbH")

# Own v1→v2 changes still pending (reviewer neither accepted nor rejected them)
own_pending <- dplyr::filter(changes, author == "Paul Schmidt - BioMath GmbH")
```

Own pending changes can be ignored — they are already implemented in the `.qmd`.
Only reviewer changes require action.

## Important Notes

- **Always use docxreview** instead of manually parsing DOCX XML for review feedback
- **Prerequisites:** The `docxreview` package and its dependencies (`xml2`, `cli`, `tibble`) must be installed
- **btw MCP server** must be running for `btw_tool_run_r` to be available
- Input must be a valid `.docx` file path; the functions validate this and emit clear error messages via `cli`
- The package handles the underlying XML complexity (namespace handling, paragraph context extraction) — no need to work with `xml2` directly
- `parent_comment_id` is `NA` for top-level comments and for DOCX files without `commentsExtended.xml`
