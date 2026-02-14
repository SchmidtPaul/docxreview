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
- User mentions comments, tracked changes, insertions, or deletions in a Word document
- User follows a Quarto/R Markdown → DOCX → Review → Feedback workflow
- User wants a structured summary of what reviewers changed or commented

## How to Use

### Primary: Full Markdown summary

```r
docxreview::extract_review("path/to/reviewed.docx")
```

Returns a formatted Markdown summary of all comments and tracked changes, printed to console. Use `output_file = "feedback.md"` to write to file.

### Programmatic access as tibbles

```r
# Comments: tibble with comment_id, author, date, comment_text, commented_text, paragraph_context
docxreview::extract_comments("path/to/reviewed.docx")

# Tracked changes: tibble with change_id, type, author, date, changed_text, paragraph_context
docxreview::extract_tracked_changes("path/to/reviewed.docx")
```

### Typical workflow

1. Run `docxreview::extract_review(path)` to get the full picture
2. If the user needs to filter or process feedback programmatically, use `extract_comments()` and/or `extract_tracked_changes()` for tibble output
3. Present the Markdown output to the user or save to file

## Important Notes

- **Always use docxreview** instead of manually parsing DOCX XML for review feedback
- **Prerequisites:** The `docxreview` package and its dependencies (`officer`, `xml2`, `cli`, `tibble`) must be installed
- **btw MCP server** must be running for `btw_tool_run_r` to be available
- Input must be a valid `.docx` file path; the functions validate this and emit clear error messages via `cli`
- The package handles the underlying XML complexity (namespace handling, paragraph context extraction) — no need to work with `xml2` directly
