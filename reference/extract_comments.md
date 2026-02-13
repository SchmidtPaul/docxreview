# Extract comments from a Word document

Reads a .docx file and extracts all comments including the commented
(highlighted) text, the comment text, author, date, and the full
paragraph context in which the comment is anchored.

## Usage

``` r
extract_comments(docx_path)
```

## Arguments

- docx_path:

  Path to a .docx file.

## Value

A [tibble::tibble](https://tibble.tidyverse.org/reference/tibble.html)
with columns:

- `comment_id`: Character. The comment ID from the document.

- `author`: Character. Name of the comment author.

- `date`: Character. ISO 8601 timestamp of the comment.

- `comment_text`: Character. The comment body text.

- `commented_text`: Character. The text range that was
  highlighted/selected.

- `paragraph_context`: Character. Full text of the paragraph(s)
  containing the commented range.

## Examples

``` r
# Extract comments from a reviewed document
if (FALSE) { # \dontrun{
comments <- extract_comments("report_reviewed.docx")
comments
} # }
```
