# Extract tracked changes from a Word document

Reads a .docx file and extracts all tracked changes (insertions and
deletions) including the changed text, author, date, and the full
paragraph context.

## Usage

``` r
extract_tracked_changes(docx_path)
```

## Arguments

- docx_path:

  Path to a .docx file.

## Value

A [tibble::tibble](https://tibble.tidyverse.org/reference/tibble.html)
with columns:

- `change_id`: Integer. Sequential ID for each change.

- `type`: Character. Either `"insertion"` or `"deletion"`.

- `author`: Character. Name of the person who made the change.

- `date`: Character. ISO 8601 timestamp of the change.

- `changed_text`: Character. The text that was inserted or deleted.

- `paragraph_context`: Character. Full text of the containing paragraph,
  with the change marked up (\~~deletion\~~ or **insertion**).

## Examples

``` r
# Extract tracked changes from a reviewed document
if (FALSE) { # \dontrun{
changes <- extract_tracked_changes("report_reviewed.docx")
changes
} # }
```
