# Extract all review feedback from a Word document

Extracts both comments and tracked changes from a reviewed .docx file
and returns a structured Markdown summary. This is the main entry point
for processing reviewer feedback.

## Usage

``` r
extract_review(docx_path, output = NULL)
```

## Arguments

- docx_path:

  Path to a .docx file.

- output:

  Optional file path to write the Markdown output to. If `NULL`
  (default), the Markdown string is returned invisibly and printed to
  the console.

## Value

Invisibly returns the Markdown string. If `output` is specified, the
Markdown is also written to that file.

## Examples

``` r
# Print review feedback to console
if (FALSE) { # \dontrun{
extract_review("report_reviewed.docx")

# Save to file
extract_review("report_reviewed.docx", output = "feedback.md")
} # }
```
