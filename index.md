# docxreview

docxreview extracts comments and tracked changes from reviewed Word
documents (.docx) and returns them as structured data or formatted
Markdown. It is designed for workflows where documents are rendered from
Quarto or R Markdown, reviewed in Microsoft Word, and feedback needs to
be processed programmatically — particularly with [Claude
Code](https://docs.anthropic.com/en/docs/claude-code) and
[btw](https://github.com/posit-dev/btw) for AI-assisted review
processing.

## Installation

Install the development version from GitHub:

``` r
# install.packages("remotes")
remotes::install_github("SchmidtPaul/docxreview")
```

## Quick Start

``` r
library(docxreview)

# Get a formatted Markdown summary of all review feedback
extract_review("report_reviewed.docx")

# Save to file
extract_review("report_reviewed.docx", output = "feedback.md")

# Programmatic access as tibbles
comments <- extract_comments("report_reviewed.docx")
changes  <- extract_tracked_changes("report_reviewed.docx")
```

## Workflow

The typical use case is a Quarto/R Markdown review cycle:

``` mermaid
flowchart LR
    A[".qmd Source"] -->|"quarto render"| B[".docx Report"]
    B -->|"send to reviewer"| C["Reviewer edits in Word"]
    C -->|"reviewed .docx"| D["docxreview extracts feedback"]
    D -->|"structured Markdown"| E["Author revises source"]
```

1.  **Render** your Quarto or R Markdown document to `.docx`.
2.  **Send** the `.docx` to a reviewer who adds comments and tracked
    changes in Microsoft Word.
3.  **Extract** the feedback with
    [`extract_review()`](https://schmidtpaul.github.io/docxreview/reference/extract_review.md)
    to get a structured Markdown summary of all comments, insertions,
    and deletions — including paragraph context.
4.  **Revise** your source document based on the extracted feedback.

## Claude Code Integration

docxreview ships with a [Claude
Code](https://docs.anthropic.com/en/docs/claude-code) skill file at
`inst/skills/docxreview/SKILL.md`. When used together with
[btw](https://github.com/posit-dev/btw) as an MCP server, Claude Code
can automatically extract and process review feedback from `.docx`
files.

To make the skill available, add the installed package’s skill directory
to your project’s `.claude/settings.json`:

``` json
{
  "permissions": {
    "allow": []
  },
  "skills": [
    "/path/to/docxreview/skills/docxreview/"
  ]
}
```

Or find the path programmatically:

``` r
system.file("skills", "docxreview", package = "docxreview")
```

## Functions

| Function                                                                                                     | Returns              | Description                                                                     |
|--------------------------------------------------------------------------------------------------------------|----------------------|---------------------------------------------------------------------------------|
| [`extract_review()`](https://schmidtpaul.github.io/docxreview/reference/extract_review.md)                   | Markdown (character) | Full formatted summary of all comments and tracked changes                      |
| [`extract_comments()`](https://schmidtpaul.github.io/docxreview/reference/extract_comments.md)               | tibble               | Comments with author, date, comment text, marked text, and paragraph context    |
| [`extract_tracked_changes()`](https://schmidtpaul.github.io/docxreview/reference/extract_tracked_changes.md) | tibble               | Insertions and deletions with author, date, changed text, and paragraph context |

See
[`vignette("docxreview")`](https://schmidtpaul.github.io/docxreview/articles/docxreview.md)
for the full workflow documentation.
