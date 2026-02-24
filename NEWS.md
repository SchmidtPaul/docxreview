# docxreview 0.0.0.9000

* Initial development version.
* Added `extract_comments()` to extract comments with paragraph context.
* Added `extract_tracked_changes()` to extract insertions and deletions with paragraph context.
* Added `extract_review()` to combine both into formatted Markdown output.
* Added `anonymize` parameter to `extract_review()`: replaces reviewer names
  with neutral aliases ("Reviewer 1" etc.) in all output fields, including
  `@mentions` in comment text. Writes a sidecar `*-reviewers.md` mapping file
  when `output_file` is set.
* Added `group_by_author` parameter to `extract_review()` to group feedback
  by reviewer.
* Added `compare_versions()` to produce a diff `.docx` with tracked changes
  between two document versions (requires Microsoft Word + Python/pywin32,
  Windows only).
* Added `page` and `section` columns to `extract_comments()` and
  `extract_tracked_changes()` output.
