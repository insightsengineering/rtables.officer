# Changelog

## rtables.officer 0.1.2.9002

- Exported functions
  [`apply_alignments()`](https://insightsengineering.github.io/rtables.officer/reference/apply_alignments.md),
  [`apply_bold_manual()`](https://insightsengineering.github.io/rtables.officer/reference/apply_bold_manual.md)
  and
  [`extract_font_and_size_from_flx()`](https://insightsengineering.github.io/rtables.officer/reference/extract_font_and_size_from_flx.md)
  ([\#66](https://github.com/insightsengineering/rtables.officer/issues/66))
- Minor fix in
  [`apply_alignments()`](https://insightsengineering.github.io/rtables.officer/reference/apply_alignments.md)
  to update alignments in specific cells of the flextable.
- [`theme_docx_default()`](https://insightsengineering.github.io/rtables.officer/reference/tt_to_flextable.md)
  and
  [`theme_html_default()`](https://insightsengineering.github.io/rtables.officer/reference/tt_to_flextable.md)
  now default to `font_size = 8` (was 9) to match the standard TLG body
  font size. Pass `font_size = 9` to restore the previous output.
  ([\#71](https://github.com/insightsengineering/rtables.officer/issues/71))

## rtables.officer 0.1.2

CRAN release: 2026-01-08

- Fix issue with extra borders when header had `"\n"` special characters
  in `tt_as_flextable()`.

## rtables.officer 0.1.1

CRAN release: 2025-09-23

- Adding letter/A4 landscape/portrait docx templates.
- Dependency version bump for `officer` version 0.7.0 and `flextable`
  version 0.9.10.

## rtables.officer 0.1.0

CRAN release: 2025-04-22

- Added option to start new pages when exporting different paginated
  tables to `.docx`. It is possible to put tables on separate pages by
  doing `export_as_docx(add_page_break = TRUE)`.
- Added exporter functions for `rlistings` objects.
- Added separator lines and padding when present in the `rlistings` or
  `rtables` objects.
- Fixed bugs impeding pagination of lists of `rlistings` or `rtables`
  objects.
- Changed handler for footnotes printout from `footers_as_text` to
  `integrate_footers`. The behavior now is directly aligned with
  `titles_as_header`.
- Improved pagination documentation related to footnotes and titles.

## rtables.officer 0.0.2

CRAN release: 2025-01-17

- Experimental pagination is now possible in `tt_as_flextable()` and
  [`export_as_docx()`](https://insightsengineering.github.io/rtables.officer/reference/export_as_docx.md).
- Added handling of widths in `tt_as_flextable()`. Now it is possible to
  change column widths for `.docx` exports.
- First version of `rtables.officer`, split from the `rtables` package.
