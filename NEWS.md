## rtables.officer 0.1.1.9004

 * Fix issue with extra borders when header had `"\n"` special characters in `tt_as_flextable()`.

## rtables.officer 0.1.1

 * Adding letter/A4 landscape/portrait docx templates.
 * Dependency version bump for `officer` version 0.7.0 and `flextable` version 0.9.10.

## rtables.officer 0.1.0

 * Added option to start new pages when exporting different paginated tables to `.docx`. It is possible to put tables on separate pages by doing `export_as_docx(add_page_break = TRUE)`.
 * Added exporter functions for `rlistings` objects.
 * Added separator lines and padding when present in the `rlistings` or `rtables` objects.
 * Fixed bugs impeding pagination of lists of `rlistings` or `rtables` objects.
 * Changed handler for footnotes printout from `footers_as_text` to `integrate_footers`. The behavior now is directly aligned with `titles_as_header`.
 * Improved pagination documentation related to footnotes and titles.

## rtables.officer 0.0.2

 * Experimental pagination is now possible in `tt_as_flextable()` and `export_as_docx()`.
 * Added handling of widths in `tt_as_flextable()`. Now it is possible to change column widths for `.docx` exports.
 * First version of `rtables.officer`, split from the `rtables` package.
