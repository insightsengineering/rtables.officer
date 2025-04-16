## rtables.officer 0.0.2.9003

 * Added option to flush pages when exporting to `.docx`. It is possible to put tables on separate pages by doing `export_as_docx(add_page_break = TRUE)`.
 * Now it is possible to transform and export also `rlistings` objects.
 * Added separator lines and padding when present in the `rlistings` or `rtables` objects.
 * Fixed bugs impeding pagination of lists of `rlistings` or `rtables` objects.
 * Changed handler for footnotes printout from `footers_as_text` to `integrate_footers`. The behavior now is directly aligned with `titles_as_header`.
 * Improved pagination documentation related to footnotes and titles.

## rtables.officer 0.0.2

 * Experimental pagination is now possible in `tt_as_flextable()` and `export_as_docx()`.
 * Added handling of widths in `tt_as_flextable()`. Now it is possible to change column widths for `.docx` exports.
 * First version of `rtables.officer`, split from the `rtables` package.
