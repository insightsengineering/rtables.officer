## rtables.officer 0.0.2.9002

### New Features
 * Added option to flush pages when exporting to `.docx`. It is possible to put tables on separate pages by doing `export_as_docx(add_page_break = TRUE)`.


## rtables.officer 0.0.2

### New Features
 * Experimental pagination is now possible in `tt_as_flextable()` and `export_as_docx()`.
 * Added handling of widths in `tt_as_flextable()`. Now it is possible to change column widths for `.docx` exports.
 * First version of `rtables.officer`, split from the `rtables` package.
