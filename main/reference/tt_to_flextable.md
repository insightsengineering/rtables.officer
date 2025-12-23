# Create a `flextable` from an `rtables` table

Principally used within
[`export_as_docx()`](https://insightsengineering.github.io/rtables.officer/reference/export_as_docx.md),
this function produces a `flextable` from an `rtables` table. If
`theme = theme_docx_default()` (default), a `.docx`-friendly table will
be produced. If `theme = NULL`, the table will be produced in an
`rtables`-like style.

## Usage

``` r
tt_to_flextable(
  tt,
  theme = theme_docx_default(),
  border = flextable::fp_border_default(width = 0.5),
  indent_size = NULL,
  titles_as_header = TRUE,
  bold_titles = TRUE,
  integrate_footers = TRUE,
  counts_in_newline = FALSE,
  paginate = FALSE,
  fontspec = NULL,
  lpp = NULL,
  cpp = NULL,
  ...,
  colwidths = NULL,
  tf_wrap = !is.null(cpp),
  max_width = cpp,
  total_page_height = 10,
  total_page_width = 10,
  autofit_to_page = TRUE
)

theme_docx_default(
  font = "Arial",
  font_size = 9,
  cell_margins = c(word_mm_to_pt(1.9), word_mm_to_pt(1.9), 0, 0),
  bold = c("header", "content_rows", "label_rows", "top_left"),
  bold_manual = NULL,
  border = flextable::fp_border_default(width = 0.5)
)

theme_html_default(
  font = "Courier",
  font_size = 9,
  cell_margins = 0.2,
  remove_internal_borders = "label_rows",
  border = flextable::fp_border_default(width = 1, color = "black")
)

word_mm_to_pt(mm)
```

## Arguments

- tt:

  (`TableTree`, `listing_df`, or related class)  
  a `TableTree` or `listing_df` object representing a populated table or
  listing.

- theme:

  (`function` or `NULL`)  
  a theme function designed to change the layout and style of a
  `flextable` object. Defaults to `theme_docx_default()`, the classic
  Microsoft Word output style. If `NULL`, a table with style similar to
  the `rtables` default will be produced. See Details below for more
  information.

- border:

  (`flextable::fp_border()`)  
  border style. Defaults to `flextable::fp_border_default(width = 0.5)`.

- indent_size:

  (`numeric(1)`)  
  indentation size. If `NULL`, the default indent size of the table (see
  [`formatters::matrix_form()`](https://insightsengineering.github.io/formatters/latest-tag/reference/matrix_form.html)
  `indent_size`, default is 2) is used. To work with `docx`, any size is
  multiplied by 1 mm (2.83 pt) by default.

- titles_as_header:

  (`flag`)  
  Controls how titles are rendered relative to the table. If `TRUE`
  (default), the main title
  ([`formatters::main_title()`](https://insightsengineering.github.io/formatters/latest-tag/reference/title_footer.html))
  and subtitles
  ([`formatters::subtitles()`](https://insightsengineering.github.io/formatters/latest-tag/reference/title_footer.html))
  are added as distinct header rows within the `flextable` object
  itself. If `FALSE`, titles are rendered as a separate paragraph of
  text placed immediately before the table.

- bold_titles:

  (`flag` or `integer`)  
  whether titles should be bold (defaults to `TRUE`). If one or more
  integers are provided, these integers are used as indices for lines at
  which titles should be bold.

- integrate_footers:

  (`flag`)  
  Controls how footers are rendered relative to the table. If `TRUE`
  (default), footers (e.g.,
  [`formatters::main_footer()`](https://insightsengineering.github.io/formatters/latest-tag/reference/title_footer.html),
  [`formatters::prov_footer()`](https://insightsengineering.github.io/formatters/latest-tag/reference/title_footer.html))
  are integrated directly into the `flextable` object, typically
  appearing as footnotes below the table body with a smaller font. If
  `FALSE`, footers are rendered as a separate paragraph of text placed
  immediately after the table.

- counts_in_newline:

  (`flag`)  
  whether column counts should be printed on a new line. In `rtables`,
  column counts (i.e. `(N=xx)`) are always printed on a new line
  (`TRUE`). For `docx` exports it may be preferred to print these counts
  on the same line (`FALSE`). Defaults to `FALSE`.

- paginate:

  (`flag`)  
  whether the `rtables` pagination mechanism should be used. If `TRUE`,
  this option splits `tt` into multiple `flextables` as different
  "pages". When using
  [`export_as_docx()`](https://insightsengineering.github.io/rtables.officer/reference/export_as_docx.md)
  we suggest setting this to `FALSE` and relying only on the default
  Microsoft Word pagination system as co-operation between the two
  mechanisms is not guaranteed. Defaults to `FALSE`.

- fontspec:

  (`font_spec`)  
  a font_spec object specifying the font information to use for
  calculating string widths and heights, as returned by
  [`font_spec()`](https://insightsengineering.github.io/formatters/latest-tag/reference/font_spec.html).

- lpp:

  (`numeric(1)`)  
  maximum lines per page including (re)printed header and context rows.

- cpp:

  (`numeric(1)` or `NULL`)  
  width (in characters) of the pages for horizontal pagination. `NA`
  (the default) indicates `cpp` should be inferred from the page size;
  `NULL` indicates no horizontal pagination should be done regardless of
  page size.

- ...:

  (`any`)  
  additional parameters to be passed to the pagination function. See
  [`rtables::paginate_table()`](https://insightsengineering.github.io/rtables/latest-tag/reference/paginate.html)
  for options. If `paginate = FALSE` this argument is ignored.

- colwidths:

  (`numeric`)  
  column widths for the resulting flextable(s). If `NULL`, the column
  widths estimated with
  [`formatters::propose_column_widths()`](https://insightsengineering.github.io/formatters/latest-tag/reference/propose_column_widths.html)
  will be used. When exporting into `.docx` these values are normalized
  to represent a fraction of the `total_page_width`. If these are
  specified, `autofit_to_page` is set to `FALSE`.

- tf_wrap:

  (`flag`)  
  whether the text for title, subtitles, and footnotes should be
  wrapped.

- max_width:

  (`integer(1)`, `string` or `NULL`)  
  width that title and footer (including footnotes) materials should be
  word-wrapped to. If `NULL`, it is set to the current print width of
  the session (`getOption("width")`). If set to `"auto"`, the width of
  the table (plus any table inset) is used. Parameter is ignored if
  `tf_wrap = FALSE`.

- total_page_height:

  (`numeric(1)`)  
  total page height (in inches) for the resulting flextable(s). Used
  only to estimate number of lines per page (`lpp`) when
  `paginate = TRUE`. Defaults to 10.

- total_page_width:

  (`numeric(1)`)  
  total page width (in inches) for the resulting flextable(s). Any
  values added for column widths are normalized by the total page width.
  Defaults to 10. If `autofit_to_page = TRUE`, this value is
  automatically set to the allowed page width.

- autofit_to_page:

  (`flag`)  
  whether column widths should be automatically adjusted to fit the
  total page width. If `FALSE`, `colwidths` is used to indicate
  proportions of `total_page_width`. Defaults to `TRUE`. See
  `flextable::set_table_properties(layout)` for more details.

- font:

  (`string`)  
  font. Defaults to `"Arial"`. If the font given is not available, the
  `flextable` default is used instead. For options, consult the family
  column from
  [`systemfonts::system_fonts()`](https://systemfonts.r-lib.org/reference/system_fonts.html).

- font_size:

  (`integer(1)`)  
  font size. Defaults to 9.

- cell_margins:

  (`numeric(1)` or `numeric(4)`)  
  a numeric or a vector of four numbers indicating
  `c("left", "right", "top", "bottom")`. It defaults to 0 for top and
  bottom, and to 0.19 `mm` in Word `pt` for left and right.

- bold:

  (`character`)  
  parts of the table text that should be in bold. Can be any combination
  of `c("header", "content_rows", "label_rows", "top_left")`. The first
  one renders all column names bold (not `topleft` content). The second
  and third option use
  [`formatters::make_row_df()`](https://insightsengineering.github.io/formatters/latest-tag/reference/make_row_df.html)
  to render content or/and label rows as bold.

- bold_manual:

  (named `list` or `NULL`)  
  list of index lists. See example for needed structure. Accepted
  groupings/names are `c("header", "body")`.

- remove_internal_borders:

  (`character`)  
  where to remove internal borders between rows. Defaults to
  `"label_rows"`. Currently there are no other options and this can be
  turned off by providing any other character value.

- mm:

  (`numeric(1)`)  
  the value in mm to transform to pt.

## Value

A `flextable` object.

## Details

If you would like to make a minor change to a pre-existing style, this
can be done by extending themes. You can do this by either adding your
own theme to the theme call (e.g.
`theme = c(theme_docx_default(), my_theme)`) or creating a new theme as
shown in the examples below. Please pay close attention to the
parameters' inputs.

It is possible to use some hidden values to build your own theme (hence
the need for the `...` parameter). In particular, `tt_to_flextable()`
uses the following variable:
`tbl_row_class = rtables::make_row_df(tt)$node_class`. This is ignored
if not used in the theme. See `theme_docx_default()` for an example on
how to retrieve and use these values.

## Functions

- `theme_docx_default()`: Main theme function for
  [`export_as_docx()`](https://insightsengineering.github.io/rtables.officer/reference/export_as_docx.md).

- `theme_html_default()`: Theme function for html outputs.

- `word_mm_to_pt()`: Padding helper functions to transform mm to pt.

## Note

Currently `cpp`, `tf_wrap`, and `max_width` are only used in pagination
and should be used cautiously if used in combination with `colwidths`
and `autofit_to_page`. If issues arise, please raise an issue on GitHub
or communicate this to the package maintainers directly.

## See also

[`export_as_docx()`](https://insightsengineering.github.io/rtables.officer/reference/export_as_docx.md)

## Examples

``` r
analysisfun <- function(x, ...) {
  in_rows(
    row1 = 5,
    row2 = c(1, 2),
    .row_footnotes = list(row1 = "row 1 - row footnote"),
    .cell_footnotes = list(row2 = "row 2 - cell footnote")
  )
}

lyt <- basic_table(
  title = "Title says Whaaaat", subtitles = "Oh, ok.",
  main_footer = "ha HA! Footer!"
) %>%
  split_cols_by("ARM") %>%
  analyze("AGE", afun = analysisfun)

tbl <- build_table(lyt, ex_adsl)

# Example 1: rtables style ---------------------------------------------------
tt_to_flextable(tbl, theme = NULL)


.cl-a466e9f2{table-layout:auto;}.cl-a46118ba{font-family:'DejaVu Sans';font-size:11pt;font-weight:bold;font-style:normal;text-decoration:none;color:rgba(0, 0, 0, 1.00);background-color:transparent;}.cl-a46118c4{font-family:'DejaVu Sans';font-size:11pt;font-weight:normal;font-style:normal;text-decoration:none;color:rgba(0, 0, 0, 1.00);background-color:transparent;}.cl-a4639216{margin:0;text-align:left;border-bottom: 0 solid rgba(0, 0, 0, 1.00);border-top: 0 solid rgba(0, 0, 0, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);padding-bottom:5pt;padding-top:5pt;padding-left:8.2pt;padding-right:5.4pt;line-height: 1;background-color:transparent;}.cl-a4639220{margin:0;text-align:left;border-bottom: 0 solid rgba(0, 0, 0, 1.00);border-top: 0 solid rgba(0, 0, 0, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);padding-bottom:5pt;padding-top:5pt;padding-left:5pt;padding-right:5pt;line-height: 1;background-color:transparent;}.cl-a4639221{margin:0;text-align:center;border-bottom: 0 solid rgba(0, 0, 0, 1.00);border-top: 0 solid rgba(0, 0, 0, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);padding-bottom:5pt;padding-top:5pt;padding-left:5pt;padding-right:5pt;line-height: 1;background-color:transparent;}.cl-a463922a{margin:0;text-align:left;border-bottom: 0 solid rgba(0, 0, 0, 1.00);border-top: 0 solid rgba(0, 0, 0, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);padding-bottom:5pt;padding-top:5pt;padding-left:5.4pt;padding-right:5.4pt;line-height: 1;background-color:transparent;}.cl-a463ac88{background-color:rgba(255, 255, 255, 1.00);vertical-align: middle;border-bottom: 0 solid rgba(102, 102, 102, 1.00);border-top: 0 solid rgba(102, 102, 102, 1.00);border-left: 0 solid rgba(102, 102, 102, 1.00);border-right: 0 solid rgba(102, 102, 102, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-a463ac92{background-color:rgba(255, 255, 255, 1.00);vertical-align: middle;border-bottom: 0 solid rgba(102, 102, 102, 1.00);border-top: 0 solid rgba(102, 102, 102, 1.00);border-left: 0 solid rgba(102, 102, 102, 1.00);border-right: 0 solid rgba(102, 102, 102, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-a463ac93{background-color:rgba(255, 255, 255, 1.00);vertical-align: middle;border-bottom: 0 solid rgba(102, 102, 102, 1.00);border-top: 0 solid rgba(102, 102, 102, 1.00);border-left: 0 solid rgba(102, 102, 102, 1.00);border-right: 0 solid rgba(102, 102, 102, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-a463ac9c{background-color:rgba(255, 255, 255, 1.00);vertical-align: middle;border-bottom: 0 solid rgba(102, 102, 102, 1.00);border-top: 0 solid rgba(102, 102, 102, 1.00);border-left: 0 solid rgba(102, 102, 102, 1.00);border-right: 0 solid rgba(102, 102, 102, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-a463ac9d{background-color:rgba(255, 255, 255, 1.00);vertical-align: middle;border-bottom: 0.5pt solid rgba(102, 102, 102, 1.00);border-top: 0 solid rgba(102, 102, 102, 1.00);border-left: 0 solid rgba(102, 102, 102, 1.00);border-right: 0 solid rgba(102, 102, 102, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-a463ac9e{background-color:rgba(255, 255, 255, 1.00);vertical-align: middle;border-bottom: 0.5pt solid rgba(102, 102, 102, 1.00);border-top: 0 solid rgba(102, 102, 102, 1.00);border-left: 0 solid rgba(102, 102, 102, 1.00);border-right: 0 solid rgba(102, 102, 102, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-a463aca6{background-color:rgba(255, 255, 255, 1.00);vertical-align: middle;border-bottom: 0.5pt solid rgba(102, 102, 102, 1.00);border-top: 0 solid rgba(102, 102, 102, 1.00);border-left: 0 solid rgba(102, 102, 102, 1.00);border-right: 0 solid rgba(102, 102, 102, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-a463aca7{background-color:rgba(255, 255, 255, 1.00);vertical-align: middle;border-bottom: 0.5pt solid rgba(102, 102, 102, 1.00);border-top: 0 solid rgba(102, 102, 102, 1.00);border-left: 0 solid rgba(102, 102, 102, 1.00);border-right: 0 solid rgba(102, 102, 102, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-a463aca8{background-color:transparent;vertical-align: middle;border-bottom: 0.5pt solid rgba(102, 102, 102, 1.00);border-top: 0.5pt solid rgba(102, 102, 102, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-a463acb0{background-color:transparent;vertical-align: middle;border-bottom: 0.5pt solid rgba(102, 102, 102, 1.00);border-top: 0.5pt solid rgba(102, 102, 102, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-a463acb1{background-color:transparent;vertical-align: middle;border-bottom: 0.5pt solid rgba(102, 102, 102, 1.00);border-top: 0.5pt solid rgba(102, 102, 102, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-a463acb2{background-color:transparent;vertical-align: middle;border-bottom: 0.5pt solid rgba(102, 102, 102, 1.00);border-top: 0.5pt solid rgba(102, 102, 102, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-a463acba{background-color:transparent;vertical-align: middle;border-bottom: 0 solid rgba(0, 0, 0, 1.00);border-top: 0 solid rgba(0, 0, 0, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-a463acbb{background-color:transparent;vertical-align: middle;border-bottom: 0 solid rgba(0, 0, 0, 1.00);border-top: 0 solid rgba(0, 0, 0, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-a463acbc{background-color:transparent;vertical-align: middle;border-bottom: 0 solid rgba(0, 0, 0, 1.00);border-top: 0 solid rgba(0, 0, 0, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-a463acbd{background-color:transparent;vertical-align: middle;border-bottom: 0 solid rgba(0, 0, 0, 1.00);border-top: 0 solid rgba(0, 0, 0, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-a463acc4{background-color:transparent;vertical-align: middle;border-bottom: 0.5pt solid rgba(102, 102, 102, 1.00);border-top: 0 solid rgba(0, 0, 0, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-a463acc5{background-color:transparent;vertical-align: middle;border-bottom: 0.5pt solid rgba(102, 102, 102, 1.00);border-top: 0 solid rgba(0, 0, 0, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-a463acce{background-color:transparent;vertical-align: middle;border-bottom: 0.5pt solid rgba(102, 102, 102, 1.00);border-top: 0 solid rgba(0, 0, 0, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-a463accf{background-color:transparent;vertical-align: middle;border-bottom: 0.5pt solid rgba(102, 102, 102, 1.00);border-top: 0 solid rgba(0, 0, 0, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-a463acd0{background-color:transparent;vertical-align: middle;border-bottom: 0 solid rgba(255, 255, 255, 0.00);border-top: 0 solid rgba(255, 255, 255, 0.00);border-left: 0 solid rgba(255, 255, 255, 0.00);border-right: 0 solid rgba(255, 255, 255, 0.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-a463acd8{background-color:transparent;vertical-align: middle;border-bottom: 0 solid rgba(255, 255, 255, 0.00);border-top: 0 solid rgba(255, 255, 255, 0.00);border-left: 0 solid rgba(255, 255, 255, 0.00);border-right: 0 solid rgba(255, 255, 255, 0.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-a463acd9{background-color:transparent;vertical-align: middle;border-bottom: 0 solid rgba(255, 255, 255, 0.00);border-top: 0 solid rgba(255, 255, 255, 0.00);border-left: 0 solid rgba(255, 255, 255, 0.00);border-right: 0 solid rgba(255, 255, 255, 0.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-a463ace2{background-color:transparent;vertical-align: middle;border-bottom: 0 solid rgba(255, 255, 255, 0.00);border-top: 0 solid rgba(255, 255, 255, 0.00);border-left: 0 solid rgba(255, 255, 255, 0.00);border-right: 0 solid rgba(255, 255, 255, 0.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}


Title says Whaaaat
```
