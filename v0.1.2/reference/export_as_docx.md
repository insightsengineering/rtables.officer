# Export to a Word document

From an `rtables` table, produce a self-contained Word document or
attach it to a template Word file (`template_file`). This function is
based on the
[`tt_to_flextable()`](https://insightsengineering.github.io/rtables.officer/reference/tt_to_flextable.md)
transformer and the `officer` package.

## Usage

``` r
export_as_docx(
  tt,
  file,
  add_page_break = FALSE,
  add_template_page_numbers = TRUE,
  titles_as_header = TRUE,
  integrate_footers = TRUE,
  section_properties = section_properties_default(),
  doc_metadata = NULL,
  template_file = NULL,
  ...
)

section_properties_default(
  page_size = c("letter", "A4"),
  orientation = c("portrait", "landscape")
)

margins_potrait()

margins_landscape()
```

## Arguments

- tt:

  (`TableTree` or related class)  
  a `TableTree` object representing a populated table.

- file:

  (`string`)  
  output file. Must have `.docx` extension.

- add_page_break:

  (`flag`)  
  whether to add a page break after the table (`TRUE`) or not (`FALSE`).

- add_template_page_numbers:

  (`flag`)  
  whether to add page numbers to the word document as page footer. This
  uses templates to achieve it. Defaults to `TRUE`. Consider adding your
  own template file if you want more customization.

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

- section_properties:

  ([`officer::prop_section`](https://davidgohel.github.io/officer/reference/prop_section.html))  
  an
  [`officer::prop_section()`](https://davidgohel.github.io/officer/reference/prop_section.html)
  object which sets margins and page size. Defaults to
  `section_properties_default()`.

- doc_metadata:

  (`list` of `string`)  
  any value that can be used as metadata by
  [`officer::set_doc_properties()`](https://davidgohel.github.io/officer/reference/set_doc_properties.html).
  Important text values are `title`, `subject`, `creator`, and
  `description`, while `created` is a date object.

- template_file:

  (`string`)  
  template file that `officer` will use as a starting point for the
  final document. Document attaches the table and uses the defaults
  defined in the template file.

- ...:

  (`any`)  
  additional arguments passed to
  [`tt_to_flextable()`](https://insightsengineering.github.io/rtables.officer/reference/tt_to_flextable.md).

- page_size:

  (`string`) page size. Can be `"letter"` or `"A4"`. Defaults to
  `"letter"`.

- orientation:

  (`string`) page orientation. Can be `"portrait"` or `"landscape"`.
  Defaults to `"portrait"`.

## Value

No return value, called for side effects

## Details

Pagination Behavior for Titles and Footers (this behavior is
experimental at the moment):

The rendering of titles and footers interacts with table pagination as
follows:

- **Titles:** When `titles_as_header = TRUE` (default), the integrated
  title header rows typically repeat at the top of each new page if the
  table spans multiple pages. Setting `titles_as_header = FALSE` renders
  titles as a separate paragraph only once before the table begins.

- **Footers:** Regardless of the `integrate_footers` setting, footers
  appear only once. Integrated footnotes (`integrate_footers = TRUE`)
  appear at the very end of the complete table, and separate text
  paragraphs (`integrate_footers = FALSE`) appear after the complete
  table. Footers do not repeat on each page.

## Functions

- `section_properties_default()`: Helper function that defines standard
  portrait properties for tables.

- `margins_potrait()`: Helper function that defines standard portrait
  margins for tables.

- `margins_landscape()`: Helper function that defines standard landscape
  margins for tables.

## Note

`export_as_docx()` has few customization options available. If you
require specific formats and details, we suggest that you use
[`tt_to_flextable()`](https://insightsengineering.github.io/rtables.officer/reference/tt_to_flextable.md)
prior to `export_as_docx()`. If the table is modified first using
[`tt_to_flextable()`](https://insightsengineering.github.io/rtables.officer/reference/tt_to_flextable.md),
the `titles_as_header` and `integrate_footers` parameters must be
re-specified.

## See also

[`tt_to_flextable()`](https://insightsengineering.github.io/rtables.officer/reference/tt_to_flextable.md)

## Examples

``` r
lyt <- basic_table() %>%
  split_cols_by("ARM") %>%
  analyze(c("AGE", "BMRKR2", "COUNTRY"))

tbl <- build_table(lyt, ex_adsl)

# See how the section_properties_portrait() function is built for customization
tf <- tempfile(tmpdir = tempdir(check = TRUE), fileext = ".docx")
export_as_docx(tbl,
  file = tf,
  section_properties = section_properties_default(orientation = "landscape")
)
```
