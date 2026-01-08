# Add Conditional Separators (horizontal line or padding) to `flextable` Rows

Modifies an existing `flextable` object by adding visual separators
(horizontal lines or bottom padding) after specific rows based on a
control vector, without changing the table's row count.

## Usage

``` r
add_flextable_separators(
  ft,
  trailing_sep,
  border = officer::fp_border(width = 1, color = "grey60"),
  padding = 10
)
```

## Arguments

- ft:

  A flextable object.

- trailing_sep:

  A vector specifying separators. Its length must equal the number of
  rows in the body of `ft`. Allowed values are NA (no separator), "-"
  (adds a horizontal line), or " " (adds bottom padding).

- border:

  The `fp_border` object to use for horizontal lines when `trailing_sep`
  is "-". Defaults to a gray line of width 1.

- padding:

  The amount of bottom padding (in points) to add when `trailing_sep` is
  " ". Defaults to 10.

## Value

The modified flextable object, or the original `ft` if all
`trailing_sep` values are NA. Throws an error for invalid inputs or
invalid characters in `trailing_sep`.

## Examples

``` r
content <- data.frame(
  USUBJID = c("S1", "S1", "S1", "S2", "S2", "S2", "S3"),
  ARM = c("A", "A", "B", "A", "A", "B", "A"),
  VAL = round(rnorm(7), 2)
)
ft <- flextable::as_flextable(content)
ft <- flextable::theme_booktabs(ft)

# Define separators: line, space, NA, line, space, NA, NA
sep_ctrl <- c("-", " ", NA, "-", " ", NA, NA)

ft_modified <- add_flextable_separators(ft, sep_ctrl)
print(ft_modified)
#> <style></style>
#> <div class="tabwid"><style>.cl-1aa3cf7a{}.cl-1a88aaec{font-family:'DejaVu Sans';font-size:11pt;font-weight:normal;font-style:normal;text-decoration:none;color:rgba(0, 0, 0, 1.00);background-color:transparent;}.cl-1a88aaf6{font-family:'DejaVu Sans';font-size:11pt;font-weight:normal;font-style:normal;text-decoration:none;color:rgba(153, 153, 153, 1.00);background-color:transparent;}.cl-1a8bfd3c{margin:0;text-align:left;border-bottom: 0 solid rgba(0, 0, 0, 1.00);border-top: 0 solid rgba(0, 0, 0, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);padding-bottom:5pt;padding-top:5pt;padding-left:5pt;padding-right:5pt;line-height: 1;background-color:transparent;}.cl-1a8bfd50{margin:0;text-align:right;border-bottom: 0 solid rgba(0, 0, 0, 1.00);border-top: 0 solid rgba(0, 0, 0, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);padding-bottom:5pt;padding-top:5pt;padding-left:5pt;padding-right:5pt;line-height: 1;background-color:transparent;}.cl-1a8bfd51{margin:0;text-align:left;border-bottom: 0 solid rgba(0, 0, 0, 1.00);border-top: 0 solid rgba(0, 0, 0, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);padding-bottom:10pt;padding-top:5pt;padding-left:5pt;padding-right:5pt;line-height: 1;background-color:transparent;}.cl-1a8bfd5a{margin:0;text-align:right;border-bottom: 0 solid rgba(0, 0, 0, 1.00);border-top: 0 solid rgba(0, 0, 0, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);padding-bottom:10pt;padding-top:5pt;padding-left:5pt;padding-right:5pt;line-height: 1;background-color:transparent;}.cl-1a8c320c{width:1.017in;background-color:transparent;vertical-align: middle;border-bottom: 0 solid rgba(102, 102, 102, 1.00);border-top: 1.5pt solid rgba(102, 102, 102, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-1a8c3216{width:0.911in;background-color:transparent;vertical-align: middle;border-bottom: 0 solid rgba(102, 102, 102, 1.00);border-top: 1.5pt solid rgba(102, 102, 102, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-1a8c3217{width:1.017in;background-color:transparent;vertical-align: middle;border-bottom: 1.5pt solid rgba(102, 102, 102, 1.00);border-top: 0 solid rgba(102, 102, 102, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-1a8c3218{width:0.911in;background-color:transparent;vertical-align: middle;border-bottom: 1.5pt solid rgba(102, 102, 102, 1.00);border-top: 0 solid rgba(102, 102, 102, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-1a8c3220{width:1.017in;background-color:transparent;vertical-align: middle;border-bottom: 1pt solid rgba(153, 153, 153, 1.00);border-top: 0 solid rgba(0, 0, 0, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-1a8c3221{width:0.911in;background-color:transparent;vertical-align: middle;border-bottom: 1pt solid rgba(153, 153, 153, 1.00);border-top: 0 solid rgba(0, 0, 0, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-1a8c3222{width:1.017in;background-color:transparent;vertical-align: middle;border-bottom: 0 solid rgba(0, 0, 0, 1.00);border-top: 1pt solid rgba(153, 153, 153, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-1a8c322a{width:0.911in;background-color:transparent;vertical-align: middle;border-bottom: 0 solid rgba(0, 0, 0, 1.00);border-top: 1pt solid rgba(153, 153, 153, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-1a8c322b{width:1.017in;background-color:transparent;vertical-align: middle;border-bottom: 0 solid rgba(0, 0, 0, 1.00);border-top: 0 solid rgba(0, 0, 0, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-1a8c322c{width:0.911in;background-color:transparent;vertical-align: middle;border-bottom: 0 solid rgba(0, 0, 0, 1.00);border-top: 0 solid rgba(0, 0, 0, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-1a8c3234{width:1.017in;background-color:transparent;vertical-align: middle;border-bottom: 1.5pt solid rgba(102, 102, 102, 1.00);border-top: 0 solid rgba(0, 0, 0, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-1a8c323e{width:0.911in;background-color:transparent;vertical-align: middle;border-bottom: 1.5pt solid rgba(102, 102, 102, 1.00);border-top: 0 solid rgba(0, 0, 0, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-1a8c323f{width:1.017in;background-color:transparent;vertical-align: middle;border-bottom: 1.5pt solid rgba(102, 102, 102, 1.00);border-top: 0 solid rgba(0, 0, 0, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-1a8c3240{width:0.911in;background-color:transparent;vertical-align: middle;border-bottom: 1.5pt solid rgba(102, 102, 102, 1.00);border-top: 0 solid rgba(0, 0, 0, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}</style><table data-quarto-disable-processing='true' class='cl-1aa3cf7a'><thead><tr style="overflow-wrap:break-word;"><th class="cl-1a8c320c"><p class="cl-1a8bfd3c"><span class="cl-1a88aaec">USUBJID</span></p></th><th class="cl-1a8c320c"><p class="cl-1a8bfd3c"><span class="cl-1a88aaec">ARM</span></p></th><th class="cl-1a8c3216"><p class="cl-1a8bfd50"><span class="cl-1a88aaec">VAL</span></p></th></tr><tr style="overflow-wrap:break-word;"><th class="cl-1a8c3217"><p class="cl-1a8bfd3c"><span class="cl-1a88aaf6">character</span></p></th><th class="cl-1a8c3217"><p class="cl-1a8bfd3c"><span class="cl-1a88aaf6">character</span></p></th><th class="cl-1a8c3218"><p class="cl-1a8bfd50"><span class="cl-1a88aaf6">numeric</span></p></th></tr></thead><tbody><tr style="overflow-wrap:break-word;"><td class="cl-1a8c3220"><p class="cl-1a8bfd3c"><span class="cl-1a88aaec">S1</span></p></td><td class="cl-1a8c3220"><p class="cl-1a8bfd3c"><span class="cl-1a88aaec">A</span></p></td><td class="cl-1a8c3221"><p class="cl-1a8bfd50"><span class="cl-1a88aaec">0.3</span></p></td></tr><tr style="overflow-wrap:break-word;"><td class="cl-1a8c3222"><p class="cl-1a8bfd51"><span class="cl-1a88aaec">S1</span></p></td><td class="cl-1a8c3222"><p class="cl-1a8bfd51"><span class="cl-1a88aaec">A</span></p></td><td class="cl-1a8c322a"><p class="cl-1a8bfd5a"><span class="cl-1a88aaec">-2.4</span></p></td></tr><tr style="overflow-wrap:break-word;"><td class="cl-1a8c322b"><p class="cl-1a8bfd3c"><span class="cl-1a88aaec">S1</span></p></td><td class="cl-1a8c322b"><p class="cl-1a8bfd3c"><span class="cl-1a88aaec">B</span></p></td><td class="cl-1a8c322c"><p class="cl-1a8bfd50"><span class="cl-1a88aaec">-0.0</span></p></td></tr><tr style="overflow-wrap:break-word;"><td class="cl-1a8c3220"><p class="cl-1a8bfd3c"><span class="cl-1a88aaec">S2</span></p></td><td class="cl-1a8c3220"><p class="cl-1a8bfd3c"><span class="cl-1a88aaec">A</span></p></td><td class="cl-1a8c3221"><p class="cl-1a8bfd50"><span class="cl-1a88aaec">0.6</span></p></td></tr><tr style="overflow-wrap:break-word;"><td class="cl-1a8c3222"><p class="cl-1a8bfd51"><span class="cl-1a88aaec">S2</span></p></td><td class="cl-1a8c3222"><p class="cl-1a8bfd51"><span class="cl-1a88aaec">A</span></p></td><td class="cl-1a8c322a"><p class="cl-1a8bfd5a"><span class="cl-1a88aaec">1.1</span></p></td></tr><tr style="overflow-wrap:break-word;"><td class="cl-1a8c322b"><p class="cl-1a8bfd3c"><span class="cl-1a88aaec">S2</span></p></td><td class="cl-1a8c322b"><p class="cl-1a8bfd3c"><span class="cl-1a88aaec">B</span></p></td><td class="cl-1a8c322c"><p class="cl-1a8bfd50"><span class="cl-1a88aaec">-1.8</span></p></td></tr><tr style="overflow-wrap:break-word;"><td class="cl-1a8c3234"><p class="cl-1a8bfd3c"><span class="cl-1a88aaec">S3</span></p></td><td class="cl-1a8c3234"><p class="cl-1a8bfd3c"><span class="cl-1a88aaec">A</span></p></td><td class="cl-1a8c323e"><p class="cl-1a8bfd50"><span class="cl-1a88aaec">-0.2</span></p></td></tr></tbody><tfoot><tr style="overflow-wrap:break-word;"><td  colspan="3"class="cl-1a8c323f"><p class="cl-1a8bfd3c"><span class="cl-1a88aaec">n: 7</span></p></td></tr></tfoot></table></div>

# Example: All NA - should return original ft
ft_all_na <- add_flextable_separators(ft, rep(NA, 7))
identical(ft, ft_all_na) # Should be TRUE
#> [1] TRUE

# Example: Invalid character - should throw error
tryCatch(
  add_flextable_separators(ft, c("-", "x", NA, "-", " ", NA, NA)),
  error = function(e) print(e)
)
#> <simpleError in add_flextable_separators(ft, c("-", "x", NA, "-", " ", NA, NA)): Invalid character(s) found as trailing separators: 'x'. Only NA (no sparator), '-' (a line), or ' ' (padding) are allowed.>
```
