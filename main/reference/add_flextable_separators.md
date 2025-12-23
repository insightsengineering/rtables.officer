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
#> <div class="tabwid"><style>.cl-7854eb38{}.cl-783cf35c{font-family:'DejaVu Sans';font-size:11pt;font-weight:normal;font-style:normal;text-decoration:none;color:rgba(0, 0, 0, 1.00);background-color:transparent;}.cl-783cf370{font-family:'DejaVu Sans';font-size:11pt;font-weight:normal;font-style:normal;text-decoration:none;color:rgba(153, 153, 153, 1.00);background-color:transparent;}.cl-783fb236{margin:0;text-align:left;border-bottom: 0 solid rgba(0, 0, 0, 1.00);border-top: 0 solid rgba(0, 0, 0, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);padding-bottom:5pt;padding-top:5pt;padding-left:5pt;padding-right:5pt;line-height: 1;background-color:transparent;}.cl-783fb237{margin:0;text-align:right;border-bottom: 0 solid rgba(0, 0, 0, 1.00);border-top: 0 solid rgba(0, 0, 0, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);padding-bottom:5pt;padding-top:5pt;padding-left:5pt;padding-right:5pt;line-height: 1;background-color:transparent;}.cl-783fb240{margin:0;text-align:left;border-bottom: 0 solid rgba(0, 0, 0, 1.00);border-top: 0 solid rgba(0, 0, 0, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);padding-bottom:10pt;padding-top:5pt;padding-left:5pt;padding-right:5pt;line-height: 1;background-color:transparent;}.cl-783fb241{margin:0;text-align:right;border-bottom: 0 solid rgba(0, 0, 0, 1.00);border-top: 0 solid rgba(0, 0, 0, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);padding-bottom:10pt;padding-top:5pt;padding-left:5pt;padding-right:5pt;line-height: 1;background-color:transparent;}.cl-783fd090{width:1.017in;background-color:transparent;vertical-align: middle;border-bottom: 0 solid rgba(102, 102, 102, 1.00);border-top: 1.5pt solid rgba(102, 102, 102, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-783fd09a{width:0.911in;background-color:transparent;vertical-align: middle;border-bottom: 0 solid rgba(102, 102, 102, 1.00);border-top: 1.5pt solid rgba(102, 102, 102, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-783fd09b{width:1.017in;background-color:transparent;vertical-align: middle;border-bottom: 1.5pt solid rgba(102, 102, 102, 1.00);border-top: 0 solid rgba(102, 102, 102, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-783fd09c{width:0.911in;background-color:transparent;vertical-align: middle;border-bottom: 1.5pt solid rgba(102, 102, 102, 1.00);border-top: 0 solid rgba(102, 102, 102, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-783fd0a4{width:1.017in;background-color:transparent;vertical-align: middle;border-bottom: 1pt solid rgba(153, 153, 153, 1.00);border-top: 0 solid rgba(0, 0, 0, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-783fd0a5{width:0.911in;background-color:transparent;vertical-align: middle;border-bottom: 1pt solid rgba(153, 153, 153, 1.00);border-top: 0 solid rgba(0, 0, 0, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-783fd0a6{width:1.017in;background-color:transparent;vertical-align: middle;border-bottom: 0 solid rgba(0, 0, 0, 1.00);border-top: 1pt solid rgba(153, 153, 153, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-783fd0a7{width:0.911in;background-color:transparent;vertical-align: middle;border-bottom: 0 solid rgba(0, 0, 0, 1.00);border-top: 1pt solid rgba(153, 153, 153, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-783fd0ae{width:1.017in;background-color:transparent;vertical-align: middle;border-bottom: 0 solid rgba(0, 0, 0, 1.00);border-top: 0 solid rgba(0, 0, 0, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-783fd0af{width:0.911in;background-color:transparent;vertical-align: middle;border-bottom: 0 solid rgba(0, 0, 0, 1.00);border-top: 0 solid rgba(0, 0, 0, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-783fd0b0{width:1.017in;background-color:transparent;vertical-align: middle;border-bottom: 1.5pt solid rgba(102, 102, 102, 1.00);border-top: 0 solid rgba(0, 0, 0, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-783fd0b1{width:0.911in;background-color:transparent;vertical-align: middle;border-bottom: 1.5pt solid rgba(102, 102, 102, 1.00);border-top: 0 solid rgba(0, 0, 0, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-783fd0b8{width:1.017in;background-color:transparent;vertical-align: middle;border-bottom: 1.5pt solid rgba(102, 102, 102, 1.00);border-top: 0 solid rgba(0, 0, 0, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-783fd0b9{width:0.911in;background-color:transparent;vertical-align: middle;border-bottom: 1.5pt solid rgba(102, 102, 102, 1.00);border-top: 0 solid rgba(0, 0, 0, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}</style><table data-quarto-disable-processing='true' class='cl-7854eb38'><thead><tr style="overflow-wrap:break-word;"><th class="cl-783fd090"><p class="cl-783fb236"><span class="cl-783cf35c">USUBJID</span></p></th><th class="cl-783fd090"><p class="cl-783fb236"><span class="cl-783cf35c">ARM</span></p></th><th class="cl-783fd09a"><p class="cl-783fb237"><span class="cl-783cf35c">VAL</span></p></th></tr><tr style="overflow-wrap:break-word;"><th class="cl-783fd09b"><p class="cl-783fb236"><span class="cl-783cf370">character</span></p></th><th class="cl-783fd09b"><p class="cl-783fb236"><span class="cl-783cf370">character</span></p></th><th class="cl-783fd09c"><p class="cl-783fb237"><span class="cl-783cf370">numeric</span></p></th></tr></thead><tbody><tr style="overflow-wrap:break-word;"><td class="cl-783fd0a4"><p class="cl-783fb236"><span class="cl-783cf35c">S1</span></p></td><td class="cl-783fd0a4"><p class="cl-783fb236"><span class="cl-783cf35c">A</span></p></td><td class="cl-783fd0a5"><p class="cl-783fb237"><span class="cl-783cf35c">0.3</span></p></td></tr><tr style="overflow-wrap:break-word;"><td class="cl-783fd0a6"><p class="cl-783fb240"><span class="cl-783cf35c">S1</span></p></td><td class="cl-783fd0a6"><p class="cl-783fb240"><span class="cl-783cf35c">A</span></p></td><td class="cl-783fd0a7"><p class="cl-783fb241"><span class="cl-783cf35c">-2.4</span></p></td></tr><tr style="overflow-wrap:break-word;"><td class="cl-783fd0ae"><p class="cl-783fb236"><span class="cl-783cf35c">S1</span></p></td><td class="cl-783fd0ae"><p class="cl-783fb236"><span class="cl-783cf35c">B</span></p></td><td class="cl-783fd0af"><p class="cl-783fb237"><span class="cl-783cf35c">-0.0</span></p></td></tr><tr style="overflow-wrap:break-word;"><td class="cl-783fd0a4"><p class="cl-783fb236"><span class="cl-783cf35c">S2</span></p></td><td class="cl-783fd0a4"><p class="cl-783fb236"><span class="cl-783cf35c">A</span></p></td><td class="cl-783fd0a5"><p class="cl-783fb237"><span class="cl-783cf35c">0.6</span></p></td></tr><tr style="overflow-wrap:break-word;"><td class="cl-783fd0a6"><p class="cl-783fb240"><span class="cl-783cf35c">S2</span></p></td><td class="cl-783fd0a6"><p class="cl-783fb240"><span class="cl-783cf35c">A</span></p></td><td class="cl-783fd0a7"><p class="cl-783fb241"><span class="cl-783cf35c">1.1</span></p></td></tr><tr style="overflow-wrap:break-word;"><td class="cl-783fd0ae"><p class="cl-783fb236"><span class="cl-783cf35c">S2</span></p></td><td class="cl-783fd0ae"><p class="cl-783fb236"><span class="cl-783cf35c">B</span></p></td><td class="cl-783fd0af"><p class="cl-783fb237"><span class="cl-783cf35c">-1.8</span></p></td></tr><tr style="overflow-wrap:break-word;"><td class="cl-783fd0b0"><p class="cl-783fb236"><span class="cl-783cf35c">S3</span></p></td><td class="cl-783fd0b0"><p class="cl-783fb236"><span class="cl-783cf35c">A</span></p></td><td class="cl-783fd0b1"><p class="cl-783fb237"><span class="cl-783cf35c">-0.2</span></p></td></tr></tbody><tfoot><tr style="overflow-wrap:break-word;"><td  colspan="3"class="cl-783fd0b8"><p class="cl-783fb236"><span class="cl-783cf35c">n: 7</span></p></td></tr></tfoot></table></div>

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
