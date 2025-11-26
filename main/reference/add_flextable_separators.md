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
#> <div class="tabwid"><style>.cl-0e34f29e{}.cl-0e2e11e0{font-family:'DejaVu Sans';font-size:11pt;font-weight:normal;font-style:normal;text-decoration:none;color:rgba(0, 0, 0, 1.00);background-color:transparent;}.cl-0e2e11ea{font-family:'DejaVu Sans';font-size:11pt;font-weight:normal;font-style:normal;text-decoration:none;color:rgba(153, 153, 153, 1.00);background-color:transparent;}.cl-0e30fba8{margin:0;text-align:left;border-bottom: 0 solid rgba(0, 0, 0, 1.00);border-top: 0 solid rgba(0, 0, 0, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);padding-bottom:5pt;padding-top:5pt;padding-left:5pt;padding-right:5pt;line-height: 1;background-color:transparent;}.cl-0e30fbbc{margin:0;text-align:right;border-bottom: 0 solid rgba(0, 0, 0, 1.00);border-top: 0 solid rgba(0, 0, 0, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);padding-bottom:5pt;padding-top:5pt;padding-left:5pt;padding-right:5pt;line-height: 1;background-color:transparent;}.cl-0e30fbbd{margin:0;text-align:left;border-bottom: 0 solid rgba(0, 0, 0, 1.00);border-top: 0 solid rgba(0, 0, 0, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);padding-bottom:10pt;padding-top:5pt;padding-left:5pt;padding-right:5pt;line-height: 1;background-color:transparent;}.cl-0e30fbbe{margin:0;text-align:right;border-bottom: 0 solid rgba(0, 0, 0, 1.00);border-top: 0 solid rgba(0, 0, 0, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);padding-bottom:10pt;padding-top:5pt;padding-left:5pt;padding-right:5pt;line-height: 1;background-color:transparent;}.cl-0e311b2e{width:1.017in;background-color:transparent;vertical-align: middle;border-bottom: 0 solid rgba(102, 102, 102, 1.00);border-top: 1.5pt solid rgba(102, 102, 102, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-0e311b2f{width:0.911in;background-color:transparent;vertical-align: middle;border-bottom: 0 solid rgba(102, 102, 102, 1.00);border-top: 1.5pt solid rgba(102, 102, 102, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-0e311b38{width:1.017in;background-color:transparent;vertical-align: middle;border-bottom: 1.5pt solid rgba(102, 102, 102, 1.00);border-top: 0 solid rgba(102, 102, 102, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-0e311b39{width:0.911in;background-color:transparent;vertical-align: middle;border-bottom: 1.5pt solid rgba(102, 102, 102, 1.00);border-top: 0 solid rgba(102, 102, 102, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-0e311b42{width:1.017in;background-color:transparent;vertical-align: middle;border-bottom: 1pt solid rgba(153, 153, 153, 1.00);border-top: 0 solid rgba(0, 0, 0, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-0e311b43{width:0.911in;background-color:transparent;vertical-align: middle;border-bottom: 1pt solid rgba(153, 153, 153, 1.00);border-top: 0 solid rgba(0, 0, 0, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-0e311b4c{width:1.017in;background-color:transparent;vertical-align: middle;border-bottom: 0 solid rgba(0, 0, 0, 1.00);border-top: 1pt solid rgba(153, 153, 153, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-0e311b4d{width:0.911in;background-color:transparent;vertical-align: middle;border-bottom: 0 solid rgba(0, 0, 0, 1.00);border-top: 1pt solid rgba(153, 153, 153, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-0e311b4e{width:1.017in;background-color:transparent;vertical-align: middle;border-bottom: 0 solid rgba(0, 0, 0, 1.00);border-top: 0 solid rgba(0, 0, 0, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-0e311b56{width:0.911in;background-color:transparent;vertical-align: middle;border-bottom: 0 solid rgba(0, 0, 0, 1.00);border-top: 0 solid rgba(0, 0, 0, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-0e311b57{width:1.017in;background-color:transparent;vertical-align: middle;border-bottom: 1.5pt solid rgba(102, 102, 102, 1.00);border-top: 0 solid rgba(0, 0, 0, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-0e311b58{width:0.911in;background-color:transparent;vertical-align: middle;border-bottom: 1.5pt solid rgba(102, 102, 102, 1.00);border-top: 0 solid rgba(0, 0, 0, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-0e311b60{width:1.017in;background-color:transparent;vertical-align: middle;border-bottom: 1.5pt solid rgba(102, 102, 102, 1.00);border-top: 0 solid rgba(0, 0, 0, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}.cl-0e311b61{width:0.911in;background-color:transparent;vertical-align: middle;border-bottom: 1.5pt solid rgba(102, 102, 102, 1.00);border-top: 0 solid rgba(0, 0, 0, 1.00);border-left: 0 solid rgba(0, 0, 0, 1.00);border-right: 0 solid rgba(0, 0, 0, 1.00);margin-bottom:0;margin-top:0;margin-left:0;margin-right:0;}</style><table data-quarto-disable-processing='true' class='cl-0e34f29e'><thead><tr style="overflow-wrap:break-word;"><th class="cl-0e311b2e"><p class="cl-0e30fba8"><span class="cl-0e2e11e0">USUBJID</span></p></th><th class="cl-0e311b2e"><p class="cl-0e30fba8"><span class="cl-0e2e11e0">ARM</span></p></th><th class="cl-0e311b2f"><p class="cl-0e30fbbc"><span class="cl-0e2e11e0">VAL</span></p></th></tr><tr style="overflow-wrap:break-word;"><th class="cl-0e311b38"><p class="cl-0e30fba8"><span class="cl-0e2e11ea">character</span></p></th><th class="cl-0e311b38"><p class="cl-0e30fba8"><span class="cl-0e2e11ea">character</span></p></th><th class="cl-0e311b39"><p class="cl-0e30fbbc"><span class="cl-0e2e11ea">numeric</span></p></th></tr></thead><tbody><tr style="overflow-wrap:break-word;"><td class="cl-0e311b42"><p class="cl-0e30fba8"><span class="cl-0e2e11e0">S1</span></p></td><td class="cl-0e311b42"><p class="cl-0e30fba8"><span class="cl-0e2e11e0">A</span></p></td><td class="cl-0e311b43"><p class="cl-0e30fbbc"><span class="cl-0e2e11e0">0.3</span></p></td></tr><tr style="overflow-wrap:break-word;"><td class="cl-0e311b4c"><p class="cl-0e30fbbd"><span class="cl-0e2e11e0">S1</span></p></td><td class="cl-0e311b4c"><p class="cl-0e30fbbd"><span class="cl-0e2e11e0">A</span></p></td><td class="cl-0e311b4d"><p class="cl-0e30fbbe"><span class="cl-0e2e11e0">-2.4</span></p></td></tr><tr style="overflow-wrap:break-word;"><td class="cl-0e311b4e"><p class="cl-0e30fba8"><span class="cl-0e2e11e0">S1</span></p></td><td class="cl-0e311b4e"><p class="cl-0e30fba8"><span class="cl-0e2e11e0">B</span></p></td><td class="cl-0e311b56"><p class="cl-0e30fbbc"><span class="cl-0e2e11e0">-0.0</span></p></td></tr><tr style="overflow-wrap:break-word;"><td class="cl-0e311b42"><p class="cl-0e30fba8"><span class="cl-0e2e11e0">S2</span></p></td><td class="cl-0e311b42"><p class="cl-0e30fba8"><span class="cl-0e2e11e0">A</span></p></td><td class="cl-0e311b43"><p class="cl-0e30fbbc"><span class="cl-0e2e11e0">0.6</span></p></td></tr><tr style="overflow-wrap:break-word;"><td class="cl-0e311b4c"><p class="cl-0e30fbbd"><span class="cl-0e2e11e0">S2</span></p></td><td class="cl-0e311b4c"><p class="cl-0e30fbbd"><span class="cl-0e2e11e0">A</span></p></td><td class="cl-0e311b4d"><p class="cl-0e30fbbe"><span class="cl-0e2e11e0">1.1</span></p></td></tr><tr style="overflow-wrap:break-word;"><td class="cl-0e311b4e"><p class="cl-0e30fba8"><span class="cl-0e2e11e0">S2</span></p></td><td class="cl-0e311b4e"><p class="cl-0e30fba8"><span class="cl-0e2e11e0">B</span></p></td><td class="cl-0e311b56"><p class="cl-0e30fbbc"><span class="cl-0e2e11e0">-1.8</span></p></td></tr><tr style="overflow-wrap:break-word;"><td class="cl-0e311b57"><p class="cl-0e30fba8"><span class="cl-0e2e11e0">S3</span></p></td><td class="cl-0e311b57"><p class="cl-0e30fba8"><span class="cl-0e2e11e0">A</span></p></td><td class="cl-0e311b58"><p class="cl-0e30fbbc"><span class="cl-0e2e11e0">-0.2</span></p></td></tr></tbody><tfoot><tr style="overflow-wrap:break-word;"><td  colspan="3"class="cl-0e311b60"><p class="cl-0e30fba8"><span class="cl-0e2e11e0">n: 7</span></p></td></tr></tfoot></table></div>

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
