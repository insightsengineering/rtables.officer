# Extract font and size from a flextable

Extract font and size from a flextable

## Usage

``` r
extract_font_and_size_from_flx(flx)
```

## Arguments

- flx:

  (`flextable`)\
  a `flextable` object from which to extract font and size.

## Value

a named `list` with 2 elements:

- ftp: contains font name, size and other text attributes from the
  title.

- ftp_footer: contains font name, size and other text attributes from
  the footer.

## See also

[`export_as_docx()`](https://insightsengineering.github.io/rtables.officer/reference/export_as_docx.md)

## Examples

``` r
df <- head(iris)
flx <- flextable::flextable(df)
extract_font_and_size_from_flx(flx)
#> $fpt
#>   font.size italic  bold underlined strike color     shading    fontname
#> 1        11  FALSE FALSE      FALSE  FALSE black transparent DejaVu Sans
#>   fontname_cs fontname_eastasia fontname.hansi vertical_align
#> 1 DejaVu Sans       DejaVu Sans    DejaVu Sans       baseline
#> 
#> $fpt_footer
#>   font.size italic  bold underlined strike color     shading    fontname
#> 1        10  FALSE FALSE      FALSE  FALSE black transparent DejaVu Sans
#>   fontname_cs fontname_eastasia fontname.hansi vertical_align
#> 1 DejaVu Sans       DejaVu Sans    DejaVu Sans       baseline
#> 
```
