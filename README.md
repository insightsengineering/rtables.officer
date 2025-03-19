
<!-- README.md is generated from README.Rmd. Please edit that file -->

# rtables.officer <a href='https://github.com/insightsengineering/rtables'><img src="man/figures/logo.png" align="right" height="200" width="200"/></a>

<!-- start badges -->

[![Check
🛠](https://github.com/insightsengineering/rtables.officer/actions/workflows/check.yaml/badge.svg)](https://github.com/insightsengineering/rtables.officer/actions/workflows/check.yaml)
[![Docs
📚](https://github.com/insightsengineering/rtables.officer/actions/workflows/docs.yaml/badge.svg)](https://insightsengineering.github.io/rtables.officer/)
[![Code Coverage
📔](https://raw.githubusercontent.com/insightsengineering/rtables.officer/_xml_coverage_reports/data/main/badge.svg)](https://raw.githubusercontent.com/insightsengineering/rtables.officer/_xml_coverage_reports/data/main/coverage.xml)

![GitHub
forks](https://img.shields.io/github/forks/insightsengineering/rtables.officer?style=social)
![GitHub repo
stars](https://img.shields.io/github/stars/insightsengineering/rtables.officer?style=social)

![GitHub commit
activity](https://img.shields.io/github/commit-activity/m/insightsengineering/rtables.officer)
![GitHub
contributors](https://img.shields.io/github/contributors/insightsengineering/rtables.officer)
![GitHub last
commit](https://img.shields.io/github/last-commit/insightsengineering/rtables.officer)
![GitHub pull
requests](https://img.shields.io/github/issues-pr/insightsengineering/rtables.officer)
![GitHub repo
size](https://img.shields.io/github/repo-size/insightsengineering/rtables.officer)
![GitHub language
count](https://img.shields.io/github/languages/count/insightsengineering/rtables.officer)
[![Project Status: Active – The project has reached a stable, usable
state and is being actively
developed.](https://www.repostatus.org/badges/latest/active.svg)](https://www.repostatus.org/#active)
[![Open
Issues](https://img.shields.io/github/issues-raw/insightsengineering/rtables.officer?color=red&label=open%20issues)](https://github.com/insightsengineering/rtables.officer/issues?q=is%3Aissue+is%3Aopen+sort%3Aupdated-desc)

[![Current
Version](https://img.shields.io/github/r-package/v/insightsengineering/rtables.officer/main?color=purple&label=Development%20Version)](https://github.com/insightsengineering/rtables.officer/tree/main)
<!-- end badges -->

## Exporting `rtables` to Microsoft Word and beyond

The `rtables.officer` package provides a framework to export tables
created with
[`rtables`](https://github.com/insightsengineering/rtables/) to
Microsoft Word documents. To do so, we use the `officer` package to
create a Word document and the `flextable` package to produce the
intermediary table object that `officer` can use to create the Word
document.

Please refer to the following packages for further information: -
[`rtables`](https://github.com/insightsengineering/rtables/) to create
tables. - [`flextable`](https://github.com/davidgohel/flextable) as an
intermediate html table object. Many aesthetic functionalities are
available at this stage. -
[`officer`](https://github.com/davidgohel/officer) to create Word
documents. Please consider also other exporter options (e.g. `html`)
that are available from `flextable`.

`rtables` and `rtables.officer` is developed and copy written by
`F. Hoffmann-La Roche` and it is released open source under Apache
License Version 2.

## Installation

`rtables.officer` is available on CRAN and you can install the latest
released version with:

``` r
install.packages("rtables.officer")
```

or you can install the latest development version directly from GitHub
with:

``` r
# install.packages("pak")
pak::pak("insightsengineering/rtables.officer")
```

## Example

Here’s a simple example demonstrating how to create a basic table
layout, perform analysis on various columns, and export the resultant
table to a Word document in landscape orientation. Further reading are
available in the
[vignettes](https://insightsengineering.github.io/rtables.officer/latest-tag/articles/).

``` r
# Define the table layout
lyt <- basic_table() %>%
  split_cols_by("ARM") %>%
  analyze(c("AGE", "BMRKR2", "COUNTRY"))

# Build the table
tbl <- build_table(lyt, ex_adsl)

# Export the table to a Word document in landscape orientation
tf <- tempfile(fileext = ".docx")
export_as_docx(tbl,
  file = tf,
  section_properties = section_properties_default(orientation = "landscape")
)

# Expected output (with default theme)
tt_to_flextable(tbl, theme = theme_docx_default())
```

<img src="man/figures/README-unnamed-chunk-2-1.png" width="2000" />

## Contributions

To contribute to this package, please fork the repository, create a
branch, make your changes, and submit a pull request. Your contributions
are greatly appreciated!
