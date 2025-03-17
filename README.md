
<!-- README.md is generated from README.Rmd. Please edit that file -->

# rtables <a href='https://github.com/insightsengineering/rtables'><img src="man/figures/logo.png" align="right" height="200" width="200"/></a>

<!-- start badges -->

[![Check
🛠](https://github.com/insightsengineering/rtables/actions/workflows/check.yaml/badge.svg)](https://github.com/insightsengineering/rtables/actions/workflows/check.yaml)
[![Docs
📚](https://github.com/insightsengineering/rtables/actions/workflows/docs.yaml/badge.svg)](https://insightsengineering.github.io/rtables/)
[![Code Coverage
📔](https://raw.githubusercontent.com/insightsengineering/rtables/_xml_coverage_reports/data/main/badge.svg)](https://raw.githubusercontent.com/insightsengineering/rtables/_xml_coverage_reports/data/main/coverage.xml)

![GitHub
forks](https://img.shields.io/github/forks/insightsengineering/rtables?style=social)
![GitHub repo
stars](https://img.shields.io/github/stars/insightsengineering/rtables?style=social)

![GitHub commit
activity](https://img.shields.io/github/commit-activity/m/insightsengineering/rtables)
![GitHub
contributors](https://img.shields.io/github/contributors/insightsengineering/rtables)
![GitHub last
commit](https://img.shields.io/github/last-commit/insightsengineering/rtables)
![GitHub pull
requests](https://img.shields.io/github/issues-pr/insightsengineering/rtables)
![GitHub repo
size](https://img.shields.io/github/repo-size/insightsengineering/rtables)
![GitHub language
count](https://img.shields.io/github/languages/count/insightsengineering/rtables)
[![Project Status: Active – The project has reached a stable, usable
state and is being actively
developed.](https://www.repostatus.org/badges/latest/active.svg)](https://www.repostatus.org/#active)
[![Open
Issues](https://img.shields.io/github/issues-raw/insightsengineering/rtables?color=red&label=open%20issues)](https://github.com/insightsengineering/rtables/issues?q=is%3Aissue+is%3Aopen+sort%3Aupdated-desc)

[![CRAN
Version](https://www.r-pkg.org/badges/version/rtables)](https://CRAN.R-project.org/package=rtables)
[![Current
Version](https://img.shields.io/github/r-package/v/insightsengineering/rtables/main?color=purple&label=Development%20Version)](https://github.com/insightsengineering/rtables/tree/main)
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
