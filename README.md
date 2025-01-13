
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

## Reporting Tables with R

The `rtables` R package was designed to create and display complex
tables with R. The cells in an `rtable` may contain any high-dimensional
data structure which can then be displayed with cell-specific formatting
instructions. Currently, `rtables` can export tables in `ascii` `html`,
and `pdf` formats.

The `rtables.officer` package is designed to support export formats
related to the Microsoft Office software suite, including Microsoft Word
(`docx`) and Microsoft PowerPoint (`pptx`).

`rtables` and `rtables.officer` are developed and copy written by
`F. Hoffmann-La Roche` and are released open source under Apache License
Version 2.

## Installation

You can install the latest development version of `rtables.officer`
directly from GitHub with:

``` r
# install.packages("pak")
pak::pak("insightsengineering/rtables.officer")
```
