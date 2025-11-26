# Exporting listings

This vignette shows how to use {rtables.officer} to export clinical
trial results listing created with {rlistings}.

1.  Setup and Data Preparation Load the necessary libraries and prepare
    the dataset:

``` r
library(rlistings)
# Loading required package: formatters
# 
# Attaching package: 'formatters'
# The following object is masked from 'package:base':
# 
#     %||%
# Loading required package: tibble
library(dplyr)
# 
# Attaching package: 'dplyr'
# The following objects are masked from 'package:stats':
# 
#     filter, lag
# The following objects are masked from 'package:base':
# 
#     intersect, setdiff, setequal, union
library(rtables.officer)
# Loading required package: magrittr
# Loading required package: rtables
# 
# Attaching package: 'rtables'
# The following object is masked from 'package:utils':
# 
#     str

lsting <- as_listing(
  df = head(formatters::ex_adae, n = 50),
  key_cols = c("USUBJID", "ARM"),
  disp_cols = c("AETOXGR", "AEDECOD", "AESEV"),
  main_title = "Listing of Adverse Events (First 50 Records)",
  main_footer = "Source: formatters::ex_adae example dataset",
  add_trailing_sep = "ARM" # for readability adds a space line between differen ARMs
)
```

1b. Decorations (optional)

``` r
# 1. Add Subtitles and Provenance Footer
subtitles(lsting) <- c(
  "Subset: Treatment-Emergent Events",
  "Protocol: XYZ-123"
)
prov_footer(lsting) <- c(
  paste("R Version:", R.version.string),
  paste("rlistings Version:", packageVersion("rlistings")),
  paste("Generated on:", Sys.time()), # Use current time
  "File: your_script_name.R" # Add script name if applicable
)
```

2.  Convert to `flextable` and Export to Word Convert the table to a
    `flextable` object and export it to a Word document:

``` r
flx_res <- tt_to_flextable(lsting)

# Create a temporary file for the output
tf <- tempfile(fileext = ".docx")

export_as_docx(lsting,
  file = tf,
  section_properties = section_properties_default(orientation = "landscape")
)
flx_res
```

[TABLE]
