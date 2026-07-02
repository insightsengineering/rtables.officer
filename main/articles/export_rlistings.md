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
# Loading required package: rtables
# Loading required package: magrittr
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

| Listing of Adverse Events (First 50 Records) |  |  |  |  |
|----|----|----|----|----|
| Subset: Treatment-Emergent Events |  |  |  |  |
| Protocol: XYZ-123 |  |  |  |  |
| Unique Subject Identifier | Description of Planned Arm | Analysis Toxicity Grade | Dictionary-Derived Term | Severity/Intensity |
| AB12345-BRA-1-id-134 | A: Drug X | 3 | dcd B.2.1.2.1 | MODERATE |
|  |  | 3 | dcd D.1.1.4.2 | MODERATE |
|  |  | 2 | dcd A.1.1.1.2 | MODERATE |
|  |  | 2 | dcd A.1.1.1.2 | MODERATE |
| AB12345-BRA-1-id-141 | C: Combination | 3 | dcd B.2.1.2.1 | MODERATE |
|  |  | 1 | dcd D.2.1.5.3 | MILD |
|  |  | 1 | dcd A.1.1.1.1 | MILD |
|  |  | 2 | dcd A.1.1.1.2 | MODERATE |
|  |  | 1 | dcd A.1.1.1.1 | MILD |
|  |  | 5 | dcd D.1.1.1.1 | SEVERE |
| AB12345-BRA-1-id-236 | B: Placebo | 5 | dcd B.1.1.1.1 | SEVERE |
|  |  | 5 | dcd B.1.1.1.1 | SEVERE |
|  |  | 5 | dcd B.1.1.1.1 | SEVERE |
| AB12345-BRA-1-id-265 | C: Combination | 2 | dcd C.2.1.2.1 | MODERATE |
|  |  | 3 | dcd D.1.1.4.2 | MODERATE |
|  |  | 5 | dcd D.1.1.1.1 | SEVERE |
|  |  | 4 | dcd C.1.1.1.3 | SEVERE |
| AB12345-BRA-1-id-42 | A: Drug X | 2 | dcd C.2.1.2.1 | MODERATE |
|  |  | 5 | dcd D.1.1.1.1 | SEVERE |
|  |  | 2 | dcd C.2.1.2.1 | MODERATE |
|  |  | 2 | dcd A.1.1.1.2 | MODERATE |
|  |  | 1 | dcd B.2.2.3.1 | MILD |
|  |  | 2 | dcd A.1.1.1.2 | MODERATE |
|  |  | 5 | dcd B.1.1.1.1 | SEVERE |
|  |  | 2 | dcd A.1.1.1.2 | MODERATE |
|  |  | 4 | dcd C.1.1.1.3 | SEVERE |
|  |  | 5 | dcd B.1.1.1.1 | SEVERE |
| AB12345-BRA-1-id-65 | B: Placebo | 2 | dcd C.2.1.2.1 | MODERATE |
|  |  | 1 | dcd D.2.1.5.3 | MILD |
|  |  | 5 | dcd B.1.1.1.1 | SEVERE |
|  |  | 4 | dcd C.1.1.1.3 | SEVERE |
| AB12345-BRA-1-id-93 | A: Drug X | 3 | dcd D.1.1.4.2 | MODERATE |
|  |  | 1 | dcd D.2.1.5.3 | MILD |
|  |  | 1 | dcd A.1.1.1.1 | MILD |
|  |  | 1 | dcd D.2.1.5.3 | MILD |
|  |  | 4 | dcd C.1.1.1.3 | SEVERE |
|  |  | 5 | dcd D.1.1.1.1 | SEVERE |
|  |  | 5 | dcd B.1.1.1.1 | SEVERE |
|  |  | 1 | dcd D.2.1.5.3 | MILD |
|  |  | 1 | dcd B.2.2.3.1 | MILD |
| AB12345-BRA-11-id-237 | C: Combination | 1 | dcd D.2.1.5.3 | MILD |
|  |  | 4 | dcd C.1.1.1.3 | SEVERE |
|  |  | 3 | dcd B.2.1.2.1 | MODERATE |
| AB12345-BRA-11-id-321 | C: Combination | 1 | dcd A.1.1.1.1 | MILD |
|  |  | 1 | dcd A.1.1.1.1 | MILD |
|  |  | 2 | dcd C.2.1.2.1 | MODERATE |
|  |  | 2 | dcd A.1.1.1.2 | MODERATE |
|  |  | 1 | dcd D.2.1.5.3 | MILD |
|  |  | 1 | dcd D.2.1.5.3 | MILD |
|  |  | 1 | dcd A.1.1.1.1 | MILD |
| Source: formatters::ex_adae example dataset |  |  |  |  |
| R Version: R version 4.5.2 (2025-10-31) |  |  |  |  |
| rlistings Version: 0.2.13 |  |  |  |  |
| Generated on: 2026-07-02 08:30:39.077785 |  |  |  |  |
| File: your_script_name.R |  |  |  |  |
