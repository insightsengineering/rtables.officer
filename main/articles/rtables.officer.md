# Introduction to {rtables.officer}

Below is a step-by-step guide demonstrating how to use {rtables.officer}
for creating and exporting a clinical trial results table.

1.  Setup and Data Preparation Load the necessary libraries and prepare
    the dataset:

``` r
library(tern)
# Loading required package: rtables
# Loading required package: formatters
# 
# Attaching package: 'formatters'
# The following object is masked from 'package:base':
# 
#     %||%
# Loading required package: magrittr
# 
# Attaching package: 'rtables'
# The following object is masked from 'package:utils':
# 
#     str
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
# Loading required package: rlistings
# Loading required package: tibble

# Load example datasets
adsl <- formatters::ex_adae
adlb <- formatters::ex_adlb

# Convert character variables to factors and handle missing levels
adsl <- df_explicit_na(adsl)
adlb <- df_explicit_na(adlb)

# Create a temporary file for the output
tf <- tempfile(fileext = ".docx")
```

2.  Data Filtering Filter the dataset for specific parameters and
    visits:

``` r
adlb_f <- adlb %>%
  dplyr::filter(
    PARAM %in% c("Alanine Aminotransferase Measurement", "C-Reactive Protein Measurement") &
      !(ACTARM == "B: Placebo" & AVISIT == "WEEK 1 DAY 8") &
      AVISIT != "SCREENING"
  )
```

3.  Define Custom Analysis Function Create a custom function to perform
    the analysis:

``` r
afun <- function(x, .var, .spl_context, ...) {
  n_fun <- sum(!is.na(x), na.rm = TRUE)
  mean_sd_fun <- if (n_fun == 0) c(NA, NA) else c(mean(x, na.rm = TRUE), sd(x, na.rm = TRUE))
  median_fun <- if (n_fun == 0) NA else median(x, na.rm = TRUE)
  min_max_fun <- if (n_fun == 0) c(NA, NA) else c(min(x), max(x))

  is_chg <- .var == "CHG"
  is_baseline <- .spl_context$value[which(.spl_context$split == "AVISIT")] == "BASELINE"
  if (is_baseline && is_chg) n_fun <- mean_sd_fun <- median_fun <- min_max_fun <- NULL

  in_rows(
    "n" = n_fun,
    "Mean (SD)" = mean_sd_fun,
    "Median" = median_fun,
    "Min - Max" = min_max_fun,
    .formats = list("n" = "xx", "Mean (SD)" = "xx.xx (xx.xx)", "Median" = "xx.xx", "Min - Max" = "xx.xx - xx.xx"),
    .format_na_strs = list("n" = "NE", "Mean (SD)" = "NE (NE)", "Median" = "NE", "Min - Max" = "NE - NE")
  )
}
```

4.  Define Table Layout Create the layout for the table:

``` r
lyt <- basic_table() %>%
  split_cols_by("ACTARM", show_colcounts = TRUE, split_fun = keep_split_levels(levels(adlb_f$ACTARM)[c(1, 2)])) %>%
  split_rows_by("PARAM",
    split_fun = drop_split_levels, label_pos = "topleft",
    split_label = obj_label(adlb_f$PARAM), page_by = TRUE
  ) %>%
  split_rows_by("AVISIT",
    split_fun = drop_split_levels, label_pos = "topleft",
    split_label = obj_label(adlb_f$AVISIT)
  ) %>%
  split_cols_by_multivar(
    vars = c("AVAL", "CHG"),
    varlabels = c("Value at Visit", "Change from Baseline")
  ) %>%
  analyze_colvars(afun = afun)
```

5.  Build and Display the Table Build the table using the defined
    layout:

``` r
result <- build_table(lyt, adlb_f)
result
#                                                      A: Drug X                              B: Placebo              
# Parameter                                            (N=1608)                                (N=1340)               
#   Analysis Visit                       Value at Visit   Change from Baseline   Value at Visit   Change from Baseline
# ————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
# Alanine Aminotransferase Measurement                                                                                
#   BASELINE                                                                                                          
#     n                                       134                                     134                             
#     Mean (SD)                           49.57 (8.32)                            50.31 (8.34)                        
#     Median                                 49.57                                   50.15                            
#     Min - Max                          23.97 - 70.85                           26.22 - 79.12                        
#   WEEK 1 DAY 8                                                                                                      
#     n                                       134                 134                  0                   0          
#     Mean (SD)                           48.60 (7.96)       -0.97 (11.70)          NE (NE)             NE (NE)       
#     Median                                 48.42               -1.02                 NE                  NE         
#     Min - Max                          27.71 - 64.64       -25.05 - 39.39         NE - NE             NE - NE       
#   WEEK 2 DAY 15                                                                                                     
#     n                                       134                 134                 134                 134         
#     Mean (SD)                           49.41 (8.47)       -0.16 (12.01)        50.25 (8.53)       -0.06 (12.53)    
#     Median                                 48.34                0.08               49.97               -0.74        
#     Min - Max                          24.34 - 71.06       -26.31 - 34.35      24.38 - 71.07       -36.75 - 28.78   
#   WEEK 3 DAY 22                                                                                                     
#     n                                       134                 134                 134                 134         
#     Mean (SD)                           50.26 (7.51)        0.69 (11.52)        49.67 (7.74)       -0.63 (11.21)    
#     Median                                 50.09                0.90               49.75               -0.36        
#     Min - Max                          33.04 - 68.98       -32.07 - 32.51      33.71 - 66.49       -37.04 - 30.00   
#   WEEK 4 DAY 29                                                                                                     
#     n                                       134                 134                 134                 134         
#     Mean (SD)                           50.70 (9.17)        1.13 (13.01)        49.26 (8.69)       -1.04 (12.59)    
#     Median                                 49.53                0.71               48.30               -2.37        
#     Min - Max                          29.75 - 78.99       -34.00 - 37.59      32.99 - 73.97       -33.16 - 34.47   
#   WEEK 5 DAY 36                                                                                                     
#     n                                       134                 134                 134                 134         
#     Mean (SD)                           50.84 (7.85)        1.26 (12.11)        49.72 (8.45)       -0.59 (12.63)    
#     Median                                 51.44                1.85               50.23                1.46        
#     Min - Max                          31.89 - 70.34       -32.06 - 32.18      30.64 - 68.11       -38.08 - 25.27   
# C-Reactive Protein Measurement                                                                                      
#   BASELINE                                                                                                          
#     n                                       134                                     134                             
#     Mean (SD)                           48.95 (9.44)                            50.08 (7.87)                        
#     Median                                 49.55                                   50.12                            
#     Min - Max                          28.96 - 75.23                           29.91 - 73.35                        
#   WEEK 1 DAY 8                                                                                                      
#     n                                       134                 134                  0                   0          
#     Mean (SD)                           51.86 (8.08)        2.91 (12.48)          NE (NE)             NE (NE)       
#     Median                                 51.23                2.94                 NE                  NE         
#     Min - Max                          28.47 - 72.90       -28.23 - 41.19         NE - NE             NE - NE       
#   WEEK 2 DAY 15                                                                                                     
#     n                                       134                 134                 134                 134         
#     Mean (SD)                           49.75 (8.27)        0.80 (12.55)        50.87 (7.42)        0.79 (10.69)    
#     Median                                 50.44                0.48               51.51                0.82        
#     Min - Max                          29.02 - 69.08       -34.08 - 30.43      30.70 - 67.70       -22.15 - 27.16   
#   WEEK 3 DAY 22                                                                                                     
#     n                                       134                 134                 134                 134         
#     Mean (SD)                           50.06 (8.32)        1.12 (12.36)        49.27 (7.52)       -0.81 (10.62)    
#     Median                                 49.93                3.14               48.78                0.15        
#     Min - Max                          26.70 - 69.37       -36.68 - 31.78      30.70 - 67.76       -30.71 - 29.92   
#   WEEK 4 DAY 29                                                                                                     
#     n                                       134                 134                 134                 134         
#     Mean (SD)                           51.61 (8.13)        2.66 (12.22)        49.47 (8.21)       -0.61 (11.10)    
#     Median                                 52.26                3.17               48.81               -1.27        
#     Min - Max                          31.45 - 71.21       -31.34 - 41.61      27.07 - 67.32       -24.90 - 23.50   
#   WEEK 5 DAY 36                                                                                                     
#     n                                       134                 134                 134                 134         
#     Mean (SD)                           49.71 (8.76)        0.76 (13.36)        50.83 (7.59)        0.75 (11.28)    
#     Median                                 49.22               -0.25               50.25                0.46        
#     Min - Max                          25.97 - 69.62       -39.98 - 33.63      33.32 - 70.27       -25.59 - 33.12
```

Assign titles and footers:

``` r
main_title(result) <- "Alanine Aminotransferase Measurement"
subtitles(result) <- c("This is a subtitle.", "This is another subtitle.")
main_footer(result) <- "This is a demo table for illustration purpose."
prov_footer(result) <- "Program: demo_poc_docx.R\nDate: 2024-11-06\nVersion: 0.0.1\n"
```

6.  Convert to `flextable` and Export to Word Convert the table to a
    `flextable` object and export it to a Word document:

``` r
flx_res <- tt_to_flextable(result)
export_as_docx(flx_res,
  file = tf,
  section_properties = section_properties_default(orientation = "landscape")
)
flx_res
```

[TABLE]

## Advanced Customizations

You can further customize your tables, such as setting column widths,
handling pagination, and more.

### Column Widths

``` r
cw <- propose_column_widths(result)
cw <- cw / sum(cw)
cw <- c(0.6, 0.1, 0.1, 0.1, 0.1)
spd <- section_properties_default(orientation = "landscape")
fin_cw <- cw * spd$page_size$width / 2 / sum(cw)

flex_tbl <- tt_to_flextable(result,
  total_page_width = spd$page_size$width / 2,
  counts_in_newline = TRUE,
  autofit_to_page = FALSE,
  bold_titles = TRUE,
  colwidths = cw
)

export_as_docx(flex_tbl, file = tf)
# Warning in FUN(X[[i]], ...): The total table width does not match the page
# width. The column widths will be resized to fit the page. Please consider
# modifying the parameter total_page_width in tt_to_flextable().
flex_tbl
```

[TABLE]

### Pagination

``` r
flx_res <- tt_to_flextable(
  result,
  paginate = TRUE,
  titles_as_header = FALSE,
  lpp = 250,
  counts_in_newline = TRUE,
  bold_titles = TRUE,
  theme = theme_docx_default()
)
export_as_docx(flx_res, file = tf, add_page_break = TRUE)
flx_res[[1]]
```

[TABLE]

### Horizontal separators (note `section_div = <chr>`)

``` r
tbl <- basic_table() %>%
  split_cols_by("ACTARM") %>%
  split_rows_by("PARAM", split_fun = drop_split_levels, section_div = "-") %>%
  split_rows_by("AVISIT", split_fun = drop_split_levels, section_div = " ") %>%
  split_cols_by_multivar(
    vars = c("AVAL", "CHG"), varlabels = c("Value at Visit", "Change from Baseline")
  ) %>%
  analyze_colvars(afun = afun) %>%
  build_table(adlb_f)
flx_res <- tt_to_flextable(tbl)
export_as_docx(flx_res, file = tf)
flx_res
```

[TABLE]
