makefakedat <- function(n = 1000) {
  datadf <- data.frame(
    stringsAsFactors = FALSE,
    ARM = c("ARM1", sample(c("ARM1", "ARM2"), n - 1, replace = TRUE)),
    SEX = c("M", sample(c("M", "F"), n - 1, replace = TRUE)),
    FACTOR2 = c("A", sample(c("A", "B", "C"), n - 1, replace = TRUE)),
    RACE = c("WHITE", sample(c("WHITE", "BLACK"), n - 1, replace = TRUE)),
    AGE = runif(n, 40, 70),
    VAR3 = c("level1", sample(c("level1", "level2"), n - 1,
      replace = TRUE
    ))
  )

  datadf$ethn_label <- c(WHITE = "Caucasian", BLACK = "African American")[datadf$RACE]
  datadf$fac2_label <- paste("Level", datadf$FACTOR2)
  datadf$gend_label <- c(M = "Male", F = "Female")[datadf$SEX]
  datadf
}

set.seed(0)
rawdat <- makefakedat()

make_big_lyt <- function() {
  lyt <- basic_table(show_colcounts = TRUE) %>%
    split_cols_by("ARM") %>%
    ## add nested column split on SEX with value lables from gend_label
    split_cols_by("SEX", "Gender", labels_var = "gend_label") %>%
    ## No row splits have been introduced, so this adds
    ## a root split and puts summary content on it labelled Overall (N)
    ## add_colby_total(label = "All") %>%
    ##    summarize_row_groups(label = "Overall (N)", format = "(N=xx)") %>%
    ## add a new subtable that splits on RACE, value labels from ethn_label
    split_rows_by("RACE", "Ethnicity", labels_var = "ethn_label", label_pos = "hidden") %>%
    summarize_row_groups("RACE", label_fstr = "%s (n)") %>%
    ##
    ## Add nested row split within Race categories for FACTOR2
    ## using a split function that excludes level C
    ## value labels from fac2_label
    split_rows_by("FACTOR2", "Factor2",
      split_fun = remove_split_levels("C"),
      labels_var = "fac2_label",
      label_pos = "hidden"
    ) %>%
    ## Add count summary within FACTOR2 categories
    summarize_row_groups("FACTOR2") %>%
    ## Add analysis/data rows by analyzing AGE variable
    ## Note afun is a function that returns 2 values in a named list
    ## this will create 2 data rows
    analyze(
      "AGE", "Age Analysis",
      afun = function(x) list(mean = mean(x), median = median(x)),
      format = "xx.xx"
    ) %>%
    ## adding more analysis vars "compounds them", placing them at the same
    ## level of nesting as all previous analysis blocks, rather than
    ## attempting to further nest them
    analyze("AGE",
      "Age Analysis redux",
      afun = range,
      format = "xx.x - xx.x",
      table_names = "AgeRedux"
    ) %>%
    ## Note nested=TRUE, this creates a NEW subtable directly under the
    ## root split
    ## afun of table() gives us k count rows, where k is the number of
    ## levels of VAR3, in this case 2.
    analyze("VAR3", "Var3 Counts", afun = list_wrap_x(table), nested = FALSE)
  lyt
}
