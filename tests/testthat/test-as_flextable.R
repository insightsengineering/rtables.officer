test_that("Can create flextable object that works with different styles", {
  analysisfun <- function(x, ...) {
    in_rows(
      row1 = 5,
      row2 = c(1, 2),
      .row_footnotes = list(row1 = "row 1 - row footnote"),
      .cell_footnotes = list(row2 = "row 2 - cell footnote")
    )
  }

  lyt <- basic_table() %>%
    split_cols_by("ARM") %>%
    split_cols_by("SEX", split_fun = keep_split_levels(c("M", "F"))) %>%
    split_rows_by("STRATA1") %>%
    summarize_row_groups() %>%
    split_rows_by("RACE", split_fun = keep_split_levels(c("WHITE", "ASIAN"))) %>%
    analyze("AGE", afun = analysisfun)


  tbl <- build_table(lyt, ex_adsl)
  ft <- tt_to_flextable(tbl, total_page_width = 20)
  expect_equal(sum(unlist(nrow(ft))), 20)

  expect_silent(ft3 <- tt_to_flextable(tbl, theme = NULL))

  # Custom theme
  special_bold <- list(
    "header" = list("i" = c(1, 2), "j" = c(1, 3)),
    "body" = list("i" = c(1, 2), "j" = 1)
  )
  custom_theme <- theme_docx_default(
    font_size = 10,
    font = "Brush Script MT",
    border = officer::fp_border(color = "pink", width = 2),
    bold = NULL,
    bold_manual = special_bold
  )
  expect_silent(tt_to_flextable(tbl, theme = custom_theme))

  # Custom theme error
  special_bold <- list(
    "header" = list("asdai" = c(1, 2), "j" = c(1, 3)),
    "body" = list("i" = c(1, 2), "j" = 1)
  )
  custom_theme <- theme_docx_default(
    font_size = 10,
    font = "Brush Script MT",
    bold = NULL,
    bold_manual = special_bold
  )
  expect_error(tt_to_flextable(tbl, theme = custom_theme), regexp = "header")

  # header colcounts not in a newline works
  topleft_t1 <- topleft_t2 <- basic_table(show_colcounts = TRUE) %>%
    split_rows_by("ARM", label_pos = "topleft") %>%
    split_cols_by("STRATA1")

  topleft_t1 <- topleft_t1 %>%
    analyze("BMRKR1") %>%
    build_table(DM)
  topleft_t1a <- tt_to_flextable(topleft_t1, counts_in_newline = FALSE)
  topleft_t1b <- tt_to_flextable(topleft_t1, counts_in_newline = TRUE)

  topleft_t2 <- topleft_t2 %>%
    split_rows_by("SEX", label_pos = "topleft") %>%
    analyze("BMRKR1") %>%
    build_table(DM) %>%
    tt_to_flextable(counts_in_newline = FALSE)

  expect_equal(flextable::nrow_part(topleft_t2, part = "header"), 2L)
  expect_equal(flextable::nrow_part(topleft_t1a, part = "header"), 1L)
  expect_equal(flextable::nrow_part(topleft_t1b, part = "header"), 1L)
})

test_that("tt_to_flextable does not create different cells when colcounts (or multiple) on different lines", {
  lyt <- basic_table(show_colcounts = TRUE) %>%
    split_rows_by("ARM", label_pos = "topleft") %>%
    split_rows_by("STRATA1", label_pos = "topleft") %>%
    split_cols_by("STRATA1", split_fun = keep_split_levels("B"), show_colcounts = TRUE) %>%
    split_cols_by("SEX", split_fun = keep_split_levels(c("F", "M"))) %>%
    split_cols_by("COUNTRY", split_fun = keep_split_levels("CHN")) %>%
    analyze("AGE")

  tbl <- build_table(lyt, ex_adsl)
  ft1 <- tt_to_flextable(tbl, counts_in_newline = FALSE)
  ft2 <- tt_to_flextable(tbl, counts_in_newline = TRUE)

  expect_equal(flextable::nrow_part(ft1, "header"), flextable::nrow_part(ft2, "header"))
})

test_that("check titles bold and html theme", {
  lyt <- basic_table(show_colcounts = TRUE) %>%
    split_rows_by("ARM", label_pos = "topleft") %>%
    split_rows_by("STRATA1", label_pos = "topleft") %>%
    split_cols_by("STRATA1", split_fun = keep_split_levels("B"), show_colcounts = TRUE) %>%
    split_cols_by("SEX", split_fun = keep_split_levels(c("F", "M"))) %>%
    split_cols_by("COUNTRY", split_fun = keep_split_levels("CHN")) %>%
    analyze("AGE")

  tbl <- build_table(lyt, ex_adsl)
  main_title(tbl) <- "Main title"
  subtitles(tbl) <- c("Some Many", "Subtitles")
  main_footer(tbl) <- c("Some Footer", "Mehr")

  expect_silent(ft1 <- tt_to_flextable(tbl, theme = theme_html_default(), bold_titles = FALSE))
  expect_silent(ft1 <- tt_to_flextable(tbl, theme = theme_html_default(), bold_titles = c(2, 3)))
  expect_error(ft1 <- tt_to_flextable(tbl, theme = theme_html_default(), bold_titles = c(2, 3, 5)))
})

test_that("check pagination", {
  lyt <- basic_table(show_colcounts = TRUE) %>%
    split_rows_by("ARM", label_pos = "topleft", page_by = TRUE) %>%
    split_rows_by("STRATA1", label_pos = "topleft") %>%
    split_cols_by("STRATA1", split_fun = keep_split_levels("B"), show_colcounts = TRUE) %>%
    split_cols_by("SEX", split_fun = keep_split_levels(c("F", "M"))) %>%
    split_cols_by("COUNTRY", split_fun = keep_split_levels("CHN")) %>%
    analyze("AGE")

  tbl <- build_table(lyt, ex_adsl)

  main_title(tbl) <- "Main title"
  subtitles(tbl) <- c("Some Many", "Subtitles")
  main_footer(tbl) <- c("Some Footer", "Mehr")
  prov_footer(tbl) <- "Some prov Footer"

  expect_no_error(out <- tt_to_flextable(tbl, paginate = TRUE, lpp = 100))
  expect_equal(length(out), 3L)
})

test_that("check colwidths in flextable object", {
  lyt <- basic_table(show_colcounts = TRUE) %>%
    split_rows_by("ARM", label_pos = "topleft", page_by = TRUE) %>%
    split_rows_by("STRATA1", label_pos = "topleft") %>%
    split_cols_by("STRATA1", split_fun = keep_split_levels("B"), show_colcounts = TRUE) %>%
    split_cols_by("SEX", split_fun = keep_split_levels(c("F", "M"))) %>%
    split_cols_by("COUNTRY", split_fun = keep_split_levels("CHN")) %>%
    analyze("AGE")

  tbl <- build_table(lyt, ex_adsl)

  main_title(tbl) <- "Main title"
  subtitles(tbl) <- c("Some Many", "Subtitles")
  main_footer(tbl) <- c("Some Footer", "Mehr")
  prov_footer(tbl) <- "Some prov Footer"

  cw <- c(0.9, 0.05, 0.05)
  spd <- section_properties_default(orientation = "landscape")
  fin_cw <- cw * spd$page_size$width / 2 / sum(cw)

  # Fixed total width is / 2
  flx_res <- tt_to_flextable(tbl,
    total_page_width = spd$page_size$width / 2,
    counts_in_newline = TRUE,
    autofit_to_page = TRUE,
    bold_titles = TRUE,
    colwidths = cw
  ) # if you add cw then autofit_to_page = FALSE
  dflx <- dim(flx_res)
  testthat::expect_equal(fin_cw, unname(dflx$widths))
})

test_that("tt_to_flextable works with {rlistings}' objects", {
  lsting <- as_listing(
    df = head(formatters::ex_adae, n = 50),
    key_cols = c("USUBJID", "ARM"),
    disp_cols = c("AETOXGR", "AEDECOD", "AESEV"),
    main_title = "Listing of Adverse Events (First 50 Records)",
    main_footer = "Source: formatters::ex_adae example dataset"
  )

  expect_no_error(out <- tt_to_flextable(lsting))
  expect_equal(flextable::nrow_part(out), nrow(lsting))
  expect_equal(flextable::nrow_part(out, part = "header"), length(all_titles(lsting)) + 1)
})
test_that("tt_to_flextable handles basic rlistings object correctly", {
  # --- Setup ---
  test_data <- head(formatters::ex_adae, n = 50)
  listing_title <- "Listing of Adverse Events (First 50 Records)"
  listing_footer <- "Source: formatters::ex_adae example dataset"
  key_cols <- c("USUBJID", "ARM")
  disp_cols <- c("AETOXGR", "AEDECOD", "AESEV")

  lsting <- as_listing(
    df = test_data,
    key_cols = key_cols,
    disp_cols = disp_cols,
    main_title = listing_title,
    main_footer = listing_footer
  )
  # Add a subtitle for header row count check robustness
  subtitles(lsting) <- "A subtitle for testing"

  expected_titles <- all_titles(lsting) # Main title + subtitle
  expected_col_keys <- c(key_cols, disp_cols)
  expected_data_rows <- nrow(test_data) # 50

  # --- Action ---
  # Use expect_no_error to ensure it runs and capture output
  out <- NULL # Initialize to avoid potential issues if expect_no_error fails early
  expect_no_error(out <- tt_to_flextable(lsting))

  # --- Checks ---
  # Check 1: Output Type
  expect_s3_class(out, "flextable")

  # Check 2: Dimensions
  expect_equal(flextable::nrow_part(out, part = "body"), expected_data_rows,
    label = "Number of body rows should match input data rows (assuming no separator rows added)"
  )

  # Column count (based on keys in flextable)
  expect_equal(length(out$col_keys), length(expected_col_keys),
    label = "Number of columns in flextable"
  )

  # Check 3: Header Structure and Content
  # Number of header rows = titles + column names row
  expect_equal(flextable::nrow_part(out, part = "header"), length(expected_titles) + 1,
    label = "Number of header rows (titles + colnames)"
  )

  # Column keys/names (order matters)
  expect_equal(out$col_keys, expected_col_keys,
    label = "Column keys/names in flextable"
  )

  # Check 4: Footer Content
  # This might need adjustment based on how tt_to_flextable handles footers
  expect_match(out$footer$dataset[[1]], listing_footer,
    fixed = TRUE,
    label = "Flextable footer should contain the listing's main footer text"
  )
})

test_that("tt_to_flextable handles rlistings with active separators", {
  # Create data where the separator column *changes*
  test_data_sep <- data.frame(
    USUBJID = paste0("S", rep(1:2, each = 3)),
    ARM = rep(c("A", "A", "B"), times = 2), # ARM changes within Subject 1 and 2
    AETOXGR = rep(1:3, times = 2),
    AEDECOD = LETTERS[1:6],
    AESEV = rep(c("MILD", "MOD", "SEVERE"), 2)
  )
  # Expected separators: After row 2 (A->B), After row 5 (A->B) = 2 separators

  lsting_sep <- as_listing(
    df = test_data_sep,
    key_cols = c("USUBJID"),
    disp_cols = c("ARM", "AETOXGR", "AEDECOD", "AESEV"),
    add_trailing_sep = "ARM"
  )

  out_sep <- NULL
  expect_no_error(out_sep <- tt_to_flextable(lsting_sep))
  expect_s3_class(out_sep, "flextable")

  # *** Crucial Check: Row count should NOT change ***
  # This assumes tt_to_flextable uses hline or padding, not adding physical rows
  expect_equal(
    flextable::nrow_part(out_sep, part = "body"),
    nrow(test_data_sep), # Should equal original data rows (6)
    label = "Body row count should match original data rows"
  )
})
