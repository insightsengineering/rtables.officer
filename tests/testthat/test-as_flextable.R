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

  expect_warning(out <- tt_to_flextable(tbl, paginate = TRUE, lpp = 100))
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
