test_that("export_as_docx works thanks to tt_to_flextable", {
  lyt <- make_big_lyt()
  tbl <- build_table(lyt, rawdat)
  top_left(tbl) <- "Ethnicity"
  main_title(tbl) <- "Main title"
  subtitles(tbl) <- c("Some Many", "Subtitles")
  main_footer(tbl) <- c("Some Footer", "Mehr")
  prov_footer(tbl) <- "Some prov Footer"
  fnotes_at_path(tbl, rowpath = c("RACE", "BLACK")) <- "factor 2"
  fnotes_at_path(tbl,
    rowpath = c("RACE", "BLACK"),
    colpath = c("ARM", "ARM1", "SEX", "F")
  ) <- "factor 3"

  # Get the flextable
  flex_tbl <- tt_to_flextable(tbl, titles_as_header = TRUE, footers_as_text = FALSE)

  doc_file <- tempfile(fileext = ".docx")

  expect_silent(export_as_docx(tbl,
    file = doc_file, doc_metadata = list("title" = "meh"),
    template_file = doc_file,
    section_properties = section_properties_default()
  ))
  # flx table in input
  expect_silent(export_as_docx(flex_tbl,
    file = doc_file, doc_metadata = list("title" = "meh"),
    template_file = doc_file,
    section_properties = section_properties_default(page_size = "A4")
  ))
  expect_silent(export_as_docx(tbl,
    file = doc_file, doc_metadata = list("title" = "meh"),
    template_file = doc_file,
    section_properties = section_properties_default(orientation = "landscape")
  ))

  expect_true(file.exists(doc_file))
})

test_that("export_as_docx produces a warning if manual column widths are used", {
  skip_if_not_installed("flextable")
  require("flextable", quietly = TRUE)

  lyt <- basic_table() %>%
    split_rows_by("Species") %>%
    analyze("Petal.Length")
  tbl <- build_table(lyt, iris)

  doc_file <- tempfile(fileext = ".docx")

  # Get the flextable
  expect_warning(
    export_as_docx(tbl,
      colwidths = c(1, 2),
      file = doc_file,
      section_properties = section_properties_default()
    ), "The total table width does not match the page width"
  )
})
