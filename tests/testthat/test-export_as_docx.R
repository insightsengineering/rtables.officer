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
  flex_tbl <- tt_to_flextable(tbl, titles_as_header = TRUE, integrate_footers = TRUE)

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

test_that("export_as_docx works thanks to tt_to_flextable", {
  lsting <- as_listing(
    df = head(formatters::ex_adae, n = 50),
    key_cols = c("USUBJID", "ARM"),
    disp_cols = c("AETOXGR", "AEDECOD", "AESEV"),
    main_title = "Listing of Adverse Events (First 50 Records)",
    main_footer = "Source: formatters::ex_adae example dataset",
    add_trailing_sep = "ARM"
  )

  doc_file <- tempfile(fileext = ".docx")
  expect_no_error(
    out <- export_as_docx(lsting, doc_file, titles_as_header = TRUE, integrate_footers = TRUE)
  )
})


test_that("Getting correct template file", {
  root <- system.file(package = "rtables.officer")
  expect_equal(
    get_template_file(section_properties_default(page_size = "A4", orientation = "portrait")),
    file.path(root, "templates/a4_portrait.docx")
  )

  expect_equal(
    get_template_file(section_properties_default(page_size = "A4", orientation = "landscape")),
    file.path(root, "templates/a4_landscape.docx")
  )

  expect_equal(
    get_template_file(section_properties_default(page_size = "letter", orientation = "portrait")),
    file.path(root, "templates/letter_portrait.docx")
  )

  expect_equal(
    get_template_file(section_properties_default(page_size = "letter", orientation = "landscape")),
    file.path(root, "templates/letter_landscape.docx")
  )
})
