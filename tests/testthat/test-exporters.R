context("Exporting to txt, pdf, rtf, and docx")

test_that("export_as_txt works with and without pagination", {
  lyt <- basic_table() %>%
    split_cols_by("ARM") %>%
    analyze(c("AGE", "BMRKR2", "COUNTRY"))

  tbl <- build_table(lyt, ex_adsl)

  tmptxtf <- tempfile()
  export_as_txt(tbl, file = tmptxtf, paginate = TRUE, lpp = 8, verbose = TRUE)
  txtlns <- readLines(tmptxtf)
  expect_identical(
    grep("\\\\s\\\\n", txtlns),
    c(9L, 17L)
  )

  expect_identical(
    toString(tbl),
    export_as_txt(tbl, file = NULL, paginate = FALSE)
  )
})




# test_that("exporting pdfs gives the correct values", {
#     if (check_pdf) {
#         lyt <- basic_table(title = " ") %>%
#             split_rows_by("SEX", page_by = TRUE) %>%
#             analyze("AGE")
#
#         # Building the table
#         tbl <- build_table(lyt, DM)
#
#         tmpf <- tempfile(fileext = ".pdf")
#         res <- export_as_pdf(tbl, file = tmpf, hsep = "=", lpp = 20)
#         res_pdf <- pdf_text(tmpf)
#
#         # Removing spaces and replacing separators
#         res_pdf <- gsub(res_pdf, pattern = "==*", replacement = "+++")
#         res_pdf <- gsub(res_pdf, pattern = "  +", replacement = " ")
#         res_pdf <- gsub(res_pdf, pattern = " \n", replacement = "")
#
#         # Pagination is present as vector in pdf_text. Doing the same with tbl
#         expected <- sapply(paginate_table(tbl), function(x) toString(x, hsep = "="), USE.NAMES = FALSE)
#         names(expected) <- NULL
#
#         # Removing spaces and replacing separators
#         expected <- gsub(expected, pattern = "==*", replacement = "+++")
#         expected <- gsub(expected, pattern = "  +", replacement = " ")
#         expected <- gsub(expected, pattern = " \n", replacement = "\n")
#         expected <- gsub(expected, pattern = "^\n", replacement = "")
#         expect_identical(res_pdf, expected)
#         ## TODO understand better how to compare exactly these outputs
#     }
# })


## https://github.com/insightsengineering/rtables/issues/308
test_that("path_enriched_df works for tables with a column that has all length 1 elements", {
  my_table <- basic_table() %>%
    split_rows_by("Species") %>%
    analyze("Petal.Length") %>%
    build_table(df = iris)
  mydf <- path_enriched_df(my_table)
  expect_identical(dim(mydf), c(3L, 2L))
})


test_that("export_as_doc works thanks to tt_to_flextable", {
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




test_that("export_as_doc produces a warning if manual column widths are used", {
  lyt <- basic_table() %>%
    split_rows_by("Species") %>%
    analyze("Petal.Length")
  tbl <- build_table(lyt, iris)

  doc_file <- tempfile(fileext = ".docx")

  # Get the flextable
  expect_no_warning(
    export_as_docx(tbl,
      colwidths = c(1, 2),
      file = doc_file,
      section_properties = section_properties_default()
    )#, "The total table width does not match the page width"
  )
})
