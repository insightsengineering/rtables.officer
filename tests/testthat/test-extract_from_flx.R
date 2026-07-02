test_that("extract_font_and_size_from_flx works as expected", {
  test_data_simple <- data.frame(
    USUBJID = paste0("S", 1:3),
    ARM = c("A", "A", "B"),
    AETOXGR = 1:3,
    AEDECOD = LETTERS[1:3],
    AESEV = c("MILD", "MOD", "SEVERE")
  ) |> formatters::var_relabel(
    USUBJID = "Subject ID",
    ARM = "Treatment Arm",
    AETOXGR = "Toxicity \nGrade", # origin of the issue
    AEDECOD = "Adverse Event"
  )
  lsting <- as_listing(
    df = test_data_simple,
    key_cols = c("USUBJID"),
    disp_cols = c("ARM", "AETOXGR", "AEDECOD", "AESEV"),
    add_trailing_sep = NULL # No separators
  )
  out <- flextable::as_flextable(lsting)

  # basic examples
  expect_no_error(res <- extract_font_and_size_from_flx(out))
  expect_equal(names(res), c("fpt", "fpt_footer"))
  expect_equal(res[["fpt"]], flextable::fp_text_default())
  expect_equal(res[["fpt_footer"]], flextable::fp_text_default())

  out <- tt_to_flextable(lsting)
  expect_no_error(res <- extract_font_and_size_from_flx(out))
  expect_equal(names(res), c("fpt", "fpt_footer"))
  expect_equal(
    res[["fpt"]],
    flextable::fp_text_default(
      font.size = 8,
      font.family = "Arial",
      hansi.family = "Arial",
      eastasia.family = "Arial",
      cs.family = "Arial"
    )
  )
  expect_equal(
    res[["fpt_footer"]],
    flextable::fp_text_default(
      font.size = 7,
      font.family = "Arial",
      hansi.family = "Arial",
      eastasia.family = "Arial",
      cs.family = "Arial"
    )
  )


  # cases with error
  expect_error(
    res <- extract_font_and_size_from_flx(NULL),
    "Assertion on 'flx' failed: Must inherit from class 'flextable', but has class 'NULL'."
  )
})
