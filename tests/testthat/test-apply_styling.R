test_that("apply_alignments works as expected", {

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
  out <- tt_to_flextable(lsting)

  alignments_header_before <- out$header$styles$pars$text.align$data
  alignments_body_before <- out$body$styles$pars$text.align$data
  alignments_footer_before <- out$footer$styles$pars$text.align$data

  # update alignments in "body"
  aligns_df <- matrix(data = "left", nrow = nrow(test_data_simple), ncol = ncol(test_data_simple))
  aligns_df[1, 1] <- "right"
  aligns_df[3, 3] <- "right"
  expected_alignments <- aligns_df
  colnames(expected_alignments) <- colnames(test_data_simple)
  expect_no_error(res <- apply_alignments(flx = out, aligns_df = aligns_df, part = "body"))
  alignments_body_after <- res$body$styles$pars$text.align$data
  expect_equal(expected_alignments, alignments_body_after)
  alignments_header_after <- res$header$styles$pars$text.align$data
  alignments_footer_after <- res$footer$styles$pars$text.align$data
  expect_equal(alignments_header_before, alignments_header_after)
  expect_equal(alignments_footer_before, alignments_footer_after)


  # update alignments in another part, e.g. in "header"
  aligns_df <- matrix(data = "left", nrow = flextable::nrow_part(out, "header"), ncol = ncol(test_data_simple))
  aligns_df[1, 3] <- "right"
  expected_alignments <- aligns_df
  colnames(expected_alignments) <- colnames(test_data_simple)
  expect_no_error(res <- apply_alignments(flx = out, aligns_df = aligns_df, part = "header"))
  alignments_header_after <- res$header$styles$pars$text.align$data
  expect_equal(expected_alignments, alignments_header_after)
  alignments_body_after <- res$body$styles$pars$text.align$data
  alignments_footer_after <- res$footer$styles$pars$text.align$data
  expect_equal(alignments_body_before, alignments_body_after)
  expect_equal(alignments_footer_before, alignments_footer_after)


  # cases with error
  expect_error(res <- apply_alignments(flx = NULL, aligns_df = aligns_df, part = "body"),
               "Assertion on 'flx' failed: Must inherit from class 'flextable', but has class 'NULL'.")

  expect_error(res <- apply_alignments(flx = out, aligns_df = aligns_df, part = "another part"),
               "Must be element of set \\{'header','body','footer'\\}, but is 'another part'.")

  aligns_df_2 <- matrix(data = "left", nrow = nrow(test_data_simple) + 1, ncol = ncol(test_data_simple))
  expect_error(res <- apply_alignments(flx = out, aligns_df = aligns_df_2, part = "body"),
               "Assertion on 'nrow\\(aligns_df\\) == flextable::nrow_part\\(flx, part\\)' failed: Must be TRUE.")

})

test_that("apply_bold_manual works as expected", {

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
  manual_bold <- list(
    "header" = list("i" = 1, "j" = c(1, 3)),
    "body" = list("i" = c(1, 2), "j" = 1)
  )
  res <- apply_bold_manual(out, manual_bold)

  expected_res <- matrix(FALSE, nrow = flextable::nrow_part(out, "header"), ncol = ncol(test_data_simple))
  colnames(expected_res) <- colnames(test_data_simple)
  expected_res[1, 1] <- TRUE
  expected_res[1, 3] <- TRUE
  actual_res <- res$header$styles$text$bold$data
  expect_equal(actual_res, expected_res)

  expected_res <- matrix(FALSE, nrow = flextable::nrow_part(out, "body"), ncol = ncol(test_data_simple))
  colnames(expected_res) <- colnames(test_data_simple)
  expected_res[1, 1] <- TRUE
  expected_res[2, 1] <- TRUE
  actual_res <- res$body$styles$text$bold$data
  expect_equal(actual_res, expected_res)

  # cases with error
  expect_error(res <- apply_bold_manual(NULL, manual_bold),
               "Assertion on 'flx' failed: Must inherit from class 'flextable', but has class 'NULL'.")

  manual_bold <- list(
    "header" = list("i" = 1, "j" = c(1, 6)),
    "body" = list("i" = c(1, 2), "j" = 1)
  )
  expect_error(res <- apply_bold_manual(out, manual_bold),
               "invalid columns selection")

  manual_bold <- list(
    "footer" = list("i" = 1, "j" = c(1, 3)),
    "body" = list("i" = c(1, 2), "j" = 1)
  )
  expect_error(res <- apply_bold_manual(out, manual_bold),
               "Assertion on 'names\\(bold_manual\\)' failed: Must be a subset of \\{'header','body'\\}, but has additional elements \\{'footer'\\}.")

})


