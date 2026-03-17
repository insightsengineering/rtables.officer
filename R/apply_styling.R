#' Apply alignments to a flextable
#'
#' @param flx (`flextable`)\cr a `flextable` object to which alignments will be applied.
#' @param aligns_df (`matrix`)\cr a `matrix` object containing the alignments that will be applied.
#' @param part (`character`)\cr the part of flx where the alignments will be applied.
#'   One of: "header", "body" or "footer".
#'
#' @returns a `flextable` object with the alignments updated.
#' @export
#'
#' @seealso [tt_to_flextable()]
#'
#' @examples
#' df <- head(iris)
#' aligns_df <- matrix(data = "right", nrow = nrow(df), ncol = ncol(df))
#' aligns_df[3, 3] <- "center"
#' aligns_df[5, 2] <- "center"
#' flx <- flextable::flextable(df)
#' apply_alignments(flx = flx, aligns_df = aligns_df, part = "body")
apply_alignments <- function(flx, aligns_df, part) {
  checkmate::assert_class(flx, "flextable")
  checkmate::assert_matrix(aligns_df)
  checkmate::assert_choice(part, choices = c("header", "body", "footer"))
  checkmate::assert_true(nrow(aligns_df) == flextable::nrow_part(flx, part))
  checkmate::assert_true(ncol(aligns_df) == ncol(flx[[part]]$dataset))

  # List of characters you want to search for
  search_chars <- unique(c(aligns_df))

  # Loop through each character and find its indexes
  for (char in search_chars) {
    indexes <- which(aligns_df == char, arr.ind = TRUE)
    tmp_inds <- as.data.frame(indexes)
    unique_cols <- unique(tmp_inds$col)
    for (j in unique_cols) {
      unique_rows <- unique(tmp_inds[tmp_inds$col == j, "row"])
      flx <- flx |>
        flextable::align(
          i = unique_rows,
          j = j,
          align = char,
          part = part
        )
    }
  }

  flx
}

#' Apply manual bolding to a flextable
#'
#' @param flx (`flextable`)\cr a `flextable` object to which manual bolding will be applied.
#' @param bold_manual (`list`)\cr a named `list` containing the specification for
#'   the manual bolding, in the format
#'   `list("header" = list("i" = c(), "j" = c()), "body" = list("i" = c(), "j" = c()))`
#'
#' @returns a `flextable` object with the bolding updated.
#' @export
#'
#' @seealso [theme_docx_default()]
#'
#' @examples
#' df <- head(iris)
#' flx <- flextable::flextable(df)
#' special_bold <- list(
#'   "header" = list("i" = 1, "j" = c(1, 3)),
#'   "body" = list("i" = c(1, 2), "j" = 1)
#' )
#' apply_bold_manual(flx, special_bold)
apply_bold_manual <- function(flx, bold_manual) {
  if (is.null(bold_manual)) {
    return(flx)
  }
  checkmate::assert_class(flx, "flextable")
  checkmate::assert_list(bold_manual)
  valid_sections <- c("header", "body") # Only valid values
  checkmate::assert_subset(names(bold_manual), valid_sections)
  for (bi in seq_along(bold_manual)) {
    bld_tmp <- bold_manual[[bi]]
    checkmate::assert_list(bld_tmp)
    if (!all(c("i", "j") %in% names(bld_tmp)) || !all(vapply(bld_tmp, checkmate::test_integerish, logical(1)))) {
      stop(
        "Found an allowed section for manual bold (", names(bold_manual)[bi],
        ") that was not a named list with i (row) and j (col) integer vectors."
      )
    }
    flx <- flextable::bold(flx,
      i = bld_tmp$i, j = bld_tmp$j,
      part = names(bold_manual)[bi]
    )
  }

  flx
}
