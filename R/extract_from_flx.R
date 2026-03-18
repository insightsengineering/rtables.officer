#' Extract font and size from a flextable
#'
#' @param flx (`flextable`)\cr a `flextable` object from which to extract font and size.
#'
#' @returns a named `list` with 2 elements:
#' - ftp: contains font name, size and other text attributes from the title.
#' - ftp_footer: contains font name, size and other text attributes from the footer.
#' @export
#'
#' @seealso [export_as_docx()]
#'
#' @examples
#' df <- head(iris)
#' flx <- flextable::flextable(df)
#' extract_font_and_size_from_flx(flx)
extract_font_and_size_from_flx <- function(flx) {
  checkmate::assert_class(flx, "flextable")
  font_sz_body <- flx$header$styles$text$font.size$data[1, 1]
  font_size_footer <- flx$footer$styles$text$font.size$data
  font_sz_footer <- if (length(font_size_footer) > 0) {
    font_size_footer[1, 1]
  } else {
    font_sz_body - 1
  }
  font_fam <- flx$header$styles$text$font.family$data[1, 1]

  # Set the test as the tt
  fpt <- officer::fp_text(font.family = font_fam, font.size = font_sz_body)
  fpt_footer <- officer::fp_text(font.family = font_fam, font.size = font_sz_footer)

  list("fpt" = fpt, "fpt_footer" = fpt_footer)
}
