#' Add Conditional Separators (horizontal line or padding) to `flextable` Rows
#'
#' Modifies an existing `flextable` object by adding visual separators
#' (horizontal lines or bottom padding) after specific rows based on a
#' control vector, without changing the table's row count.
#'
#' @param ft A flextable object.
#' @param trailing_sep A vector specifying separators. Its length must equal
#'   the number of rows in the body of `ft`. Allowed values are NA (no
#'   separator), "-" (adds a horizontal line), or " " (adds bottom padding).
#' @param border The `fp_border` object to use for horizontal lines when
#'   `trailing_sep` is "-". Defaults to a gray line of width 1.
#' @param padding The amount of bottom padding (in points) to add when
#'   `trailing_sep` is " ". Defaults to 10.
#'
#' @return The modified flextable object, or the original `ft` if all
#'   `trailing_sep` values are NA. Throws an error for invalid inputs or
#'   invalid characters in `trailing_sep`.
#'
#' @examples
#' content <- data.frame(
#'   USUBJID = c("S1", "S1", "S1", "S2", "S2", "S2", "S3"),
#'   ARM = c("A", "A", "B", "A", "A", "B", "A"),
#'   VAL = round(rnorm(7), 2)
#' )
#' ft <- flextable::as_flextable(content)
#' ft <- flextable::theme_booktabs(ft)
#'
#' # Define separators: line, space, NA, line, space, NA, NA
#' sep_ctrl <- c("-", " ", NA, "-", " ", NA, NA)
#'
#' ft_modified <- add_flextable_separators(ft, sep_ctrl)
#' print(ft_modified)
#'
#' # Example: All NA - should return original ft
#' ft_all_na <- add_flextable_separators(ft, rep(NA, 7))
#' identical(ft, ft_all_na) # Should be TRUE
#'
#' # Example: Invalid character - should throw error
#' tryCatch(
#'   add_flextable_separators(ft, c("-", "x", NA, "-", " ", NA, NA)),
#'   error = function(e) print(e)
#' )
#'
#' @export
add_flextable_separators <- function(ft,
                                     trailing_sep,
                                     border = officer::fp_border(width = 1, color = "grey60"),
                                     padding = 10) {
  # --- Input Validation ---
  if (!inherits(ft, "flextable")) {
    stop("Input 'ft' must be a flextable object.")
  }
  checkmate::assert_character(trailing_sep)

  # Use stats:: explicitly if function is internal and stats not imported
  n_rows_body <- flextable::nrow_part(ft, part = "body")

  # Handle empty flextable case
  if (n_rows_body == 0) {
    if (length(trailing_sep) != 0) {
      warning(
        "Input flextable 'ft' has 0 body rows, but 'trailing_sep' is not empty. ",
        "Returning original empty flextable.",
        call. = FALSE
      )
    }
    return(ft) # Return the empty flextable
  }

  # Check length consistency
  if (length(trailing_sep) != n_rows_body) {
    warning(sprintf(
      "Length of 'trailing_sep' (%d) must equal the number of body rows in 'ft' (%d).",
      length(trailing_sep), n_rows_body
    ))
    return(ft)
  }

  # --- Check for only allowed non-NA values ---
  allowed_chars <- c("-", " ")
  # Use stats:: explicitly if function is internal and stats not imported
  present_chars <- unique(stats::na.omit(trailing_sep))
  invalid_chars <- setdiff(present_chars, allowed_chars)

  if (length(invalid_chars) > 0) {
    stop(
      sprintf(
        paste0(
          "Invalid character(s) found as trailing separators: '%s'. Only NA (no sparator), ",
          "'-' (a line), or ' ' (padding) are allowed."
        ),
        paste(invalid_chars, collapse = ", ") # Quote characters for clarity
      )
    )
  }

  # --- Handle "All NA" case (optimization) ---
  if (all(is.na(trailing_sep))) {
    return(ft)
  }

  # --- Apply Separators ---
  # Loop through each row index of the flextable body
  for (i in seq_len(n_rows_body)) {
    sep_char <- trailing_sep[i]

    # Skip if NA
    if (is.na(sep_char)) {
      next
    }

    # Apply action based on character
    if (sep_char == "-") {
      # Add hline below row i
      ft <- hline(ft, i = i, border = border, part = "body")
    } else if (sep_char == " ") {
      # Add padding below row i
      # Check flextable docs if padding is cumulative; assume it sets for the row.
      ft <- padding(ft, i = i, padding.bottom = padding, part = "body")
    }
    # No 'else' needed because invalid characters were checked earlier
  }

  # --- Return Modified Flextable ---
  ft
}
