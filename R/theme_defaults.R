# Themes -----------------------------------------------------------------------
#

#' @describeIn tt_to_flextable Main theme function for [export_as_docx()].
#'
#' @param font (`string`)\cr font. Defaults to `"Arial"`. If the font given is not available, the `flextable` default
#'   is used instead. For options, consult the family column from `systemfonts::system_fonts()`.
#' @param font_size (`integer(1)`)\cr font size. Defaults to 9.
#' @param cell_margins (`numeric(1)` or `numeric(4)`)\cr a numeric or a vector of four numbers indicating
#'   `c("left", "right", "top", "bottom")`. It defaults to 0 for top and bottom, and to 0.19 `mm` in Word `pt`
#'   for left and right.
#' @param bold (`character`)\cr parts of the table text that should be in bold. Can be any combination of
#'   `c("header", "content_rows", "label_rows", "top_left")`. The first one renders all column names bold
#'   (not `topleft` content). The second and third option use [formatters::make_row_df()] to render content or/and
#'   label rows as bold.
#' @param bold_manual (named `list` or `NULL`)\cr list of index lists. See example for needed structure. Accepted
#'   groupings/names are `c("header", "body")`.
#' @param border (`flextable::fp_border()`)\cr border style. Defaults to `flextable::fp_border_default(width = 0.5)`.
#'
#' @examples
#' # Example 4: Creating a custom theme -----------------------------------------
#' special_bold <- list(
#'   "header" = list("i" = 1, "j" = c(1, 3)),
#'   "body" = list("i" = c(1, 2), "j" = 1)
#' )
#' custom_theme <- theme_docx_default(
#'   font_size = 10,
#'   font = "Brush Script MT",
#'   border = flextable::fp_border_default(color = "pink", width = 2),
#'   bold = NULL,
#'   bold_manual = special_bold
#' )
#' tt_to_flextable(tbl,
#'   border = flextable::fp_border_default(color = "pink", width = 2),
#'   theme = custom_theme
#' )
#'
#' # Example 5: Extending the docx theme ----------------------------------------
#' my_theme <- function(font_size = 6) { # here can pass additional arguments for default theme
#'   function(flx, ...) {
#'     # First apply theme_docx_default
#'     flx <- theme_docx_default(font_size = font_size)(flx, ...)
#'
#'     # Then apply additional styling
#'     flx <- flextable::border_inner(flx,
#'       part = "body",
#'       border = flextable::fp_border_default(width = 0.5)
#'     )
#'
#'     return(flx)
#'   }
#' }
#' flx <- tt_to_flextable(tbl, theme = my_theme())
#'
#' @export
theme_docx_default <- function(font = "Arial",
                               font_size = 9,
                               cell_margins = c(word_mm_to_pt(1.9), word_mm_to_pt(1.9), 0, 0), # Default in docx
                               bold = c("header", "content_rows", "label_rows", "top_left"),
                               bold_manual = NULL,
                               border = flextable::fp_border_default(width = 0.5)) {
  function(flx, ...) {
    if (!inherits(flx, "flextable")) {
      stop(sprintf(
        "Function `%s` supports only flextable objects.",
        "theme_box()"
      ))
    }
    checkmate::assert_int(font_size, lower = 6, upper = 12)
    checkmate::assert_string(font)
    checkmate::assert_subset(bold,
      eval(formals(theme_docx_default)$bold),
      empty.ok = TRUE
    )
    if (length(cell_margins) == 1) {
      cell_margins <- rep(cell_margins, 4)
    }
    checkmate::assert_numeric(cell_margins, lower = 0, len = 4)

    # Setting values coming from ...
    args <- list(...)
    tbl_row_class <- args$tbl_row_class
    tbl_ncol_body <- flextable::ncol_keys(flx) # tbl_ncol_body respects if rownames = FALSE (only rlistings)

    # Font setting
    flx <- flextable::fontsize(flx, size = font_size, part = "all") %>%
      flextable::fontsize(size = font_size - 1, part = "footer") %>%
      flextable::font(fontname = font, part = "all")

    # Add all borders (very specific fix too)
    flx <- .add_borders(flx, border = border, ncol = tbl_ncol_body)

    # Vertical alignment -> all top for now
    flx <- flx %>%
      flextable::valign(j = seq(2, tbl_ncol_body), valign = "top", part = "body") %>%
      flextable::valign(j = 1, valign = "top", part = "all") %>%
      # topleft styling (-> bottom aligned) xxx merge_at() could merge these, but let's see
      flextable::valign(j = 1, valign = "top", part = "header") %>%
      flextable::valign(j = seq(2, tbl_ncol_body), valign = "top", part = "header")

    flx <- .apply_indentation_and_margin(flx,
      cell_margins = cell_margins, tbl_row_class = tbl_row_class,
      tbl_ncol_body = tbl_ncol_body
    )

    # Vertical padding/spaces - rownames
    if (any(tbl_row_class == "LabelRow")) { # label rows - 3pt top
      flx <- flextable::padding(flx,
        j = 1, i = which(tbl_row_class == "LabelRow"),
        padding.top = 3 + cell_margins[3], padding.bottom = cell_margins[4], part = "body"
      )
    }
    if (any(tbl_row_class == "ContentRow")) { # content rows - 1pt top
      flx <- flextable::padding(flx,
        # j = 1, # removed because I suppose we want alignment with body
        i = which(tbl_row_class == "ContentRow"),
        padding.top = 1 + cell_margins[3], padding.bottom = cell_margins[4], part = "body"
      )
    }
    # single line spacing (for safety) -> space = 1
    flx <- flextable::line_spacing(flx, space = 1, part = "all")

    # Bold settings
    if (any(bold == "header")) {
      flx <- flextable::bold(flx, j = seq(2, tbl_ncol_body), part = "header") # Done with theme
    }
    # Content rows are effectively our labels in row names
    if (any(bold == "content_rows")) {
      if (is.null(tbl_row_class)) {
        stop('bold = "content_rows" needs tbl_row_class = rtables::make_row_df(tt).')
      }
      flx <- flextable::bold(flx, j = 1, i = which(tbl_row_class == "ContentRow"), part = "body")
    }
    if (any(bold == "label_rows")) {
      if (is.null(tbl_row_class)) {
        stop('bold = "content_rows" needs tbl_row_class = rtables::make_row_df(tt).')
      }
      flx <- flextable::bold(flx, j = 1, i = which(tbl_row_class == "LabelRow"), part = "body")
    }
    # topleft information is also bold if content or label rows are bold
    if (any(bold == "top_left")) {
      flx <- flextable::bold(flx, j = 1, part = "header")
    }

    # If you want specific cells to be bold
    flx <- .apply_bold_manual(flx, bold_manual)

    flx
  }
}

#' @describeIn tt_to_flextable Theme function for html outputs.
#' @param remove_internal_borders (`character`)\cr where to remove internal borders between rows. Defaults to
#'   `"label_rows"`. Currently there are no other options and this can be turned off by providing any other character
#'   value.
#'
#' @examples
#' # html theme
#'
#' # Define a layout for the table
#' lyt <- basic_table() %>%
#'   # Split columns by the "ARM" variable
#'   split_cols_by("ARM") %>%
#'   # Analyze the "AGE", "BMRKR2", and "COUNTRY" variables
#'   analyze(c("AGE", "BMRKR2", "COUNTRY"))
#'
#' # Build the table using the defined layout and example data 'ex_adsl'
#' tbl <- build_table(lyt, ex_adsl)
#'
#' # Convert the table to a flextable object suitable for HTML,
#' # applying the default HTML theme and setting the orientation to landscape
#' tbl_html <- tt_to_flextable(
#'   tbl,
#'   theme = theme_html_default(),
#'   section_properties = section_properties_default(orientation = "landscape")
#' )
#'
#' # Save the flextable as an HTML file named "test.html"
#' flextable::save_as_html(tbl_html,
#'   path = tempfile(tmpdir = tempdir(check = TRUE), fileext = ".html")
#' )
#'
#' @export
theme_html_default <- function(font = "Courier",
                               font_size = 9,
                               cell_margins = 0.2,
                               remove_internal_borders = "label_rows",
                               border = flextable::fp_border_default(width = 1, color = "black")) {
  function(flx, ...) {
    if (!inherits(flx, "flextable")) {
      stop(sprintf(
        "Function `%s` supports only flextable objects.",
        "theme_box()"
      ))
    }
    checkmate::assert_int(font_size, lower = 6, upper = 12)
    checkmate::assert_string(font)
    if (length(cell_margins) == 1) {
      cell_margins <- rep(cell_margins, 4)
    }
    checkmate::assert_numeric(cell_margins, lower = 0, len = 4)
    checkmate::assert_character(remove_internal_borders)

    # Setting values coming from ...
    args <- list(...)
    tbl_row_class <- args$tbl_row_class # This is internal info
    nc_body <- flextable::ncol_keys(flx) # respects if rownames = FALSE (only rlistings)
    nr_header <- flextable::nrow_part(flx, "header")

    # Font setting
    flx <- flextable::fontsize(flx, size = font_size, part = "all") %>%
      flextable::fontsize(size = font_size - 1, part = "footer") %>%
      flextable::font(fontname = font, part = "all")

    # all borders
    flx <- .add_borders(flx, border = border, ncol = nc_body)

    if (any(remove_internal_borders == "label_rows") && any(tbl_row_class == "LabelRow")) {
      flx <- flextable::border(flx,
        j = seq(2, nc_body - 1),
        i = which(tbl_row_class == "LabelRow"), part = "body",
        border.left = flextable::fp_border_default(width = 0),
        border.right = flextable::fp_border_default(width = 0)
      ) %>%
        flextable::border(
          j = 1,
          i = which(tbl_row_class == "LabelRow"), part = "body",
          border.right = flextable::fp_border_default(width = 0)
        ) %>%
        flextable::border(
          j = nc_body,
          i = which(tbl_row_class == "LabelRow"), part = "body",
          border.left = flextable::fp_border_default(width = 0)
        )
    }
    flx <- flextable::bg(flx, i = seq_len(nr_header), bg = "grey", part = "header")

    flx
  }
}

# Helper functions -------------------------------------------------------------
#

.add_borders <- function(flx, border, ncol) {
  # all borders
  flx <- flx %>%
    flextable::border_outer(part = "body", border = border) %>%
    flextable::border_outer(part = "header", border = border) %>%
    flextable::border(
      part = "header", j = 1,
      border.left = border,
      border.right = border
    ) %>%
    flextable::border(
      part = "header", j = 1, i = 1,
      border.top = border
    ) %>%
    flextable::border(
      part = "header", j = 1, i = flextable::nrow_part(flx, "header"),
      border.bottom = border
    ) %>%
    flextable::border(
      part = "header", j = seq(2, ncol),
      border.left = border,
      border.right = border
    )

  # Special bottom and top for when there is no empty row
  raw_header <- flx$header$content$data # HACK xxx
  extracted_header <- NULL
  for (ii in seq_len(nrow(raw_header))) {
    extracted_header <- rbind(
      extracted_header,
      sapply(raw_header[ii, ], function(x) x$txt)
    )
  }
  for (ii in seq_len(nrow(extracted_header))) {
    for (jj in seq(2, ncol)) {
      if (extracted_header[ii, jj] != " ") {
        flx <- flextable::border(
          flx,
          part = "header", j = jj, i = ii,
          border.bottom = border
        )
      }
    }
  }

  flx
}

.apply_bold_manual <- function(flx, bold_manual) {
  if (is.null(bold_manual)) {
    return(flx)
  }
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

.apply_indentation_and_margin <- function(flx, cell_margins, tbl_row_class, tbl_ncol_body) {
  flx <- flx %>% # summary/data rows and cells
    flextable::padding(
      padding.top = cell_margins[3],
      padding.bottom = cell_margins[4], part = "body"
    )

  # Horizontal padding all table margin 0.19 mm
  flx <- flextable::padding(flx,
    j = seq(2, tbl_ncol_body),
    padding.left = cell_margins[1],
    padding.right = cell_margins[2]
  )

  # Vertical padding/spaces - header (3pt after)
  flx <- flx %>%
    flextable::padding(
      j = seq(1, tbl_ncol_body), # also topleft
      padding.top = cell_margins[3],
      padding.bottom = cell_margins[4],
      part = "header"
    )

  flx
}

# Remove vertical borders from both sides (for titles)
remove_vborder <- function(flx, part, ii) {
  flx <- flextable::border(flx,
    i = ii, part = part,
    border.left = flextable::fp_border_default(width = 0),
    border.right = flextable::fp_border_default(width = 0)
  )
}


#' @describeIn tt_to_flextable Padding helper functions to transform mm to pt.
#' @param mm (`numeric(1)`)\cr the value in mm to transform to pt.
#'
#' @export
word_mm_to_pt <- function(mm) {
  mm / 0.3527777778
}

# Padding helper functions to transform mm to pt and viceversa
# # General note for word: 1pt -> 0.3527777778mm -> 0.013888888888889"
word_inch_to_pt <- function(inch) { # nocov
  inch / 0.013888888888889 # nocov
}
