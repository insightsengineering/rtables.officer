# Flextable conversion ---------------------------------------------------------
#

#' Create a `flextable` from an `rtables` table
#'
#' Principally used within [export_as_docx()], this function produces a `flextable` from an `rtables` table.
#' If `theme = theme_docx_default()` (default), a `.docx`-friendly table will be produced.
#' If `theme = NULL`, the table will be produced in an `rtables`-like style.
#'
#' @inheritParams rtables::gen_args
#' @inheritParams rtables::paginate_table
#' @param theme (`function` or `NULL`)\cr a theme function designed to change the layout and style of a `flextable`
#'   object. Defaults to `theme_docx_default()`, the classic Microsoft Word output style. If `NULL`, a table with style
#'   similar to the `rtables` default will be produced. See Details below for more information.
#' @param indent_size (`numeric(1)`)\cr indentation size. If `NULL`, the default indent size of the table (see
#'   [formatters::matrix_form()] `indent_size`, default is 2) is used. To work with `docx`, any size is multiplied by
#'   1 mm (2.83 pt) by default.
#' @param titles_as_header (`flag`)\cr whether the table should be self-contained and additional header rows created for
#'   the [formatters::main_title()] string and [formatters::subtitles()] character vector (one row per element).
#'   Defaults to `TRUE`. If `FALSE`, titles and subtitles are added as a paragraph of text above the table.
#' @param bold_titles (`flag` or `integer`)\cr whether titles should be bold (defaults to `TRUE`). If one or more
#'   integers are provided, these integers are used as indices for lines at which titles should be bold.
#' @param footers_as_text (`flag`)\cr whether footers should be added as a new paragraph after the table (`TRUE`) or
#'   the table should be self-contained, implementing `flextable`-style footnotes (`FALSE`) with the same style but a
#'   smaller font. Defaults to `FALSE`.
#' @param counts_in_newline (`flag`)\cr whether column counts should be printed on a new line. In `rtables`, column
#'   counts (i.e. `(N=xx)`) are always printed on a new line (`TRUE`). For `docx` exports it may be preferred to print
#'   these counts on the same line (`FALSE`). Defaults to `FALSE`.
#' @param paginate (`flag`)\cr whether the `rtables` pagination mechanism should be used. If `TRUE`, this option splits
#'   `tt` into multiple `flextables` as different "pages". When using [export_as_docx()] we suggest setting this to
#'   `FALSE` and relying only on the default Microsoft Word pagination system as co-operation between the two mechanisms
#'   is not guaranteed. Defaults to `FALSE`.
#' @param total_page_width (`numeric(1)`)\cr total page width (in inches) for the resulting flextable(s). Any values
#'   added for column widths are normalized by the total page width. Defaults to 10. If `autofit_to_page = TRUE`, this
#'   value is automatically set to the allowed page width.
#' @param total_page_height (`numeric(1)`)\cr total page height (in inches) for the resulting flextable(s). Used only
#'   to estimate number of lines per page (`lpp`) when `paginate = TRUE`. Defaults to 10.
#' @param colwidths (`numeric`)\cr column widths for the resulting flextable(s). If `NULL`, the column widths estimated
#'   with [formatters::propose_column_widths()] will be used. When exporting into `.docx` these values are normalized
#'   to represent a fraction of the `total_page_width`. If these are specified, `autofit_to_page` is set to `FALSE`.
#' @param autofit_to_page (`flag`)\cr whether column widths should be automatically adjusted to fit the total page
#'   width. If `FALSE`, `colwidths` is used to indicate proportions of `total_page_width`. Defaults to `TRUE`.
#'   See `flextable::set_table_properties(layout)` for more details.
#' @param ... (`any`)\cr additional parameters to be passed to the pagination function. See [rtables::paginate_table()]
#'   for options. If `paginate = FALSE` this argument is ignored.
#'
#' @return A `flextable` object.
#'
#' @note
#' Currently `cpp`, `tf_wrap`, and `max_width` are only used in pagination and should be used cautiously if used in
#' combination with `colwidths` and `autofit_to_page`. If issues arise, please raise an issue on GitHub or communicate
#' this to the package maintainers directly.
#'
#' @details
#' If you would like to make a minor change to a pre-existing style, this can be done by extending themes. You can do
#' this by either adding your own theme to the theme call (e.g. `theme = c(theme_docx_default(), my_theme)`) or creating
#' a new theme as shown in the examples below. Please pay close attention to the parameters' inputs.
#'
#' It is possible to use some hidden values to build your own theme (hence the need for the `...` parameter). In
#' particular, [tt_to_flextable()] uses the following variable: `tbl_row_class = rtables::make_row_df(tt)$node_class`.
#' This is ignored if not used in the theme. See [theme_docx_default()] for an example on how to retrieve and use these
#' values.
#'
#' @seealso [export_as_docx()]
#'
#' @examples
#' analysisfun <- function(x, ...) {
#'   in_rows(
#'     row1 = 5,
#'     row2 = c(1, 2),
#'     .row_footnotes = list(row1 = "row 1 - row footnote"),
#'     .cell_footnotes = list(row2 = "row 2 - cell footnote")
#'   )
#' }
#'
#' lyt <- basic_table(
#'   title = "Title says Whaaaat", subtitles = "Oh, ok.",
#'   main_footer = "ha HA! Footer!"
#' ) %>%
#'   split_cols_by("ARM") %>%
#'   analyze("AGE", afun = analysisfun)
#'
#' tbl <- build_table(lyt, ex_adsl)
#'
#' # Example 1: rtables style ---------------------------------------------------
#' tt_to_flextable(tbl, theme = NULL)
#'
#' # Example 2: docx style ------------------------------------------------------
#' tt_to_flextable(tbl, theme = theme_docx_default(font_size = 6))
#'
#' # Example 3: Extending the docx theme ----------------------------------------
#' my_theme <- function(x, ...) {
#'   flextable::border_inner(x, part = "body", border = flextable::fp_border_default(width = 0.5))
#' }
#' flx <- tt_to_flextable(tbl, theme = c(theme_docx_default(), my_theme))
#'
#' @export
tt_to_flextable <- function(tt,
                            theme = theme_docx_default(),
                            border = flextable::fp_border_default(width = 0.5),
                            indent_size = NULL,
                            titles_as_header = TRUE,
                            bold_titles = TRUE,
                            footers_as_text = FALSE,
                            counts_in_newline = FALSE,
                            paginate = FALSE,
                            fontspec = NULL,
                            lpp = NULL,
                            cpp = NULL,
                            ...,
                            colwidths = NULL,
                            tf_wrap = !is.null(cpp),
                            max_width = cpp,
                            total_page_height = 10, # portrait 11 landscape 8.5
                            total_page_width = 10, # portrait 8.5 landscape 11
                            autofit_to_page = TRUE) {
  if (!inherits(tt, "VTableTree")) {
    stop("Input table is not an rtables' object.")
  }
  checkmate::assert_flag(titles_as_header)
  checkmate::assert_flag(footers_as_text)
  checkmate::assert_flag(counts_in_newline)
  checkmate::assert_flag(autofit_to_page)
  checkmate::assert_number(total_page_width, lower = 1)
  checkmate::assert_number(total_page_height, lower = 1)
  checkmate::assert_numeric(colwidths, lower = 0, len = ncol(tt) + 1, null.ok = TRUE)
  if (!is.null(colwidths)) {
    autofit_to_page <- FALSE
  }

  left_right_fixed_margins <- word_mm_to_pt(1.9)

  ## if we're paginating, just call -> pagination happens also afterwards if needed
  if (paginate) {
    # Lets find out the row heights in inches with flextable
    # Capture all current arguments in a list
    args <- as.list(environment())

    # Modify the 'paginate' argument
    args$paginate <- FALSE

    # Use do.call to call the same function with modified arguments
    tmp_flx <- do.call(tt_to_flextable, args)

    # Determine line per pages (lpp) expected from heights of rows (in inches)
    row_heights <- dim(tmp_flx)$heights
    nr_header <- flextable::nrow_part(tmp_flx, part = "header")
    nr_body <- flextable::nrow_part(tmp_flx, part = "body")
    nr_footer <- flextable::nrow_part(tmp_flx, part = "footer")
    if (sum(nr_header, nr_body, nr_footer) != length(row_heights)) {
      stop("Something went wrong with the row heights. Maybe \\n? Contact maintener.") # nocov
    }
    rh_df <- data.frame(rh = row_heights, part = c(
      rep("header", nr_header), rep("body", nr_body), rep("footer", nr_footer)
    ))
    needed_height_header_footer <- sum(rh_df$rh[rh_df$part %in% c("header", "footer")])
    starting_lpp <- nr_header + nr_footer
    cumsum_page_heights <- needed_height_header_footer + cumsum(rh_df$rh[rh_df$part == "body"])
    expected_lpp <- starting_lpp + max(which(cumsum_page_heights < total_page_height))
    if (is.null(lpp)) {
      lpp <- expected_lpp
    } else if (expected_lpp < lpp) {
      # lpp needs to be estimated along with cpp if not provided
      warning(
        "lpp is too large for the given total_page_height. Change the parameters or",
        " each table will be too long to fit each page."
      )
    }
    tabs <- rtables::paginate_table(tt,
      fontspec = fontspec,
      lpp = lpp,
      cpp = cpp, tf_wrap = tf_wrap, max_width = max_width, # This can only be trial an error
      ...
    )
    cinds <- lapply(tabs, function(tb) c(1, .figure_out_colinds(tb, tt) + 1L))
    args$colwidths <- NULL
    args$tt <- NULL
    cl <- if (!is.null(colwidths)) {
      lapply(cinds, function(ci) colwidths[ci])
    } else {
      lapply(cinds, function(ci) {
        return(NULL)
      })
    }
    return(mapply(tt_to_flextable,
      tt = tabs, colwidths = cl,
      MoreArgs = args,
      SIMPLIFY = FALSE
    ))
  }

  # Extract relevant information
  matform <- rtables::matrix_form(tt, fontspec = fontspec, indent_rownames = FALSE)
  body <- formatters::mf_strings(matform) # Contains header
  spans <- formatters::mf_spans(matform) # Contains header
  mpf_aligns <- formatters::mf_aligns(matform) # Contains header
  hnum <- formatters::mf_nlheader(matform) # Number of lines for the header
  rdf <- rtables::make_row_df(tt) # Row-wise info

  # decimal alignment pre-proc
  if (any(grepl("dec", mpf_aligns))) {
    body <- formatters::decimal_align(body, mpf_aligns)
    # Coercion for flextable
    mpf_aligns[mpf_aligns == "decimal"] <- "center"
    mpf_aligns[mpf_aligns == "dec_left"] <- "left"
    mpf_aligns[mpf_aligns == "dec_right"] <- "right"
  }

  # Fundamental content of the table
  content <- as.data.frame(body[-seq_len(hnum), , drop = FALSE])

  # Fix for empty strings -> they used to get wrong font and size
  content[content == ""] <- " "

  flx <- flextable::qflextable(content) %>%
    # Default rtables if no footnotes
    .remove_hborder(part = "body", w = "bottom")

  # Header addition -> NB: here we have a problem with (N=xx)
  hdr <- body[seq_len(hnum), , drop = FALSE]

  # Change of (N=xx) behavior as we need it in the same cell, even if on diff lines
  if (hnum > 1) { # otherwise nothing to do
    det_nclab <- apply(hdr, 2, grepl, pattern = "\\(N=[0-9]+\\)$")
    has_nclab <- apply(det_nclab, 1, any) # vector of rows with (N=xx)
    whsnc <- which(has_nclab) # which rows have it
    if (any(has_nclab)) {
      for (i in seq_along(whsnc)) {
        wi <- whsnc[i]
        what_is_nclab <- det_nclab[wi, ] # extract detected row

        colcounts_split_chr <- if (isFALSE(counts_in_newline)) {
          " "
        } else {
          "\n"
        }

        # condition for popping the interested row by merging the upper one
        hdr[wi, what_is_nclab] <- paste(hdr[wi - 1, what_is_nclab],
          hdr[wi, what_is_nclab],
          sep = colcounts_split_chr
        )
        hdr[wi - 1, what_is_nclab] <- ""

        # Removing unused rows if necessary
        row_to_pop <- wi - 1

        # Case where topleft is not empty, we reconstruct the header pushing empty up
        what_to_put_up <- hdr[row_to_pop, what_is_nclab, drop = FALSE]
        if (all(!nzchar(what_to_put_up)) && row_to_pop > 1) {
          reconstructed_hdr <- rbind(
            cbind(
              hdr[seq(row_to_pop), !what_is_nclab],
              rbind(
                what_to_put_up,
                hdr[seq(row_to_pop - 1), what_is_nclab]
              )
            ),
            hdr[seq(row_to_pop + 1, nrow(hdr)), ]
          )
          row_to_pop <- 1
          hdr <- reconstructed_hdr
        }

        # We can remove the row if they are all ""
        if (all(!nzchar(hdr[row_to_pop, ]))) {
          hdr <- hdr[-row_to_pop, , drop = FALSE]
          spans <- spans[-row_to_pop, , drop = FALSE]
          body <- body[-row_to_pop, , drop = FALSE]
          mpf_aligns <- mpf_aligns[-row_to_pop, , drop = FALSE]
          hnum <- hnum - 1
          # for multiple lines
          whsnc <- whsnc - 1
          det_nclab <- det_nclab[-row_to_pop, , drop = FALSE]
        }
      }
    }
  }

  # Fix for empty strings
  hdr[hdr == ""] <- " "

  flx <- flx %>%
    flextable::set_header_labels( # Needed bc headers must be unique
      values = setNames(
        as.vector(hdr[hnum, , drop = TRUE]),
        names(content)
      )
    )

  # If there are more rows -> add them
  if (hnum > 1) {
    for (i in seq(hnum - 1, 1)) {
      sel <- formatters::spans_to_viscell(spans[i, ])
      flx <- flextable::add_header_row(
        flx,
        top = TRUE,
        values = as.vector(hdr[i, sel]),
        colwidths = as.integer(spans[i, sel]) # xxx to fix
      )
    }
  }

  # Re-set the number of row count
  nr_body <- flextable::nrow_part(flx, part = "body")
  nr_header <- flextable::nrow_part(flx, part = "header")

  # Polish the inner horizontal borders from the header
  flx <- flx %>%
    .remove_hborder(part = "header", w = "all") %>%
    .add_hborder("header", ii = c(0, hnum), border = border)

  # ALIGNS - horizontal
  flx <- flx %>%
    .apply_alignments(mpf_aligns[seq_len(hnum), , drop = FALSE], "header") %>%
    .apply_alignments(mpf_aligns[-seq_len(hnum), , drop = FALSE], "body")

  # Rownames indentation
  checkmate::check_number(indent_size, null.ok = TRUE)
  if (is.null(indent_size)) {
    # Default indent_size in {rtables} is 2 characters
    indent_size <- matform$indent_size * word_mm_to_pt(1) # default is 2mm (5.7pt)
  } else {
    indent_size <- indent_size * word_mm_to_pt(1)
  }

  # rdf contains information about indentation
  for (i in seq_len(nr_body)) {
    flx <- flextable::padding(flx,
      i = i, j = 1,
      padding.left = indent_size * rdf$indent[[i]] + left_right_fixed_margins, # margins
      padding.right = left_right_fixed_margins, # 0.19 mmm in pt (so not to touch the border)
      part = "body"
    )
  }

  # TOPLEFT
  # Principally used for topleft indentation, this is a bit of a hack xxx
  for (i in seq_len(nr_header)) {
    leading_spaces_count <- nchar(hdr[i, 1]) - nchar(stringi::stri_replace(hdr[i, 1], regex = "^ +", ""))
    header_indent_size <- leading_spaces_count * word_mm_to_pt(1)
    hdr[i, 1] <- stringi::stri_replace(hdr[i, 1], regex = "^ +", "")

    # This solution does not keep indentation
    # top_left_tmp2 <- paste0(top_left_tmp, collapse = "\n") %>%
    #   flextable::as_chunk() %>%
    #   flextable::as_paragraph()
    # flx <- flextable::compose(flx, i = hnum, j = 1, value = top_left_tmp2, part = "header")
    flx <- flextable::padding(flx,
      i = i, j = 1,
      padding.left = header_indent_size + left_right_fixed_margins, # margins
      padding.right = left_right_fixed_margins, # 0.19 mmm in pt (so not to touch the border)
      part = "header"
    )
  }

  # Adding referantial footer line separator if present
  if (length(matform$ref_footnotes) > 0 && isFALSE(footers_as_text)) {
    flx <- flextable::add_footer_lines(flx, values = matform$ref_footnotes) %>%
      .add_hborder(part = "body", ii = nrow(tt), border = border)
  }

  # Footer lines
  if (length(formatters::all_footers(tt)) > 0 && isFALSE(footers_as_text)) {
    flx <- flextable::add_footer_lines(flx, values = formatters::all_footers(tt)) %>%
      .add_hborder(part = "body", ii = nrow(tt), border = border)
  }

  # Apply the theme
  flx <- .apply_themes(flx, theme = theme, tbl_row_class = rtables::make_row_df(tt)$node_class)

  # lets do some digging into the choice of fonts etc
  if (is.null(fontspec)) {
    fontspec <- .extract_fontspec(flx)
  }
  # Calculate the needed colwidths
  if (is.null(colwidths)) {
    # what about margins?
    colwidths <- formatters::propose_column_widths(matform, fontspec = fontspec, indent_size = indent_size)
  }

  # Title lines (after theme for problems with lines)
  if (titles_as_header && length(formatters::all_titles(tt)) > 0 && any(nzchar(formatters::all_titles(tt)))) {
    flx <- .add_titles_as_header(flx, all_titles = formatters::all_titles(tt), bold = bold_titles) %>%
      flextable::border(
        part = "header", i = length(formatters::all_titles(tt)),
        border.bottom = border
      )
  }

  # xxx FIXME missing transformer from character based widths to mm or pt
  final_cwidths <- total_page_width * colwidths / sum(colwidths)

  flx <- flextable::width(flx, width = final_cwidths)

  # These final formatting need to work with colwidths
  flx <- flextable::set_table_properties(flx,
    layout = ifelse(autofit_to_page, "autofit", "fixed"),
    align = "left",
    opts_word = list(
      "split" = FALSE,
      "keep_with_next" = TRUE
    )
  )

  # NB: autofit or fixed may be switched if widths are correctly staying in the page
  flx <- flextable::fix_border_issues(flx) # Fixes some rendering gaps in borders

  flx
}


# only used in pagination
.tab_to_colpath_set <- function(tt) {
  vapply(
    rtables::collect_leaves(rtables::coltree(tt)),
    function(y) paste(rtables:::pos_to_path(rtables:::tree_pos(y)), collapse = " "),
    ""
  )
}
.figure_out_colinds <- function(subtab, fulltab) {
  match(
    .tab_to_colpath_set(subtab),
    .tab_to_colpath_set(fulltab)
  )
}

.add_titles_as_header <- function(flx, all_titles, bold = TRUE) {
  all_titles <- all_titles[nzchar(all_titles)] # Remove empty titles (use " ")

  flx <- flx %>%
    flextable::add_header_lines(values = all_titles, top = TRUE) %>%
    # Remove the added borders
    flextable::border(
      part = "header", i = seq_along(all_titles),
      border.top = flextable::fp_border_default(width = 0),
      border.bottom = flextable::fp_border_default(width = 0),
      border.left = flextable::fp_border_default(width = 0),
      border.right = flextable::fp_border_default(width = 0)
    ) %>%
    flextable::bg(part = "header", i = seq_along(all_titles), bg = "white")

  if (isTRUE(bold)) {
    flx <- flextable::bold(flx, part = "header", i = seq_along(all_titles))
  } else if (checkmate::test_integerish(bold)) {
    if (any(bold > length(all_titles))) {
      stop("bold values are greater than the number of titles lines.")
    }
    flx <- flextable::bold(flx, part = "header", i = bold)
  }

  flx
}

.apply_themes <- function(flx, theme, tbl_row_class = "") {
  if (is.null(theme)) {
    return(flx)
  }
  # Wrap theme in a list if it's not already a list
  theme_list <- if (is.list(theme)) theme else list(theme)
  # Loop through the themes
  for (them in theme_list) {
    flx <- them(
      flx,
      tbl_row_class = tbl_row_class # These are ignored if not in the theme
    )
  }

  flx
}

.extract_fontspec <- function(test_flx) {
  font_sz <- test_flx$header$styles$text$font.size$data[1, 1]
  font_fam <- test_flx$header$styles$text$font.family$data[1, 1]
  font_fam <- "Courier" # Fix if we need it -> coming from gpar and fontfamily Arial not being recognized

  formatters::font_spec(font_family = font_fam, font_size = font_sz, lineheight = 1)
}

.apply_alignments <- function(flx, aligns_df, part) {
  # List of characters you want to search for
  search_chars <- unique(c(aligns_df))

  # Loop through each character and find its indexes
  for (char in search_chars) {
    indexes <- which(aligns_df == char, arr.ind = TRUE)
    tmp_inds <- as.data.frame(indexes)
    flx <- flx %>%
      flextable::align(
        i = tmp_inds[["row"]],
        j = tmp_inds[["col"]],
        align = char,
        part = part
      )
  }

  flx
}

# Themes -----------------------------------------------------------------------
#

#' @describeIn tt_to_flextable Main theme function for [export_as_docx()].
#'
#' @param font (`string`)\cr font. Defaults to `"Arial"`. If the font given is not available, the `flextable` default
#'   is used instead. For options, consult the family column from [systemfonts::system_fonts()].
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

    return(flx)
  }
}

.add_borders <- function(flx, border, ncol) {
  # all borders
  flx <- flx %>%
    flextable::border_outer(part = "body", border = border) %>%
    # flextable::border_outer(part = "header", border = border) %>%
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

# Polish horizontal borders
.remove_hborder <- function(flx, part, w = c("top", "bottom", "inner")) {
  # If you need to remove all of them
  if (length(w) == 1 && w == "all") {
    w <- eval(formals(.remove_hborder)$w)
  }

  if (any(w == "top")) {
    flx <- flextable::hline_top(flx,
      border = flextable::fp_border_default(width = 0),
      part = part
    )
  }
  if (any(w == "bottom")) {
    flx <- flextable::hline_bottom(flx,
      border = flextable::fp_border_default(width = 0),
      part = part
    )
  }
  # Inner horizontal lines removal
  if (any(w == "inner")) {
    flx <- flextable::border_inner_h(
      flx,
      border = flextable::fp_border_default(width = 0),
      part = part
    )
  }
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

# Add horizontal border
.add_hborder <- function(flx, part, ii, border) {
  if (any(ii == 0)) {
    flx <- flextable::border(flx, i = 1, border.top = border, part = part)
    ii <- ii[!(ii == 0)]
  }
  if (length(ii) > 0) {
    flx <- flextable::border(flx, i = ii, border.bottom = border, part = part)
  }
  flx
}
