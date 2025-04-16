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
#' @param tt (`TableTree`, `listing_df`, or related class)\cr a `TableTree` or `listing_df` object representing
#'   a populated table or listing.
#' @param theme (`function` or `NULL`)\cr a theme function designed to change the layout and style of a `flextable`
#'   object. Defaults to `theme_docx_default()`, the classic Microsoft Word output style. If `NULL`, a table with style
#'   similar to the `rtables` default will be produced. See Details below for more information.
#' @param indent_size (`numeric(1)`)\cr indentation size. If `NULL`, the default indent size of the table (see
#'   [formatters::matrix_form()] `indent_size`, default is 2) is used. To work with `docx`, any size is multiplied by
#'   1 mm (2.83 pt) by default.
#' @param titles_as_header (`flag`)\cr Controls how titles are rendered relative to the table.
#'   If `TRUE` (default), the main title ([formatters::main_title()]) and subtitles
#'   ([formatters::subtitles()]) are added as distinct header rows within the
#'   `flextable` object itself. If `FALSE`, titles are rendered as a separate paragraph
#'   of text placed immediately before the table.
#' @param bold_titles (`flag` or `integer`)\cr whether titles should be bold (defaults to `TRUE`). If one or more
#'   integers are provided, these integers are used as indices for lines at which titles should be bold.
#' @param integrate_footers (`flag`)\cr Controls how footers are rendered relative to the table.
#'   If `TRUE` (default), footers (e.g., [formatters::main_footer()], [formatters::prov_footer()])
#'   are integrated directly into the `flextable` object, typically appearing as footnotes
#'   below the table body with a smaller font. If `FALSE`, footers are rendered as a
#'   separate paragraph of text placed immediately after the table.
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
                            integrate_footers = TRUE,
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
  if (inherits(tt, "list")) {
    stop("Please use paginate = TRUE or mapply() to create multiple outputs. export_as_docx accepts lists.")
  }
  if (!inherits(tt, "VTableTree") && !inherits(tt, "listing_df")) {
    stop("Input object is not an rtables' or rlistings' object.")
  }
  checkmate::assert_flag(titles_as_header)
  checkmate::assert_flag(integrate_footers)
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
    # expected_lpp is still not really usable atm
    expected_lpp <- starting_lpp + max(which(cumsum_page_heights < total_page_height))

    if (!is.null(lpp) && starting_lpp + 1 > lpp) {
      stop("Header rows are more than selected lines per pages (lpp).")
    }

    # Main tabulation system
    tabs <- rtables::paginate_table(tt,
      fontspec = fontspec,
      lpp = lpp,
      cpp = cpp, tf_wrap = tf_wrap, max_width = max_width, # This can only be trial an error
      ...
    )

    # Indices for column width
    cinds <- lapply(tabs, function(tb) c(1, .figure_out_colinds(tb, tt) + 1L))
    args$colwidths <- NULL
    args$tt <- NULL
    cl <- if (!is.null(colwidths)) {
      lapply(cinds, function(ci) colwidths[ci])
    } else {
      lapply(cinds, function(ci) {
        NULL
      })
    }

    # Main return
    return(
      mapply(tt_to_flextable,
        tt = tabs, colwidths = cl,
        MoreArgs = args,
        SIMPLIFY = FALSE
      )
    )
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
  if (length(matform$ref_footnotes) > 0 && isTRUE(integrate_footers)) {
    flx <- flextable::add_footer_lines(flx, values = matform$ref_footnotes) %>%
      .add_hborder(part = "body", ii = nrow(tt), border = border)
  }

  # Footer lines
  if (length(formatters::all_footers(tt)) > 0 && isTRUE(integrate_footers)) {
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

  # Handling of horizontal separators -> done afterwards because otherwise count of lines is sloppy
  if (!all(is.na(matform$row_info$trailing_sep))) {
    flx <- add_flextable_separators(
      flx,
      matform$row_info$trailing_sep,
      border = officer::fp_border(width = 1, color = "grey60"),
      padding = 10
    )
  }

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
    flextable::align(part = "header", i = seq_along(all_titles), align = "left") %>%
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
