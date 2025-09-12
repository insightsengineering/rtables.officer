# docx (flextable) -----------------------------------------------------------
#' Export to a Word document
#'
#' From an `rtables` table, produce a self-contained Word document or attach it to a template Word
#' file (`template_file`). This function is based on the [tt_to_flextable()] transformer and
#' the `officer` package.
#'
#' @inheritParams rtables::gen_args
#' @inheritParams tt_to_flextable
#' @param file (`string`)\cr output file. Must have `.docx` extension.
#' @param add_page_break (`flag`)\cr whether to add a page break after the table (`TRUE`) or not (`FALSE`).
#' @param add_template_page_numbers (`flag`)\cr whether to add page numbers to the word document as page footer. This
#'   uses templates to achieve it. Defaults to `TRUE`. Consider adding your own template file if you want
#'   more customization.
#' @param section_properties (`officer::prop_section`)\cr an [officer::prop_section()] object which sets margins and
#'   page size. Defaults to [section_properties_default()].
#' @param doc_metadata (`list` of `string`)\cr any value that can be used as metadata by
#'   [officer::set_doc_properties()]. Important text values are `title`, `subject`, `creator`, and `description`,
#'   while `created` is a date object.
#' @param template_file (`string`)\cr template file that `officer` will use as a starting point for the final
#'   document. Document attaches the table and uses the defaults defined in the template file.
#' @param ... (`any`)\cr additional arguments passed to [tt_to_flextable()].
#'
#' @note `export_as_docx()` has few customization options available. If you require specific formats and details,
#'   we suggest that you use [tt_to_flextable()] prior to `export_as_docx()`. If the table is modified first using
#'   [tt_to_flextable()], the `titles_as_header` and `integrate_footers` parameters must be re-specified.
#'
#' @return No return value, called for side effects
#'
#' @details
#' Pagination Behavior for Titles and Footers (this behavior is experimental at the moment):
#'
#' The rendering of titles and footers interacts with table pagination as follows:
#' * **Titles:** When `titles_as_header = TRUE` (default), the integrated title
#'     header rows typically repeat at the top of each new page if the table spans
#'     multiple pages. Setting `titles_as_header = FALSE` renders titles as a
#'     separate paragraph only once before the table begins.
#' * **Footers:** Regardless of the `integrate_footers` setting, footers appear
#'     only once. Integrated footnotes (`integrate_footers = TRUE`) appear at the
#'     very end of the complete table, and separate text paragraphs
#'     (`integrate_footers = FALSE`) appear after the complete table. Footers
#'     do not repeat on each page.
#'
#' @seealso [tt_to_flextable()]
#'
#' @examples
#' lyt <- basic_table() %>%
#'   split_cols_by("ARM") %>%
#'   analyze(c("AGE", "BMRKR2", "COUNTRY"))
#'
#' tbl <- build_table(lyt, ex_adsl)
#'
#' # See how the section_properties_portrait() function is built for customization
#' tf <- tempfile(tmpdir = tempdir(check = TRUE), fileext = ".docx")
#' export_as_docx(tbl,
#'   file = tf,
#'   section_properties = section_properties_default(orientation = "landscape")
#' )
#'
#' @export
export_as_docx <- function(tt,
                           file,
                           add_page_break = FALSE,
                           add_template_page_numbers = TRUE,
                           titles_as_header = TRUE,
                           integrate_footers = TRUE,
                           section_properties = section_properties_default(),
                           doc_metadata = NULL,
                           template_file = NULL,
                           ...) {
  # Checks
  checkmate::assert_flag(add_page_break)
  checkmate::assert_flag(add_template_page_numbers)
  do_tt_error <- FALSE

  # tt can be a VTableTree, a flextable, or a list of VTableTree or flextable objects
  if (inherits(tt, "VTableTree") || inherits(tt, "listing_df")) {
    tt <- tt_to_flextable(tt,
      titles_as_header = titles_as_header,
      integrate_footers = integrate_footers,
      ...
    )
  }

  if (inherits(tt, "flextable")) {
    flex_tbl_list <- list(tt)
  } else if (inherits(tt, "list")) {
    if (inherits(tt[[1]], "VTableTree") || inherits(tt[[1]], "listing_df")) {
      flex_tbl_list <- mapply(
        tt_to_flextable,
        tt = tt,
        MoreArgs = list(
          titles_as_header = titles_as_header,
          integrate_footers = integrate_footers,
          ...
        ),
        SIMPLIFY = FALSE
      )
    } else if (inherits(tt[[1]], "flextable")) {
      flex_tbl_list <- tt
    } else {
      do_tt_error <- TRUE
    }
  } else {
    do_tt_error <- TRUE
  }

  # Error handling for wrong tt object
  if (isTRUE(do_tt_error)) {
    stop("tt must be a TableTree/listing_df, a flextable, or a list of TableTree/listing_df or flextable objects.")
  }

  # If additional text needs to be added, we need to have info about the font and size
  if (isFALSE(titles_as_header) || isFALSE(integrate_footers)) {
    flx_fpt <- .extract_font_and_size_from_flx(flex_tbl_list[[1]]) # Using the first only
  }

  # Check if the template file is inserted and exists
  if (!is.null(template_file) && !file.exists(template_file)) {
    template_file <- NULL
  }


  # Create a new empty Word document
  if (!is.null(template_file)) {
    doc <- officer::read_docx(template_file)
  } else {
    if (isTRUE(add_template_page_numbers)) {
      template_file <- .get_template_file(section_properties)
    }
    doc <- officer::read_docx(template_file)
    if (!is.null(section_properties) && isFALSE(add_template_page_numbers)) {
      doc <- officer::body_set_default_section(doc, section_properties)
    }
  }


  # Check page widths
  flex_tbl_list <- lapply(flex_tbl_list, function(flx) {
    if (flx$properties$layout != "autofit") { # fixed layout
      page_width <- section_properties$page_size$width
      dflx <- dim(flx)
      if (abs(sum(unname(dflx$widths)) - page_width) > 1e-2) {
        warning(
          "The total table width does not match the page width. The column widths",
          " will be resized to fit the page. Please consider modifying the parameter",
          " total_page_width in tt_to_flextable()."
        )

        final_cwidths <- page_width * unname(dflx$widths) / sum(unname(dflx$widths))
        flx <- flextable::width(flx, width = final_cwidths)
      }
    }
    flx
  })

  # Extract title
  if (isFALSE(titles_as_header) && inherits(tt, "VTableTree")) {
    ts_tbl <- formatters::all_titles(tt)
    if (length(ts_tbl) > 0) {
      doc <- add_text_par(doc, ts_tbl, flx_fpt$fpt)
    }
  }

  # Add the flextable(s) to the document
  for (flex_tbl_i in flex_tbl_list) {
    doc <- flextable::body_add_flextable(doc, flex_tbl_i, align = "left")
    # Add a page break after each table
    if (isTRUE(add_page_break)) {
      doc <- body_add_break(doc)
    }
  }

  # add footers as paragraphs
  if (isFALSE(integrate_footers) && inherits(tt, "VTableTree")) {
    # Adding referential footer line separator if present
    # (this is usually done differently, i.e. inside footnotes)
    matform <- rtables::matrix_form(tt, indent_rownames = TRUE)
    if (length(matform$ref_footnotes) > 0) {
      doc <- add_text_par(doc, matform$ref_footnotes, flx_fpt$fpt_footer)
    }
    # Footer lines
    if (length(formatters::all_footers(tt)) > 0) {
      doc <- add_text_par(doc, formatters::all_footers(tt), flx_fpt$fpt_footer)
    }
  }

  if (!is.null(doc_metadata)) {
    # Checks for values rely on officer function
    doc <- do.call(officer::set_doc_properties, c(list("x" = doc), doc_metadata))
  }

  # Save the Word document to a file
  print(doc, target = file)

  invisible(TRUE)
}

.extract_font_and_size_from_flx <- function(flx) {
  # Ugly but I could not find a getter for font.size
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

# Shorthand to add text paragraph
add_text_par <- function(doc, chr_v, text_format) {
  for (ii in seq_along(chr_v)) {
    cur_fp <- officer::fpar(officer::ftext(chr_v[ii], prop = text_format))
    doc <- officer::body_add_fpar(doc, cur_fp)
  }
  doc
}

#' @describeIn export_as_docx Helper function that defines standard portrait properties for tables.
#' @param page_size (`string`) page size. Can be `"letter"` or `"A4"`. Defaults to `"letter"`.
#' @param orientation (`string`) page orientation. Can be `"portrait"` or `"landscape"`. Defaults to `"portrait"`.
#'
#' @export
section_properties_default <- function(page_size = c("letter", "A4"),
                                       orientation = c("portrait", "landscape")) {
  page_size <- page_size[1]
  orientation <- orientation[1]
  checkmate::assert_choice(
    page_size,
    eval(formals(section_properties_default)$page_size)
  )
  checkmate::assert_choice(
    orientation,
    eval(formals(section_properties_default)$orientation)
  )

  if (page_size == "letter") {
    page_size <- officer::page_size(
      orient = orientation,
      width = 8.5, height = 11
    )
  } else { # A4
    page_size <- officer::page_size(
      orient = orientation,
      width = 8.27, height = 11.69
    )
  }

  # Final output
  officer::prop_section(
    page_size = page_size,
    type = "continuous",
    page_margins = margins_potrait()
  )
}

#' @describeIn export_as_docx Helper function that defines standard portrait margins for tables.
#' @export
margins_potrait <- function() {
  officer::page_mar(bottom = 0.98, top = 0.95, left = 1.5, right = 1, gutter = 0)
}
#' @describeIn export_as_docx Helper function that defines standard landscape margins for tables.
#' @export
margins_landscape <- function() {
  officer::page_mar(bottom = 1, top = 1.5, left = 0.98, right = 0.95, gutter = 0)
}


.get_template_file <- function(section_properties) {
  orient <- section_properties$page_size$orient
  page_sz <- section_properties$page_size

  size <- NULL
  warn_msg <- NULL
  if (orient == "landscape") {
    if (page_sz$width == 11 && page_sz$height == 8.5) {
      size <- "letter"
    } else if (page_sz$width == 11.69 && page_sz$height == 8.27) {
      size <- "A4"
    } else {
      warn_msg <- c(
        "Adding page numbers is supported only A4 and letter size.",
        "Page numbers will not be added."
      )
    }
  } else if (orient == "portrait") {
    if (page_sz$width == 8.5 && page_sz$height == 11) {
      size <- "letter"
    } else if (page_sz$width == 8.27 && page_sz$height == 11.69) {
      size <- "A4"
    } else {
      warn_msg <- c(
        "Adding page numbers is supported only A4 and letter size.",
        "Page numbers will not be added."
      )
    }
  } else {
    stop(
      "Adding page numbers is supported only for landscape and portrait orientation.",
      "Page numbers will not be added."
    )
  }

  if (!is.null(warn_msg)) {
    warning(warn_msg)
  }

  if (is.null(size)) {
    ret <- NULL
  } else if (size == "A4") {
    ret <- file.path(
      system.file(package = "rtables.officer"),
      ifelse(orient == "landscape",
        "docx_templates/a4_landscape.docx",
        "docx_templates/a4_portrait.docx"
      )
    )
  } else if (size == "letter") {
    ret <- file.path(
      system.file(package = "rtables.officer"),
      ifelse(orient == "landscape",
        "docx_templates/letter_landscape.docx",
        "docx_templates/letter_portrait.docx"
      )
    )
  }

  ret
}
