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
#' @param titles_as_header (`flag`)\cr whether the table should be self-contained with additional header rows created
#'   for titles and subtitles (`TRUE`), or titles and subtitles should be added as a paragraph of text above the table
#'   (`FALSE`). Defaults to `FALSE`.
#' @param footers_as_text (`flag`)\cr whether footers should be added as a new paragraph after the table (`TRUE`) or
#'   the table should be self-contained, implementing `flextable`-style footnotes (`FALSE`) with the same style but a
#'   smaller font. Defaults to `TRUE`.
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
#'   [tt_to_flextable()], the `titles_as_header` and `footer_as_text` parameters must be re-specified.
#'
#' @return No return value, called for side effects
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
#' tf <- tempfile(fileext = ".docx")
#' export_as_docx(tbl,
#'   file = tf,
#'   section_properties = section_properties_default(orientation = "landscape")
#' )
#'
#' @export
export_as_docx <- function(tt,
                           file,
                           add_page_break = FALSE,
                           titles_as_header = FALSE,
                           footers_as_text = TRUE,
                           section_properties = section_properties_default(),
                           doc_metadata = NULL,
                           template_file = NULL,
                           ...) {
  # Checks
  checkmate::assert_flag(add_page_break)

  # tt can be a VTableTree, a flextable, or a list of VTableTree or flextable objects
  if (inherits(tt, "VTableTree")) {
    flex_tbl_list <- tt_to_flextable(tt,
      titles_as_header = titles_as_header,
      footers_as_text = footers_as_text,
      ...
    ) %>%
      list()

  } else if (inherits(tt, "flextable")) {
    flex_tbl_list <- list(tt)
  } else if (inherits(tt, "list")) {
    if (inherits(tt[[1]], "VTableTree")) {
      flex_tbl_list <- mapply(
        tt_to_flextable,
        tt = tt,
        MoreArgs = list(
          titles_as_header = titles_as_header,
          footers_as_text = footers_as_text,
          ...
        )
      )
    } else if (inherits(tt[[1]], "flextable")) {
      flex_tbl_list <- tt
    } else {
      stop("tt must be a VTableTree, a flextable, or a list of VTableTree or flextable objects.")
    }
  } else {
    stop("The table must be a VTableTree, a flextable, or a list of VTableTree or flextable objects.")
  }

  # If additional text needs to be added, we need to have info about the font and size
  if (isFALSE(titles_as_header) || isTRUE(footers_as_text)) {
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
    doc <- officer::read_docx()
  }

  # page width and orientation settings
  doc <- officer::body_set_default_section(doc, section_properties)

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
  if (isTRUE(footers_as_text) && inherits(tt, "VTableTree")) {
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
