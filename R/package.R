#' @keywords internal
"_PACKAGE"

#' @importFrom lifecycle deprecated
#' @importFrom magrittr %>%
#' @import methods
#' @import rtables
#' @import officer
#' @import flextable
NULL


#' General argument conventions
#'
#' @param ... additional parameters passed to methods or tabulation functions.
#' @param alt_counts_df (`data.frame` or `tibble`)\cr alternative full dataset the rtables framework will use
#'   *only* when calculating column counts.
#' @param colwidths (`numeric`)\cr a vector of column widths for use in vertical pagination.
#' @param cvar (`string`)\cr the variable, if any, that the content function should accept. Defaults to `NA`.
#' @param df (`data.frame` or `tibble`)\cr dataset.
#' @param hsep (`string`)\cr set of characters to be repeated as the separator between the header and body of
#'   the table when rendered as text. Defaults to a connected horizontal line (unicode 2014) in locals that use a UTF
#'   charset, and to `-` elsewhere (with a once per session warning). See [formatters::set_default_hsep()] for further
#'   information.
#' @param indent_size (`numeric(1)`)\cr number of spaces to use per indent level. Defaults to 2.
#' @param inset (`numeric(1)`)\cr number of spaces to inset the table header, table body, referential footnotes, and
#'   main_footer, as compared to alignment of title, subtitle, and provenance footer. Defaults to 0 (no inset).
#' @param label (`string`)\cr a label (not to be confused with the name) for the object/structure.
#' @param label_pos (`string`)\cr location where the variable label should be displayed. Accepts `"hidden"`
#'   (default for non-analyze row splits), `"visible"`, `"topleft"`, and `"default"` (for analyze splits only). For
#'   `analyze` calls, `"default"` indicates that the variable should be visible if and only if multiple variables are
#'   analyzed at the same level of nesting.
#' @param na_str (`string`)\cr string that should be displayed when the value of `x` is missing. Defaults to `"NA"`.
#' @param obj (`ANY`)\cr the object for the accessor to access or modify.
#' @param object (`ANY`)\cr the object to modify in place.
#' @param page_prefix (`string`)\cr prefix to be appended with the split value when forcing pagination between
#'   the children of a split/table.
#' @param path (`character`)\cr a vector path for a position within the structure of a `TableTree`. Each element
#'   represents a subsequent choice amongst the children of the previous choice.
#' @param pos (`numeric`)\cr which top-level set of nested splits should the new layout feature be added to. Defaults
#'   to the current split.
#' @param section_div (`string`)\cr string which should be repeated as a section divider after each group defined
#'   by this split instruction, or `NA_character_` (the default) for no section divider.
#' @param spl (`Split`)\cr a `Split` object defining a partitioning or analysis/tabulation of the data.
#' @param table_inset (`numeric(1)`)\cr number of spaces to inset the table header, table body, referential footnotes,
#'   and main footer, as compared to alignment of title, subtitles, and provenance footer. Defaults to 0 (no inset).
#' @param topleft (`character`)\cr override values for the "top left" material to be displayed during printing.
#' @param tr (`TableRow` or related class)\cr a `TableRow` object representing a single row within a populated table.
#' @param tt (`TableTree` or related class)\cr a `TableTree` object representing a populated table.
#' @param value (`ANY`)\cr the new value.
#' @param verbose (`flag`)\cr whether additional information should be displayed to the user. Defaults to `FALSE`.
#' @param x (`ANY`)\cr an object.
#'
#' @return No return value.
#'
#' @family conventions
#' @name gen_args
#' @keywords internal
gen_args <- function(df, alt_counts_df, spl, pos, tt, tr, verbose, colwidths, obj, x,
                     value, object, path, label, label_pos, # visible_label,
                     cvar, topleft, page_prefix, hsep, indent_size, section_div, na_str, inset,
                     table_inset,
                     ...) {
  NULL
}
