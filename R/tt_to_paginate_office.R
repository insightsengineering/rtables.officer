pos_to_path <- function(pos) {
  spls <- pos_splits(pos)
  vals <- pos_splvals(pos)

  path <- character()
  for (i in seq_along(spls)) {
    nm <- obj_name(spls[[i]])
    val_i <- value_names(vals[[i]])
    path <- c(
      path,
      obj_name(spls[[i]]),
      ## rawvalues(vals[[i]]))
      if (!is.na(val_i)) val_i
    )
  }
  path
}
