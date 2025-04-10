test_that("Errors are thrown for invalid inputs and values", {
  # Setup
  test_data <- data.frame(colA = 1:3, colB = letters[1:3])
  ft_base <- flextable(test_data)

  # Case 1: Invalid Character in trailing_sep
  sep_invalid_char <- c("-", "x", NA)
  expect_error(
    add_flextable_separators(ft_base, sep_invalid_char),
    regexp = "Invalid character\\(s\\) found.*'x'"
  )

  # Case 2: Length Mismatch
  sep_wrong_len <- c("-", NA) # Length 2 vs 3 rows
  expect_warning(
    out <- add_flextable_separators(ft_base, sep_wrong_len),
    regexp = "Length of 'trailing_sep'.*\\(2\\).*must equal.*\\(3\\)"
  )

  # Case 3: Invalid ft input (not a flextable)
  expect_error(
    add_flextable_separators(test_data, rep(NA, 3)), # Pass data.frame
    regexp = "Input 'ft' must be a flextable object"
  )

  # Case 4: Invalid trailing_sep input (not atomic vector)
  expect_error(
    add_flextable_separators(ft_base, list(NA, "-", NA)) # Pass list
  )
})

test_that("Separator '-' successfully applies hline", {
  # Setup
  test_data <- data.frame(colA = 1:4, colB = letters[1:4])
  ft_base <- flextable(test_data)
  sep_ctrl <- c("-", NA, "-", NA) # Apply hline after row 1 and 3
  border_rows <- which(sep_ctrl == "-")

  # Action
  ft_mod <- NULL
  expect_no_error(ft_mod <- add_flextable_separators(ft_base, sep_ctrl))

  # Basic Checks
  expect_s3_class(ft_mod, "flextable")
  expect_equal(flextable::nrow_part(ft_mod, "body"), nrow(test_data),
               label = "Row count remains unchanged")

  # Optional (potentially brittle) check: Verify border applied
  # Check border width > 0 for the first column of rows where '-' was specified
  # Assumes default border width is > 0
  if (length(border_rows) > 0) {
    # Note: Accessing internal structure $body$styles might change between versions
    expect_true(all(ft_mod$body$styles$cells$border.width.bottom$data[border_rows, 1] == 1.0),
                label = "Bottom border width > 0 expected where '-' was applied")
  }
  # Check rows where NA was specified (should have default border width, likely 0)
  na_rows <- which(is.na(sep_ctrl))
  # Last row is bottom border -> remove
  na_rows <- na_rows[na_rows != length(sep_ctrl)]
  if (length(na_rows) > 0) {
    expect_true(all(ft_mod$body$styles$cells$border.width.bottom$data[na_rows, 1] == 0),
                label = "Bottom border width should be 0 where NA was applied")
  }
})

test_that("Separator ' ' successfully applies padding", {
  # Setup
  test_data <- data.frame(colA = 1:4, colB = letters[1:4])
  ft_base <- flextable(test_data)
  sep_ctrl <- c(" ", NA, " ", NA) # Apply padding after row 1 and 3
  padding_rows <- which(sep_ctrl == " ")
  custom_padding_val <- 12 # Use a non-default value for easier checking

  # Action
  ft_mod <- NULL
  expect_no_error(
    ft_mod <- add_flextable_separators(ft_base, sep_ctrl, padding = custom_padding_val)
  )

  # Basic Checks
  expect_s3_class(ft_mod, "flextable")
  expect_equal(flextable::nrow_part(ft_mod, "body"), nrow(test_data),
               label = "Row count remains unchanged")

  # Optional (potentially brittle) check: Verify padding applied
  # Check bottom padding for the first column of rows where ' ' was specified
  if (length(padding_rows) > 0) {
    # Note: Accessing internal structure $body$styles might change between versions
    expect_equal(ft_mod$body$styles$pars$padding.bottom$data[padding_rows, 1],
                 rep(custom_padding_val, length(padding_rows)),
                 label = "Bottom padding should match specified value where ' ' was applied")
  }
  # Check rows where NA was specified (should have default padding)
  na_rows <- which(is.na(sep_ctrl))
  if (length(na_rows) > 0) {
    # Get default padding from original table for comparison
    default_padding <- ft_mod$body$styles$pars$padding.bottom$data[na_rows[1], ][[1]]
    expect_equal(ft_mod$body$styles$pars$padding.bottom$data[na_rows, 1],
                 rep(default_padding, length(na_rows)),
                 label = "Bottom padding should be default where NA was applied")
  }
})

test_that("Separator NA results in no changes", {
  # Setup
  test_data <- data.frame(colA = 1:3, colB = letters[1:3])
  # Add some non-default property to base to make identical() check more meaningful
  ft_base <- flextable(test_data) %>% set_caption("NA Test")
  sep_all_na <- rep(NA, nrow(test_data))

  # Action
  ft_mod <- NULL
  expect_no_error(ft_mod <- add_flextable_separators(ft_base, sep_all_na))

  # Expectation
  # Function should return the exact same object if no changes were made
  expect_identical(ft_mod, ft_base)
})
