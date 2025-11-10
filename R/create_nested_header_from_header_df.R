#' Produce column merge metadata for nested Excel headers
#'
#' @description
#' Internal helper that determines which cells in each column of a header
#' definition data frame (`header_df`) should be vertically merged when
#' building hierarchical headers in Excel.
#'
#' @param header_df A data frame or tibble, where each column corresponds to a
#' header column in Excel and each row represents a header level (e.g. l1–l4).
#'
#' @return
#' A list of length `ncol(header_df)`, where each element is a list with two
#' elements:
#' \describe{
#'   \item{`start`}{Row index of the first mergeable cell (NA if none).}
#'   \item{`end`}{Row index of the last mergeable cell.}
#' }
#'
#' @details
#' For each column, the function finds the **last non-`NA` value** in the column.
#' If that value is not in the final row, the rows from that position to the
#' bottom are marked for merging. If it *is* already the bottom, no merge is
#' needed.
#'
#' @examples
#' # Example header definition
#' hdr <- tibble::tibble(
#'   l1 = c("Group A", NA, NA),
#'   l2 = c("Sub A1", "Sub A2", NA),
#'   l3 = c(NA, NA, "Detail A3")
#' )
#' produce_col_merge(hdr)
#'
#' @export
produce_col_merge <- function(header_df) {
  lapply(seq_len(ncol(header_df)), function(i) {
    # Extract column i as vector
    temp <- header_df[, i] %>% dplyr::pull()
    
    # Find the last non-NA row index
    start <- which(!is.na(temp)) %>% max()
    
    # If the last non-NA is at the bottom, no merge needed
    if (start == nrow(header_df)) {
      list(start = NULL, end = NULL)
    } else {
      list(start = start, end = nrow(header_df))
    }
  })
}

#' Create nested Excel headers from a header definition data frame
#'
#' @description
#' Builds multi-level (nested) headers in an Excel worksheet based on a structured header definition
#' table (`header_df`).
#'
#' Each row in `header_df` represents one header level (e.g. top → bottom),
#' and each column corresponds to an Excel column.
#'
#' @param wb An openxlsx workbook object.
#' @param sheet A sheet name (character string) within the workbook.
#' @param header_df A tibble or data frame with one or more header levels as
#' rows (`l1`, `l2`, …). Columns represent the final Excel columns.
#'
#' @return
#' The modified workbook object (`wb`) with merged cells and styled nested
#' headers.
#'
#' @details
#' The function:
#' \enumerate{
#'   \item Computes vertical merge metadata using \code{\link{produce_col_merge}}.
#'   \item Iterates through header levels to compute merge groupings and nesting
#'         structure between consecutive header rows.
#'   \item Calls \code{TGexcel::create_nested_header_style()} to apply merged-cell
#'         formatting for each header level pair.
#'   \item Finally merges vertically identical column header segments using
#'         \pkg{openxlsx} cell merge operations.
#' }
#'
#' @examples
#' \dontrun{
#' wb <- openxlsx::createWorkbook()
#' openxlsx::addWorksheet(wb, "Sheet1")
#'
#' hdr <- tibble::tibble(
#'   l1 = c("Group A", "Group A", "Group B"),
#'   l2 = c("Total", "Male", "Total")
#' )
#' create_nested_header_from_header_df(wb, "Sheet1", hdr)
#' openxlsx::saveWorkbook(wb, "nested_headers.xlsx", overwrite = TRUE)
#' }
#'
#' @export
create_nested_header_from_header_df <- function(wb, sheet, header_df) {
  # --- 1️⃣ Compute column merge information
  col_merge <- produce_col_merge(header_df)
  nesting_vecs <- list()
  header_cols <- tibble::tibble(col = seq_len(ncol(header_df)))
  
  # --- 2️⃣ Loop over header levels (rows)
  for (i in seq_len(nrow(header_df))) {
    row <- header_df[i, 1:ncol(header_df)]
    row_vec <- as.character(row) %>%
      stats::setNames(NULL) %>%
      tidyr::replace_na("")
    
    # Columns with complete (non-NA) header info
    merge_vec <- which(colSums(is.na(row)) == 0) %>% stats::setNames(NULL)
    merge_vec <- c(seq_len(merge_vec[1]), merge_vec[2:length(merge_vec)])
    
    # Current and previous level variable names
    var_name <- paste0("l", i)
    var_nameLast <- paste0("l", i - 1)
    
    temp <- tibble::tibble(col = merge_vec, l = merge_vec) %>%
      stats::setNames(c("col", var_name))
    
    # Join temporary structure
    header_cols <- header_cols %>%
      dplyr::left_join(temp, "col")
    
    # --- 3️⃣ Propagate previous level labels downward if missing
    if (i > 1) {
      header_cols <- header_cols %>%
        dplyr::mutate(!!var_name := dplyr::coalesce(!!rlang::sym(var_nameLast), !!rlang::sym(var_name)))
    }
    
    # Identify valid (non-NA) indices
    index_vec <- header_cols %>%
      dplyr::filter(!is.na(!!rlang::sym(var_name))) %>%
      dplyr::pull(!!rlang::sym(var_name))
    
    # Extract label values and nesting structure
    vars <- row_vec[index_vec]
    if (index_vec[length(index_vec)] != nrow(header_cols)) {
      nesting <- diff(c(index_vec, (nrow(header_cols) + 1)))
    } else {
      nesting <- diff(index_vec)
    }
    
    if (i == 1) {
      vars_next_level <- row_vec[merge_vec]
    } else {
      vars_next_level <- row_vec
    }
    
    nesting_vecs[[i]] <- list(
      index = index_vec,
      vars = vars,
      vars_next_level = vars_next_level,
      nesting = nesting
    )
  }
  
  # --- 4️⃣ Create merged cells and apply nested styles
  for (i in seq_len(length(nesting_vecs) - 1)) {
    create_nested_header_style(
      wb,
      sheet,
      vars_level1 = nesting_vecs[[i]]$vars,
      vars_level2 = nesting_vecs[[i + 1]]$vars_next_level,
      startRow = i + 2,
      nesting = nesting_vecs[[i]]$nesting
    )
  }
  
  # --- 5️⃣ Apply vertical merges for identical column headers
  for (k in seq_along(col_merge)) {
    elem <- col_merge[[k]]
    if (is.null(elem[["start"]])) next
    
    row_range <- (elem[["start"]]:elem[["end"]]) + 2
    openxlsx::removeCellMerge(wb, sheet, cols = k, rows = row_range)
    openxlsx::mergeCells(wb, sheet, cols = k, rows = row_range)
  }
  
  # --- 6️⃣ Return workbook
  wb
}
