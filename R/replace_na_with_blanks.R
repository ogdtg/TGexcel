#' replace_na_with_blanks
#' 
#' Replaces NAs with blanks in Workbook
#'
#' @param wb a workbook
#' @param style_df result of create_style_df
#'
#'
replace_na_with_blanks <- function(wb, style_df) {
  for (i in 1:nrow(style_df)) {
    writeData(
      wb,
      sheet = style_df$sheet[i],
      startRow = style_df$r[i],
      startCol = style_df$c[i],
      x = ""
    )
  }
}