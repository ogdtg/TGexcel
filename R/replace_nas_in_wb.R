#' replace_nas_in_wb
#' 
#' Replaces NAs in Workbook so that they do not appear in the Header of a Table
#'
#' @param wb a workbook object
#'
#' @export
#'
replace_nas_in_wb <- function(wb) {
  style_df <- create_style_df(wb)
  replace_na_with_blanks(wb,style_df)
  
}