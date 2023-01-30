#' create_header_style
#'
#' @param wb a workbook object
#' @param sheet the name of the Sheet in the workbook object
#' @param ncol the number of columns that the header should be long
#' @param text the text of the header
#' @param startRow row in which the header should be placed (default = 1)
#' @param style an openxlsx::Style element. Default is current "Internettabellen" header style
#'
#' @export
#'
create_header_style <- function(wb, sheet, ncol,text, startRow = 1,style = header_bold_12_left,...) {
  
  create_style_generic(wb = wb, sheet = sheet, ncol= ncol,text=text, startRow = startRow,style = style,...)


}
