#' create_subheader_style
#'
#' @param wb a workbook object
#' @param sheet the name of the Sheet in the workbook object
#' @param ncol the number of columns that the subheader should be long
#' @param text the text of the subheader
#' @param startRow row in which the subheader should be placed (default = 2)
#' @param style an openxlsx::Style element. Default is current "Internettabellen" subheader style
#' @param rowHeight height of the row(s)
#' @param ... see create_style_generic
#'
#' @export
#'
create_subheader_style <- function(wb, sheet, ncol,text, startRow = 2, style = subheader_normal_10_left,rowHeight = 15,...) {

  create_style_generic(wb = wb, sheet = sheet, ncol= ncol,text=text, startRow = startRow,style = style,...)
  openxlsx::setRowHeights(wb,sheet,rows=startRow,heights = rowHeight)
  

}
