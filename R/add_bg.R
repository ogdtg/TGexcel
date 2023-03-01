#' add_bg
#' 
#' Adds fill color to specified cells
#'
#' @param wb a workbook object
#' @param sheet sheet name
#' @param rows rows the style should be added to
#' @param cols columns the style should be added to
#' @param color fillcolor of the specified cells (Rcolor or HEX Value)
#'
#' @export
#'
add_bg <- function(wb,sheet,rows,cols,color){
  feature_names <- c("fillFg")
  
  
  color <- openxlsx:::validateColour(color)
  
  add_style_feature(wb,sheet, rows, cols = cols, feature_names = feature_names, feature_values = list(list(rgb = color)))
}