#' add_hline
#' 
#' Adds a horizontal line by setting cell borders
#'
#' @param wb a workbook object
#' @param sheet sheet name
#' @param rows rows the the line should be placed in
#' @param ncol number of columns the line should embrace
#' @param color color of the line (Rcolor or HEX Value)
#' @param style see [openxlsx::createStyle()]
#' @param below logical whether line should be placed below or above selected rows (default is `TRUE`)
#' @param startCol column where the line should start (default is 1)
#'
#' @export
#'
add_hline <- function(wb,sheet,rows,ncol,color,style = "thin", below = TRUE, startCol = 1){
  if(below == TRUE) {
    feature_names <- c("borderBottom","borderBottomColour")
  } else {
    feature_names <- c("borderTop","borderTopColour")
  }
  
  color <- openxlsx:::validateColour(color)
  
  add_style_feature(wb,sheet, rows, cols = 1:ncol, feature_names = feature_names, feature_values = c(style,list(rgb = color)))
}



