#' add_decoration
#' 
#' Adds a font decoration to specified cells
#'
#' @param wb a workbook object
#' @param sheet sheet name
#' @param rows rows the style should be added to
#' @param cols columns the style should be added to
#' @param decoration see [openxlsx::createStyle()]
#'
#' @export
#'
add_decoration <- function(wb,sheet,rows,cols,decoration){
  
  decoration <- tolower(decoration)
  
  if (!decoration %in% c("bold", "strikeout", 
                         "italic", "underline", "underline2", "accounting", 
                         "accounting2", "")){
    stop("Invalid decoration name")
  } else {
    feature_values <- c(toupper(decoration))
  }
  
  feature_names <- c("fontDecoration")
  
  
  add_style_feature(wb,sheet, rows, cols = cols, feature_names = feature_names, feature_values = feature_values)
}