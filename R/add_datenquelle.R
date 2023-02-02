#' add_datenquelle
#' 
#' Function to add datenwuelle
#'
#' @param wb a workbook object
#' @param sheet the name of the Sheet in the workbook object
#' @param startRow row where to start
#' @param col column where to insert
#' @param datenquelle 
#'
#' @export
#'
add_datenquelle <- function(wb,sheet,startRow, col = 1, datenquelle){


  create_style_generic(
    wb,
    sheet,
    startRow = startRow,
    ncol = col,
    style = datenquelle_italic_9_left,
    startCol = col,
    text = paste0("Datenquelle: ", datenquelle)
  )
  
}