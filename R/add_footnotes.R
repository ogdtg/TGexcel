#' add_footnotes
#' 
#' Function to add footnodes
#'
#' @param wb a workbook object
#' @param sheet the name of the Sheet in the workbook object
#' @param startRow row where to start
#' @param col column where to insert
#' @param footnotes character vector or string with footnotes
#'
#' @export
#'
add_footnotes <- function(wb,sheet,startRow,col = 1,footnotes) {
  footnote_style <- customize_style(template = generic_number_data, fontSize = 9) 
  
  footer <- lapply(seq_along(footnotes), function(i) {
    create_style_generic(wb,sheet, startRow = startRow+i-1,ncol=col, style = footnote_style, startCol = col, text = footnotes[i])
    
  })
}