#' copy_heights_widths
#' 
#' Copy colWidths and rowHeights from one sheet to another one
#'
#' @param wb a workbook object
#' @param sheet the name of the Sheet in the workbook object from which to copy the row heights and col widths
#' @param newSheet new sheet where the row heights and col widths should be inserted
#'
#' @export
#'
copy_heights_widths <- function(wb,sheet,newSheet) {
  
  if (!newSheet %in% openxlsx::sheets(wb)){
    addWorksheet(wb, newSheet)
    message(paste0("Write to new sheet ",newSheet))
  } else {
    message(paste0("Write to existing sheet ",newSheet))
  }
  
  index <- which(sheets(wb) == sheet)
  
  rH <- wb$rowHeights[[index]]
  cW <- wb$colWidths[[index]]
  
  row_names <- names(rH)
  row_heights <- rH
  
  col_names <- names(cW)
  col_widths <- cW
  
  
  for (i in 1:length(col_names)) {
    openxlsx::setColWidths(wb, newSheet, cols = col_names[i], widths = as.numeric(col_widths[i]))
  }
  
  for (k in 1:length(row_names)) {
    openxlsx::setRowHeights(wb, newSheet, rows = row_names[k], heights = as.numeric(row_heights[k]))
  }
  
}