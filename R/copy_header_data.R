#' copy_header_data
#' 
#' Copy the header without formating from one sheet to another.
#'
#' @param wb a workbook object
#' @param sheet the name of the Sheet in the workbook object from which to copy the data
#' @param newSheet new sheet where the data should be copied
#' @param dataStartRow row where the body data starts
#'
#' @export
#'
copy_header_data <- function(wb,sheet,newSheet, dataStartRow) {
  
  if (!newSheet %in% openxlsx::sheets(wb)){
    addWorksheet(wb, newSheet)
    message(paste0("Write to new sheet ",newSheet))
  } else {
    message(paste0("Write to existing sheet ",newSheet))
  }
  
  invisible(lapply(1:(dataStartRow-1),function(x){
    temp_df <- read.xlsx(wb, sheet = sheet,skipEmptyRows = F, skipEmptyCols = F ,colNames = F, rows = x)
    writeData(wb, sheet = newSheet, startRow = x, colNames = F, x = temp_df)
    
  }))
}