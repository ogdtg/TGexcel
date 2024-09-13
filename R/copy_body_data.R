#' copy_body_data
#' 
#' Copy the data body from one sheet to another
#'
#' @param wb a workbook object
#' @param sheet the name of the Sheet in the workbook object from which to copy the data
#' @param newSheet new sheet where the data should be copied
#' @param dataStartRow row where the body data starts
#' @param year index of the column where a year variable is located
#'
#' @export
#'
copy_body_data <- function(wb,sheet,newSheet, dataStartRow, year) {
  
  if (!newSheet %in% openxlsx::sheets(wb)){
    addWorksheet(wb, newSheet)
    message(paste0("Write to new sheet ",newSheet))
  } else {
    message(paste0("Write to existing sheet ",newSheet))
  }
  
  data <- read.xlsx(wb,sheet = sheet, startRow = dataStartRow, skipEmptyRows = F, colNames = F)
  create_data_style(wb,sheet = newSheet, startRow = dataStartRow, year = year, data = data)
}