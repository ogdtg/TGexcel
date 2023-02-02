#' write_styles_to_wb
#' 
#' Functions to see which cell styles are present in the workbook
#'
#' @param wb a workbook object
#' @param sheet the name of the Sheet in the workbook object
#' @param filename where to store the resulting xlsx file
#'
#' @export
#'
write_styles_to_wb <- function(wb,sheet,filename) {
  
  filename = normalizePath(filename)
  
  new_wb <- openxlsx::createWorkbook()
  addWorksheet(new_wb,sheetName = sheet)
  
  df <- data.frame(text="das ist Text",dec=45.834567890, year = 2018, number =2123412, date = Sys.Date())
  
  lapply(seq_along(wb$styleObjects), function(i) {
    
    writeData(new_wb,sheet,startCol = 1,startRow = i, x= df, colNames = FALSE)
    addStyle(new_wb, sheet, rows = i,cols = 1:ncol(df), gridExpand = TRUE, style = wb$styleObjects[[i]]$style)
  })
  
  saveWorkbook(new_wb, filename, overwrite = TRUE)
}
