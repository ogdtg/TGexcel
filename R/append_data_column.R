#' append_data_column
#' 
#' Append column to the end of the data
#'
#' @param wb a workbook object
#' @param sheet the name of the Sheet in the workbook object
#' @param varname name of the new variable 
#' @param data a data.frame with the data you want to insert 
#' @param dataStart row where the data starts
#'
#' @export
#'
append_data_column <- function(wb,sheet,varname,data , dataStart) {
  
  numRows <- nrow(data)
  numCols <- ncol(data)
  
  wb_data <- openxlsx::read.xlsx(wb, skipEmptyRows = F, colNames = F, sheet= sheet)
  nrow_wb <- nrow(wb_data)
  ncol_wb <- ncol(wb_data)
  
  
  writeData(wb,sheet, startCol = ncol_wb+1, startRow = dataStart-1,x=varname)
  
  create_data_style(wb,sheet,startRow = dataStart,startCol = ncol_wb+1, data = data)
  
  style_col <- get_styles_of_area(wb,sheet,rows = 1:nrow_wb,cols = ncol_wb)
  
  invisible(lapply(style_col, function(x){
    addStyle(wb,
             sheet,
             style = x$style$copy(),
             rows = x$rows,
             cols = x$cols+1)
  }))
  
  
  
}