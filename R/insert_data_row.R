#' insert_data_row
#' 
#' Function to insert Data (similar to insert Row in Excel).
#'
#' @param wb a workbook object
#' @param sheet the name of the Sheet in the workbook object
#' @param data a data.frame with the data you want to insert 
#' @param insertRow the index of the row the inserted data should start (styles will be taken from this particular row)
#' @param dataStartRow start of the data part of the sheet (below the header, subheader and varnames)
#'
#' @export
#'
insert_data_row <- function(wb,sheet,data, insertRow, dataStartRow){
  
  
  
  numRows <- nrow(data)
  
  wb_data <- openxlsx::read.xlsx(wb, skipEmptyRows = F, colNames = F, sheet= sheet)
  nrow_wb <- nrow(wb_data)
  ncol_wb <- ncol(wb_data)
  
  # Remove all merges below the header, otherwise styles cannot be apllied correctly
  removeCellMerge(wb, sheet,cols =1:ncol_wb ,rows = dataStartRow:nrow_wb)
  
  # styles above the insertrow
  styles_top <- get_styles_of_area(wb,sheet, rows = insertRow-1,cols = 1:ncol_wb)
  
  #Styles below the insertrow
  styles_bottom <- get_styles_of_area(wb,sheet, rows = insertRow:nrow_wb,cols = 1:ncol_wb)
  
  #Data below the insertrow
  data_bottom <- wb_data[insertRow:nrow_wb,1:ncol_wb]
  
  
  # Data will be moved so that there is space for the data to insert
  for (i in 1:ncol(data_bottom)) {
    temp = suppressWarnings(as.numeric(data_bottom[,i]))
    na_num <- sum(is.na(temp))
    na_char <- sum(is.na(data_bottom[,i]))
    
    if(na_num != na_char){
      create_data_style(wb,sheet,startRow = insertRow+numRows,startCol = i, data = data.frame(x1 =data_bottom[,i]))
    } else {
      
      create_data_style(wb,sheet,startRow = insertRow+numRows,startCol = i, data = data.frame(x1 = temp))
    }
  }
  
  # Apply styles to moved data
  invisible(lapply(styles_bottom, function(x){
    addStyle(wb,
             sheet,
             style = x$style$copy(),
             rows = x$rows+numRows,
             cols = x$cols)
  }))
  
  # Insert the data
  create_data_style(wb,sheet, startRow = insertRow, data = data)
  
  # apply styles from top
  invisible(lapply(styles_top, function(x){
    addStyle(wb,
             sheet,
             style = x$style$copy(),
             rows = suppressWarnings((x$rows+1):(x$rows+numRows)),
             cols = x$cols,
             gridExpand = T)
  }))
  
}