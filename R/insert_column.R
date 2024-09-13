#' insert_column
#' 
#' Function to insert a data column somewhere in the table.
#'
#' @param wb a workbook object
#' @param sheet sheet name
#' @param patternCol column from which the styles will be copied (default is NULL -> last column of table)
#' @param col column index where to insert data
#' @param endCol column where to end style copying (default = 200)
#' @param startRow row index where to start copying the styles (default = 1)
#' @param endRow row index where to end copying the styles (default is NULL -> endrow = nrow(data)+dataStart-1)
#' @param dataStart row index where the data starts
#' @param data data.frame (ncol must be 1) to insert
#' @param colNames if TRUE, colNames of data will be included (default is FALSE)
#'
#' @export
#'
insert_column <- function(wb,sheet, patternCol=NULL, col,endCol = NULL, startRow = 1, endRow = NULL,dataStart,data, colNames = FALSE){
  
  if (is.null(patternCol)) {
    patternCol <- col +1
  }
  if (is.null(endRow)){
    endRow <- nrow(data)+dataStart-1
  }
  styles_insert <- get_styles_of_area(wb,sheet,cols = patternCol, rows = startRow:endRow)
  
  datastart_ex <- dataStart-1
  
  if (colNames) {
    dataStart_char <- dataStart
    dataStart <- datastart_ex
    
  }
  
  if (is.null(endCol)) {
    endCol <- 200
  }
  
  styles_existing <- get_styles_of_area(wb,sheet,cols = col:endCol, rows = startRow:endRow)
  df_existing <- openxlsx::readWorkbook(wb,sheet, rows = startRow:endRow, cols = col:endCol, skipEmptyRows = FALSE,  rowNames = FALSE)
  
  df_existing_char <- df_existing
  df_existing_char[df_existing_char>0] <- NA
  
  char_vals <- invisible(lapply(seq_along(df_existing_char), function(i){
    vec <- which(!is.na(df_existing_char[,i]))
    rows <- vec
    cols <- rep(i,length(vec))
    if (length(rows)==0){
      return(NULL)
    }
    return(data.frame(rows = rows, cols = cols, value = df_existing[rows,i]))
  }))
  char_df <- char_vals %>% 
    dplyr::bind_rows()
  
  
  
  df_existing <-
    df_existing %>% 
    dplyr::mutate_all(as.numeric) %>% 
    suppressWarnings()
  
  writeData(wb,
            sheet,
            startRow = datastart_ex,
            startCol = col+1,
            x = df_existing,
            colNames = TRUE)
  
  invisible(apply(char_df, 1, function(x){
    writeData(wb,
              sheet,
              startRow = as.numeric(x["rows"]) +dataStart_char-1,
              startCol = as.numeric(x["cols"]) + col,
              x = x["value"])
  }))
  
  invisible(lapply(styles_existing, function(x) {
    addStyle(wb,
             sheet,
             style = x$style,
             rows = x$rows,
             cols = x$cols+1)
  }))
  
  writeData(wb,
            sheet,
            startRow = dataStart,
            startCol = col,
            x = data,
            colNames = colNames)
}