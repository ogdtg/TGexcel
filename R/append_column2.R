#' append_column (modified)
#' 
#' Function to append new column at the end of a table
#'
#' @param wb a workbook object
#' @param sheet sheet name
#' @param patternCol column from which the styles will be copied (default is NULL -> last column of table)
#' @param col column index where to insert data
#' @param startRow row index where to start copying the styles (default = 1)
#' @param endRow row index where to end copying the styles (default is NULL -> endrow = nrow(data)+dataStart-1)
#' @param dataStart row index where the data starts
#' @param data data.frame (ncol must be 1) to insert
#' @param colNames if TRUE, colNames of data will be included (default is FALSE)
#' 
append_column2 <-
  function (wb,
            sheet,
            patternCol = NULL,
            col,
            startRow = 1,
            endRow = NULL,
            dataStart,
            data,
            colNames = FALSE) {
    if (is.null(patternCol)) {
      patternCol <- col - 1
    }
    if (is.null(endRow)) {
      endRow <- nrow(data) + dataStart - 1
    }
    styles <- get_styles_of_area(wb, sheet, cols = patternCol,
                                 rows = startRow:endRow)
    if (colNames) {
      dataStart <- dataStart - 1
    }
    writeData(
      wb,
      sheet,
      x = data,
      startCol = col,
      startRow = dataStart,
      colNames = colNames
    )
    invisible(lapply(styles, function(x) {
      addStyle(
        wb,
        sheet,
        style = x$style,
        rows = x$rows,
        cols = col
      )
    }))
    
    return(wb)
  }