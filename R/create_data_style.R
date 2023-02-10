#' create_data_style
#' 
#' Function to insert and style data in the style of the Internettabellen
#'
#' @param wb a workbook object
#' @param sheet the name of the Sheet in the workbook object
#' @param startRow row in which the data should start
#' @param year index of a year variable (default=NULL)
#' @param startCol col in which the data should start
#' @param data a data.frame
#' @param gemeinde_format if true formating for Gemeinde tables will be included (bold for Bezirk and Kanton) (default = FALSE)
#'
#' @export
#'
create_data_style <- function(wb,sheet,startRow,year = NULL, startCol=1, data, gemeinde_format = FALSE) {



  if (is.null(year)) {
    year = 0
  }
  

  openxlsx::writeData(x = data,
                      wb = wb,
                      sheet = sheet,
                      startCol = startCol,
                      startRow = startRow,
                      colNames = FALSE)
  
  for (i in 1:ncol(data)) {
    if (year == i) {
      style <- generic_year_data$copy()
    } else if (is.numeric(data[[i]])) {
      if (any(round(data[i]) != data[i], na.rm = T)) {
        style <- generic_decimal_data$copy()
      } else {
        style <- generic_number_data$copy()
      }
    } else if (class(data[[i]])=="Date") {
      style <- customize_style(template = generic_number_data, date = TRUE)
    } else {
      style <- generic_number_data$copy()
    }
    

    openxlsx::addStyle(wb = wb,
                       sheet = sheet,
                       cols = startCol+i-1,
                       rows = startRow:(startRow+nrow(data)-1),
                       style = style)
  if (gemeinde_format) {
    style_gemeinde <- customize_style(template = style, decoration = "BOLD")
    openxlsx::addStyle(wb = wb,
                       sheet = sheet,
                       cols = startCol+i-1,
                       rows = gemeinde_format_vec+startRow-1,
                       style = style_gemeinde)
    
    openxlsx::setRowHeights(wb=wb, sheet = sheet, rows = gemeinde_format_vec+startRow-1, heights = 30)
    
  }
    
    
    
  }


}

