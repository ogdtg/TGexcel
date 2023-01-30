#' create_data_style
#'
#' @param wb a workbook object
#' @param sheet the name of the Sheet in the workbook object
#' @param startRow row in which the data should start
#' @param date index of a date variable (default=NULL)
#' @param year index of a year variable (default=NULL)
#' @param startCol col in which the data should start
#' @param data a data.frame
#' @param datenquelle data source (default = NULL)
#'
#' @export
#'
create_data_style <- function(wb,sheet,startRow, date = NULL,year = NULL, startCol=1, data, datenquelle = NULL, gemeinde = FALSE) {

  style_number <- number_normal_10_right
  style_year <- jahr_normal_10_right

  style_gemeinde_number <- gemeinde_number_normal_10_right
  style_gemeinde_year <- gemeinde_jahr_normal_10_right

  style_gemeinde_year$fontDecoration <- "BOLD"
  style_gemeinde_number$fontDecoration <- "BOLD"

  if (is.null(year)) {
    year = 0
  }

  openxlsx::writeData(x = data,
                      wb = wb,
                      sheet = sheet,
                      startCol = startCol,
                      startRow = startRow,
                      colNames = FALSE)

  a <- lapply(1:ncol(data), function(x) {

    if (year == x) {
      openxlsx::addStyle(wb = wb,
                         sheet = sheet,
                         cols = startCol+x-1,
                         rows = startRow:(startRow+nrow(data)-1),
                         style = style_year)


    } else {
      openxlsx::addStyle(wb = wb,
                         sheet = sheet,
                         cols = startCol+x-1,
                         rows = startRow:(startRow+nrow(data)-1),
                         style = style_number)

    }



  })

  if (gemeinde == TRUE) {
    b <- lapply(1:ncol(data), function(x) {
      if (year == x) {
        openxlsx::addStyle(wb = wb,
                           sheet = sheet,
                           cols = startCol+x-1,
                           rows = gemeinde_format_vec+startRow-1,,
                           style = style_gemeinde_year)



      } else {

        openxlsx::addStyle(wb = wb,
                           sheet = sheet,
                           cols = startCol+x-1,
                           rows = gemeinde_format_vec+startRow-1,
                           style = style_gemeinde_number)
      }

    })
  }



  # Datenquelle
  if (!is.null(datenquelle)) {
    datenquelle_italic_9_left <- datenquelle_italic_9_left

    openxlsx::writeData(x = paste0("Datenquelle: ",datenquelle),
                        wb = wb,
                        sheet = sheet,
                        startCol = startCol,
                        startRow = startRow+nrow(data)+2,
                        colNames = FALSE)

    openxlsx::addStyle(wb = wb,
                       sheet = sheet,
                       cols = startCol,
                       rows = startRow+nrow(data)+2,
                       style = datenquelle_italic_9_left)

  }


}
