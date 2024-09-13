#' Function to append a year as column
#' 
#' @description Appends a data column where the year has to be the name of the column
#'
#' @param wb a workbook object
#' @param year the year of the data
#' @param data A dataset with one column where the name is the respective year. Can also be a list of datasets
#' @param dataStart row index where the data starts
#' @param limit limit for small values. Values below this threshold will be marked with "-"
#' @param sheet sheet number. default is 1. If more than one sheet, a range can be passed to the function
#'
#' @return workbook object
#' @export
#'
add_data_spalte_jahr <- function(wb,year,data,dataStart,limit = NULL,sheet = 1){
  
  # Aktuelle Daten lesen um korrekte Spalte zum einfügen zu ermitteln
  current_data <- openxlsx::read.xlsx(wb,sheet=sheet)
  
  # Spalte aus der die Styles kommen ist die letzte Dtaenspalte
  patterncol = ncol(current_data)
  
  # Daten auf eine Spalte (Jahr) reduzieren
  data_mod <- data |>
    dplyr::select(all_of(as.character(year)))
  
  wb <- append_column2(wb,
                       sheet = sheets(wb)[sheet],
                       data = data_mod,
                       col = patterncol+1,
                       dataStart = dataStart,
                       patternCol = patterncol,
                       colNames = T)
  
  
  col_widths <- wb$colWidths[[sheet]]
  
  setColWidths(wb,sheet=sheet,widths =col_widths[length(col_widths)], cols = patterncol+1)
  
  if (!is.null(limit)){
    wb <- replace_small_values(wb,data = data_mod,limit = limit,col = patterncol+1,year = year,dataStart=dataStart,sheet = sheet)
  }
  return(wb)
}



#' Function to append a year as row
#' 
#' @description Appends a data row where the the year has to be in one variable "jahr" 
#'
#' @param wb a workbook object
#' @param year the year of the data
#' @param data A dataset with one row where the column "jahr" contains the respective year. Can also be a list of datasets
#' @param dataStart row index where the data starts
#' @param sheet sheet number. default is 1. If more than one sheet, a range can be passed to the function
#'
#' @return workbook object
#' @export
add_data_zeile_jahr <- function(wb,year,data,dataStart,sheet = 1){
  
  
  # Aktuelle Daten lesen um korrekte Spalte zum einfügen zu ermitteln
  current_data <- openxlsx::read.xlsx(wb,sheet=sheet,skipEmptyRows = F)
  
  if (year %in% current_data[[1]]){
    stop(paste0(year," bereits in den Daten vorhanden."))
  }
  
  # Position der
  insertRow <- which(current_data[[1]]==year-1)+2
  
  
  #
  data_mod <- data |>
    dplyr::filter(jahr == year)
  
  insert_data_row2(wb=wb,
                   sheet = sheets(wb)[sheet],
                   data = data_mod,
                   insertRow=insertRow,
                   dataStartRow=dataStart  )
  
  
}