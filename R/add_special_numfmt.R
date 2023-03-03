#' add_special_numfmt
#' 
#' Function writes data to workbook and adds special numFmt for matching cells
#'
#' @param wb a workbook object
#' @param sheet name of the sheet
#' @param data a data.frame. should only conatin the numbers you want to write to the excel but no leading column with string information
#' @param pattern a regular expression that should be looked for in the data
#' @param startCol where the data insert should start
#' @param dataStart row in which the numeric data in the file starts (e.g. first row after header)
#' @param ... see `special_numFmt`

#'
#' @export
#'
add_special_numfmt <- function(wb,sheet,data, pattern,startCol,dataStart,...) {
  index_df <- create_index_df(data = data, pattern = pattern) #index_df aus Rohdaten
  data <- data %>% #Rohdaten zu numerc
    dplyr::mutate_all(.funs = function(x){str_remove_all(x,pattern)}) %>% 
    dplyr::mutate_all( .funs = as.numeric)
  openxlsx::writeData(wb=wb,sheet=sheet,x = data,startRow = dataStart, 
                      startCol = startCol, colNames = FALSE) #Daten ins File schreiben
  if (nrow(data)!=0){
    special_numFmt(wb=wb,sheet=sheet,index_df = index_df,dataStart=dataStart,colShift =startCol-1,...) #Special formats anwenden 
  } else {
    warning("No cell matches the pattern. No special numFmt will be applied.")
  }
  
}
