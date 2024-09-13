#' Create a new Sheet (reiter) for data of the respective year
#'
#' @param xlsx_path path to the excel table where the row should be added
#' @param data data.frame for the respective year
#' @param year year (will later be the name of the sheet)
#' @param dataStart start of the data part of the sheet (below the header, subheader and varnames)
#' @param dataStartCol column index where the data starts. Default is `2`
#'
#' @return workvook object
#' @export
#'
create_xlsx_reiter <- function(xlsx_path,data,year,dataStart,dataStartCol=2){
  
  # Ordner erstellen und Zielpfad generieren
  # new_xlsx_path <- str_replace_all(xlsx_path,paste0(year-1),paste0(year))
  # create_directory(new_xlsx_path)
  
  # Workbook laden
  wb <- loadWorkbook(file = xlsx_path)
  
  
  
  # Neues Sheet erstellen
  openxlsx::cloneWorksheet(wb, sheetName = paste0(year),clonedSheet = paste0(year-1))
  
  # Spalte hinzufügen
  order <- worksheetOrder(wb)
  worksheetOrder(wb) <- c(length(order),1:(length(order)-1))
  
  
  
  writeData(wb,sheet = paste0(year),x = data,startCol = dataStartCol,startRow = dataStart,colNames = F)
  
  activeSheet(wb) <- paste0(year)
  
  # Untertitel updaten
  update_untertitel(wb, year)
  
  
  
  
  #abspeichern
  return(wb)
  # save_tg_workbook(wb,new_xlsx_path,tg_header = F,overwrite = T)
  
  
}