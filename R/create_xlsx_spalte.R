#' Append a new column for a new year
#'
#' @param xlsx_path path to the excel table where the row should be added
#' @param data data.frame for the respective year. The column that should be inserted must be named after the year (e.g. 2024). Can also be a list of data.frames, if multiple sheets should be added.
#' @param year year of the data
#' @param dataStart start of the data part of the sheet (below the header, subheader and varnames)
#' @param limit all values below this value will be replaced with `-`
#' @param sheets range of the sheets that should be edited. Default is NULL, e.g. first sheet will be edited.
#'
#' @return workbook object
#' @export
#'
create_xlsx_spalte <- function(xlsx_path,data,year,dataStart,limit = NULL,sheets =NULL){
  
  # Ordner erstellen und Zielpfad generieren
  # new_xlsx_path <- str_replace_all(xlsx_path,paste0(year-1),paste0(year))
  # create_directory(new_xlsx_path)
  
  # Workbook laden
  wb <- loadWorkbook(file = xlsx_path)
  
  # Spalte hinzufügen
  
  if (!is.null(sheets)){
    for (sheet in sheets){
      wb <- add_data_spalte_jahr(wb, year=year,data = data[[sheet]],dataStart = dataStart,limit=limit, sheet = sheet)
      # Untertitel updaten
      update_untertitel(wb,year,sheet = sheet)
    }
  } else {
    wb <- add_data_spalte_jahr(wb, year=year,data = data,dataStart = dataStart,limit=limit)
    
    # Untertitel updaten
    update_untertitel(wb,year)
    
    
    # Worksheet Namen updaten
    if (str_detect(sheets(wb)[1],paste0(year-1))) {
        rename_sheet(wb,year)
      
    }
    
  }
  
  #abspeichern
  # save_tg_workbook(wb,new_xlsx_path,tg_header = F,overwrite = T)
  return(wb)
  
}