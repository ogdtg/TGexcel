#' Append a data row with the respective year
#'
#' @param xlsx_path path to the excel table where the row should be added
#' @param data the data.frame. Must contain one row with a column called "jahr" where the respective year is stored. Function wil only take the row where jahr==year. Can also be a list of data.frames if data should be added in multiple sheets
#' @param year the year of the data
#' @param dataStart start of the data part of the sheet (below the header, subheader and varnames)
#' @param sheets if a list of data.frames is given, the range of sheet indexes can be passed to the function here (e.g 1:3). Default is `NULL`
#'
#' @return a workbook object
#' @export
#'
create_xlsx_zeile <- function(xlsx_path,data,year,dataStart,sheets =NULL){
  
  # Ordner erstellen und Zielpfad generieren
  # new_xlsx_path <- str_replace_all(xlsx_path,paste0(year-1),paste0(year))
  # create_directory(new_xlsx_path)
  
  # Workbook laden
  wb <- loadWorkbook(file = xlsx_path)
  
  
  if (!is.null(sheets)){
    for (sheet in sheets){
      wb <- add_data_zeile_jahr(wb, year=year,data = data[[sheet]],dataStart = dataStart, sheet = sheet)
      # Untertitel updaten
      update_untertitel(wb,year,sheet = sheet)
    }
  } else {
    wb <- add_data_zeile_jahr(wb, year=year,data = data,dataStart = dataStart)
    
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
