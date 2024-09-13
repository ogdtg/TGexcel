#' Update subtitle to given year
#'
#' @param wb a workbook object
#' @param year the year of the data
#' @param sheet sheet number
#'
#'
update_untertitel <- function(wb,year,sheet = 1){
  data <- read.xlsx(wb)
  
  untertitel <- data[1,1]
  
  new_untertitel <- str_replace(untertitel,as.character(year-1),as.character(year)) |>
    get_superscript_numbers()
  
  create_subheader_style(wb,sheet = sheets(wb)[sheet],ncol = ncol(data),text = new_untertitel)
  
}


#' Rename sheet to given year
#'
#' @param wb a workbook object
#' @param year the year of the data
rename_sheet <- function(wb,year){
  
  sheet = sheets(wb)[1]
  new_name <- str_replace(sheet,"\\d\\d\\d\\d$",as.character(year))
  
  renameWorksheet(wb,sheet, new_name)
}



#' Get superscript numbers from text
#'
#' @param text  text
#'
get_superscript_numbers <- function(text){
  # Regex pattern to match subscripts (a digit at the end of a number or text)
  pattern <- "(\\d{4})(\\d)|([a-zA-Zäöüß]+)(\\d)"
  
  # Replace subscripts with [subscript_num]
  # \\2 and \\4 capture the subscript numbers from the regex groups
  gsub(pattern, "\\1\\3[\\2\\4]", text)
}


#' Replace small values for anonymisation
#'
#' @param wb a workbook object
#' @param data a data.frame with the data you want to insert 
#' @param limit all values smaller than this will be repalced with the replacement
#' @param col index of the col
#' @param replacement replacement. Default is "-"
#' @param year year of the data
#' @param dataStart start of the data part of the sheet (below the header, subheader and varnames)
#' @param sheet sheet number
#'
#' @return workbook object
#'
replace_small_values <- function(wb,data,limit,col,replacement = "-",year,dataStart,sheet = 1){
  
  index <- which(data[[paste0(year)]]<limit)
  
  for (i in index){
    writeData(wb,sheet = sheet,startCol = col,startRow = i+(dataStart-1),colNames = F,x = replacement)
  }
  
  return(wb)
  
}


#' Create directory from file path
#'
#' @param xlsx_path file path
#'
create_directory <- function(xlsx_path){
  dir_path <- dirname(xlsx_path)
  
  # Create the directory, suppress warning if it already exists
  suppressWarnings({
    if (!dir.exists(dir_path)) {
      dir.create(dir_path, recursive = TRUE)
    }
  })
}