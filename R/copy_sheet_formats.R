#' copy_sheet_formats
#' 
#' Copy all styles from one sheet to another one
#'
#' @param wb a workbook object
#' @param sheet the name of the Sheet in the workbook object from which to copy the styles
#' @param newSheet new sheet where the styles should be copied
#'
#' @export
#'
copy_sheet_formats <- function(wb,sheet, newSheet){
  
  if (!newSheet %in% openxlsx::sheets(wb)){
    addWorksheet(wb, newSheet)
    message(paste0("Write to new sheet ",newSheet))
  } else {
    message(paste0("Write to existing sheet ",newSheet))
  }
  
  
  styles <- lapply(wb$styleObjects, function(x) {
    if (x$sheet == sheet) {
      return(x)
    }
  })
  styles <- styles[lengths(styles) != 0]
  
  invisible(lapply(styles, function(x){
    addStyle(wb,
             sheet = newSheet,
             rows = x$rows,
             cols = x$cols,
             style = x$style$copy())
  }))
}