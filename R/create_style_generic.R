#' create_style_generic
#'
#' @param wb a workbook object
#' @param sheet the name of the Sheet in the workbook object
#' @param ncol the number of columns
#' @param text a string or a vector
#' @param startRow row in which the content should be placed
#' @param style an openxlsx::Style element
#' @param vectorised if a vector is given for text, vectorised must be set to TRUE (default = FALSE)
#' @param style_only if set to TRUE, text will be ignored and only the style will be set (default = FALSE)
#'
create_style_generic <- function(wb, sheet, ncol,text, startRow, style, vectorised = FALSE, style_only = FALSE) {

  openxlsx::addStyle(wb = wb,
                     sheet = sheet,
                     cols = 1:ncol,
                     rows = startRow,
                     style = style)

  if (style_only == FALSE) {
    if (vectorised==TRUE) {
      temp <- lapply(seq_along(text), function(i){
        openxlsx::writeData(x = text[i],
                            wb = wb,
                            sheet = sheet,
                            startCol = i,
                            startRow = startRow)
      })

    } else {
      openxlsx::writeData(x = text,
                          wb = wb,
                          sheet = sheet,
                          startCol = 1,
                          startRow = startRow)

    }
  }




}
