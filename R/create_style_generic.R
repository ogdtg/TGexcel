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
#' @param startCol col in which the function should start (default = 1)
#'
#' @export
#' 
create_style_generic <- function(wb, sheet, ncol,text, startRow, style, vectorised = FALSE, style_only = FALSE, startCol = 1) {



  if (style_only == FALSE) {
    if (vectorised==TRUE) {
      temp <- lapply(seq_along(text), function(i){

        startCol_mod = startCol + i -1

        if (grepl("\\[.*\\]", text[i]) | grepl("\\~.*\\~", text[i])) {
          addSuperSubScriptToCell(wb=wb,
                                  sheet = sheet,
                                  col = startCol_mod,
                                  row = startRow,
                                  texto = text[i],
                                  style = style)
        } else {

          openxlsx::writeData(x = text[i],
                              wb = wb,
                              sheet = sheet,
                              startCol = startCol_mod,
                              startRow = startRow)

        }


      })

    } else {

      if (grepl("\\[.*\\]", text) | grepl("\\~.*\\~", text)) {
        addSuperSubScriptToCell(wb=wb,
                                sheet = sheet,
                                col = startCol,
                                row = startRow,
                                texto = text,
                                style = style)
      } else {
        openxlsx::writeData(x = text,
                            wb = wb,
                            sheet = sheet,
                            startCol = startCol,
                            startRow = startRow)


      }
    }
  }
  openxlsx::addStyle(wb = wb,
                     sheet = sheet,
                     cols = startCol:(ncol+startCol-1),
                     rows = startRow,
                     style = style)
}
