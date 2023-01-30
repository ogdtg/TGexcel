#' create_varnames_style
#'
#' @param wb a workbook object
#' @param sheet the name of the Sheet in the workbook object
#' @param var_names vector of variable names
#' @param startRow row in which the subheader should be placed (default = 2)
#' @param style an openxlsx::Style element. Default is current "Internettabellen" variable names style
#' @param halign horizontal alignment (default is left)
#' @param ncol only necessary if var_names are NULL (style_only). Provide a positive integer.
#'
#' @export
#'
create_varnames_style <- function(wb, sheet, var_names = NULL,startRow = 3,style = varname_normal_10_left, halign = "left",...) {

  # If necessary, change alignmeent
  if (halign == "right") {
    style[["halign"]] <- "right"
  }
  if (halign == "center") {
    style[["halign"]] <- "center"
  }

  if (is.null(var_names)) {

    if(is.null(ncol)) {
      stop("ncol and var_names cannot be NULL at the same time. Please provide one of them")
    }

    create_style_generic(wb = wb, sheet = sheet,text=NULL, startRow = startRow,style = style,vectorised=TRUE,...)

  } else {
    create_style_generic(wb = wb, sheet = sheet, ncol= length(var_names),text=var_names, startRow = startRow,style = style,vectorised=TRUE,...)
  }

  #
  # openxlsx::addStyle(wb = wb,
  #                    sheet = sheet,
  #                    cols = 1:length(var_names),
  #                    rows = startRow,
  #                    style = style)
  #
  # temp <- lapply(seq_along(var_names), function(i){
  #   openxlsx::writeData(x = var_names[i],
  #                       wb = wb,
  #                       sheet = sheet,
  #                       startCol = i,
  #                       startRow = startRow)
  # })


}
