#' create_varnames_style
#' 
#' Function to insert and style variable names in the style of the Internettabellen
#'
#' @param wb a workbook object
#' @param sheet the name of the Sheet in the workbook object
#' @param var_names vector of variable names
#' @param startRow row in which the subheader should be placed (default = 2)
#' @param style an openxlsx::Style element. Default is current "Internettabellen" variable names style
#' @param ncol only necessary if var_names are NULL (style_only). Provide a positive integer.
#'
#' @export
#'
create_varnames_style <- function(wb, sheet, var_names = NULL,startRow = 3,style = varname_normal_10_left,...) {


  if (is.null(var_names)) {

    if(is.null(ncol)) {
      stop("ncol and var_names cannot be NULL at the same time. Please provide one of them")
    }

    create_style_generic(wb = wb, sheet = sheet,text=NULL, startRow = startRow,style = style,vectorised=TRUE,...)

  } else {
    create_style_generic(wb = wb, sheet = sheet, ncol= length(var_names),text=var_names, startRow = startRow,style = style,vectorised=TRUE,...)
  }


}
