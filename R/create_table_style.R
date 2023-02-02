#' create_table_style
#' 
#' Function to create a complete table in excel in the format of the Internettabellen
#'
#' @param wb a workbook object
#' @param sheet the name of the Sheet in the workbook object
#' @param header text for header
#' @param subheader text for subheader
#' @param varnames variable names
#' @param var_categories variable categories for nested header (nesting must be a numeric vector or 'even', otherwise var_categories will be ignored)
#' @param data a data.frame
#' @param datenquelle data source (default = NULL)
#' @param footnotes a character vector with footnotes. Please provide the correct order and do not forget the corresponding number (example: '1 Lorem Ipsum'). default = NULL
#' @param startRow row in which the header should be placed (default = 1)
#' @param gemeinde_format if TRUE formating for Gemeinde tables will be included (bold for Bezirk and Kanton) (default = FALSE)
#' @param nesting if the variables should be evenly distributed, then provide the string "even". If the variables should NOT be evenly distributed, please provide a numeric vector with the same length as var_categories The given integers indicate how many variables from level 2 are nested below one variable from level 1 (default=NULL)
#' @param header_style style for header (default = header_bold_12_left)
#' @param subheader_style style for subheader (default = subheader_normal_10_left )
#' @param varnames_style style for varnames (default = varname_normal_10_left )
#' @param var_categories_style style for var_categories (default = varname_normal_10_left )
#'
#' @export
#'
create_table_style <- function(wb,
                               sheet,
                               header,
                               subheader,
                               varnames,
                               var_categories = NULL,
                               data,
                               datenquelle = NULL,
                               footnotes = NULL,
                               startRow = 1,
                               gemeinde_format = FALSE,
                               nesting = NULL,
                               header_style = header_bold_12_left,
                               subheader_style = subheader_normal_10_left,
                               varnames_style = varname_normal_10_left,
                               var_categories_style = varname_normal_10_left) {
  if (!is.null(nesting) & is.null(var_categories)) {
    stop("If nesting != NULL, var_categories must be provided.")
  }
  if (is.null(nesting) & !is.null(var_categories)) {
    message(
      "nested == NULL. var_categories will be ignored. Please provide a numeric vector or 'even' if you want nested variables."
    )
  }
  if (length(varnames) != ncol(data)) {
    stop(
      paste0("Varnames has the length ",length(varnames),", data has ",ncol(data)," columns.")
    )
  }

  if (nrow(data) != nrow(gemeinde)) {
    message(paste0("data has ",nrow(data)," rows, gemeinde data has ",nrow(gemeinde)))
  }
  
  create_header_style(
    wb,
    sheet = sheet,
    text = header,
    ncol = ncol(data),
    style = header_style,
    startRow = startRow
  )
  create_subheader_style(
    wb,
    sheet = sheet,
    text = subheader,
    ncol = ncol(data),
    style = subheader_style,
    startRow = startRow + 1
  )
  
  if (!is.null(nesting)) {
    create_nested_header_style(
      wb,
      sheet = sheet,
      vars_level1 = var_categories,
      vars_level2 = varnames,
      nesting = nesting,
      startRow = startRow + 2,
      vars_level1_style = var_categories_style,
      vars_level2_style = varnames_style
    )
    start_data = startRow + 4
  } else {
    create_varnames_style(wb,
                          sheet = sheet,
                          var_names = varnames,
                          style = varnames_style)
    start_data = startRow + 3
    
  }
  create_data_style(
    wb,
    sheet = sheet,
    startRow = start_data,
    data = data,
    datenquelle = datenquelle,
    gemeinde_fomat = gemeinde_format
  )
  
  # Footnotes
  
  if(length(footnotes)> 0) {
    footnote_style <- customize_style(template = generic_number_data, fontSize = 9) 

    footer <- lapply(seq_along(footnotes), function(i) {
      create_style_generic(wb,sheet, startRow = start_data+nrow(data)+i,ncol=1, style = footnote_style, startCol = 1, text = footnotes[i])
  
    })
  }
  

  
  # Datenquelle
  if (!is.null(datenquelle)) {
    
    create_style_generic(wb,sheet, startRow = start_data+nrow(data)+length(footnotes)+2,ncol=1, style = datenquelle_italic_9_left, startCol = 1, text = paste0("Datenquelle: ",datenquelle))
    
  }
  
}