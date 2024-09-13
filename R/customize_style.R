#' customize_style
#' 
#' Function to customize a style for the common Internettabellen
#'
#' @param template can either be "header","subheader","varname", "data_year", "data_decimal" or an object of class openxlsx::Style. Default is Arial, 10, no decoration, no alignment, with thousands separator
#' @param halign horizontal alignment ("left", "right", "center" or "justify"). Default = NULL
#' @param valign vertical alignment ("top", "bottom" or "center"). Default = NULL
#' @param decoration font decoration ("bold", "strikeout", "italic", "underline", "underline2", "accounting" or"accounting2"). Default = NULL
#' @param fontName name of the Font. Default = "Arial"
#' @param fontSize font size in pt. Default = 10
#' @param date whether format should represent a date (default = FALSE)
#' @param dateformat formating of dates (default = 'dd.mm.yyyy')
#'
#' @return customized style of template openxlsx::Style
#' @export
#'
customize_style <- 
  function(template = NULL,
           halign = NULL,
           valign = NULL,
           decoration = NULL,
           fontName = NULL,
           fontSize = NULL,
           date = FALSE,
           dateformat = "dd.mm.yyyy") {
    
    if (!is.null(template)) {
      if (class(template) == "Style") {
        style <- template$copy()
        
      } else if (template == "header") {
        style <- generic_header$copy()
        
      } else if (template == "subheader") {
        style <- generic_subheader$copy()
        
      } else if (template == "varname") {
        style <- generic_varname$copy()
        
      } else if (template == "data_year") {
        style <- generic_year_data$copy()
        
      } else if (template == "data_decimal") {
        style <- generic_decimal_data$copy()
      
      } else {
        style <- generic_number_data$copy()
        
      }
    } else {
      style <- generic_number_data$copy()
      
    }
    
    
    if(date) {
      options(openxlsx.dateFormat = dateformat)
      style$numFmt$numFmtId <- "14"
    }
    
    # Customize style
    if (!is.null(halign)) {
      style$halign <- halign
    }
    
    if (!is.null(valign)) {
      style$valign <- valign
    }
    
    if (!is.null(decoration)) {
      style$fontDecoration <- valign
    }
    if (!is.null(decoration)) {
      decoration <- tolower(decoration)
      if (!all(
        decoration %in% c(
          "bold",
          "strikeout",
          "italic",
          "underline",
          "underline2",
          "accounting",
          "accounting2"
          )
      )) {
        stop("Invalid textDecoration!")
      }
      style$fontDecoration <- toupper(decoration)
      
    }
    if (!is.null(fontName)) {
      style$fontName <- list(val = fontName)
    }
    
    if (!is.null(fontSize)) {
      style$fontSize <- list(val = fontSize)
    }
    
    return(style)
  }