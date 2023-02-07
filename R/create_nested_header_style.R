#' create_nested_header_style
#'
#' @param wb a workbook object
#' @param sheet the name of the Sheet in the workbook object
#' @param vars_level1 vector of variable names (first level)
#' @param vars_level2 vector of variable names (second level)
#' @param nesting if the variables should be evenly distributed, then provide the string "even" (= default). If the variables should NOT be evenly distributed then provide a numeric vector with the same length as vars_level1. The given integers indicate how many variables from level 2 are nested below one variable from level 1
#' @param startRow row in which the variable names should be placed (default = 3)
#' @param vars_level1_style style for main variables. Default is varname_normal_10_left
#' @param vars_level2_style style for sub variables. Default is varname_normal_10_left
#' @param ... see create_style_generic
#' 
#' @export
#'
create_nested_header_style <-  function(wb, sheet, vars_level1, vars_level2, nesting = "even", startRow = 3, vars_level1_style = varname_normal_10_left, vars_level2_style= varname_normal_10_left,...){

  lv1 <- length(vars_level1)
  lv2 <- length(vars_level2)

  # Can the levels be evenly distributed
  modulu_slots <- lv2 %% lv1

  # Length of nesting vec
  length_nest <- length(nesting)

  if (length_nest < 2) {
    if (modulu_slots != 0 & nesting == "even") {
      stop("Number of vars_level1 must be dividable by the length of vars_level2. Even distribution is not possible.")
    }

    if (nesting == "even") {
      message("No nesting vector given. Variables will be evenly distributed")
    }
  }


  if (!is.character(nesting)) {
    nesting_sum <- sum(nesting)
    nesting_length <- length(nesting)

    if (nesting_sum != lv2) {
      stop(paste0("The given nesting units do not match the length of vars_level2 (",nesting_sum,"!=",lv2,")"))
    }

    if (nesting_length != lv1) {
      stop(paste0("The number of given nesting units does not fit the length of vars_level1 (",nesting_length,"!=",lv1,")"))
    }


  }
  
  vars_level2_style_mod <- vars_level2_style$copy()
  vars_level2_style_mod$borderTop <- NULL
  vars_level2_style_mod$borderBottom <- NULL

  create_varnames_style(wb = wb, sheet=sheet, startRow = startRow+1,var_names = vars_level2,style = vars_level2_style_mod,...)
  create_varnames_style(wb = wb, sheet=sheet, startRow = startRow,ncol = lv2,style_only = TRUE,style = vars_level1_style,...)
  
  # # sTYLES WITHOUT BORDER
  varname_no_border <- vars_level1_style$copy()
  varname_no_border$borderTop <- NULL
  varname_no_border$borderBottom <- NULL
  # varname_no_border_bottom <- vars_level1_style$copy()
  # varname_no_border_bottom$borderBottom <- NULL
  

  if (nesting=="even") {
    merge_size <- lv2/lv1
    start = 1
    for (i in 1: lv1) {
      
      
      
      openxlsx::mergeCells(wb,sheet,cols = start:(start+merge_size-1), rows = startRow )
      
      if (grepl("\\[.*\\]", vars_level1[i]) | grepl("\\~.*\\~", vars_level1[i])) {
        addSuperSubScriptToCell(wb=wb,
                                sheet = sheet,
                                col = start,
                                row = startRow,
                                texto = vars_level1[i],
                                style = vars_level1_style)
      } else {
        openxlsx::writeData(x = vars_level1[i],
                                  wb = wb,
                                  sheet = sheet,
                                  startCol = start,
                                  startRow = startRow)
      }
      # Correct border settings
      if (merge_size==1) {
        openxlsx::addStyle(wb,
                           sheet,
                           rows = startRow,
                           cols = start,
                           style = varname_no_border)
      }

      

      start = start + merge_size


    }
  } else {
    nesting = unlist(nesting)
    start = 1
    for (i in 1:length(nesting)) {
      
      
      
      merge_size <- nesting[i]
      openxlsx::mergeCells(wb,sheet,cols = start:(start+merge_size-1), rows = startRow)
      
      if (grepl("\\[.*\\]", vars_level1[i]) | grepl("\\~.*\\~", vars_level1[i])) {
        addSuperSubScriptToCell(wb=wb,
                                sheet = sheet,
                                col = start,
                                row = startRow,
                                texto = vars_level1[i],
                                style = vars_level1_style)
      } else {
        openxlsx::writeData(x = vars_level1[i],
                            wb = wb,
                            sheet = sheet,
                            startCol = start,
                            startRow = startRow)
      }
      
      # Correct border settings
      if (nesting[i]==1) {
        openxlsx::addStyle(wb,
                           sheet,
                           rows = startRow,
                           cols = start,
                           style = varname_no_border)
      }
      
      

      start <- merge_size+start

    }
    


  }




}
