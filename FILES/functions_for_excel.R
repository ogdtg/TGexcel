create_header_style <- function(wb, sheet, ncol,text) {
  style <- readRDS("header_bold_12_left.rds")
  
  openxlsx::addStyle(wb = wb,
           sheet = sheet,
           cols = 1:ncol,
           rows = 1,
           style = style)
  
  openxlsx::writeData(x = text,
                      wb = wb,
                      sheet = sheet,
                      startCol = 1,
                      startRow = 1)
  
  return(wb)
  
}

create_subheader_style <- function(wb, sheet, ncol,text) {
  style <- readRDS("subheader_normal_10_left.rds")
  
  openxlsx::addStyle(wb = wb,
                     sheet = sheet,
                     cols = 1:ncol,
                     rows = 2,
                     style = style)
  
  openxlsx::writeData(x = text,
                      wb = wb,
                      sheet = sheet,
                      startCol = 1,
                      startRow = 2)
  
  return(wb)
  
}

create_varnames_style <- function(wb, sheet,var_names, halign = "left", startRow) {
  
  style <- readRDS("varname_normal_10_left.rds")  

  if (halign == "right") {
    style[["halign"]] <- "right"
  }
  if (halign == "center") {
    style[["halign"]] <- "center"
  }
  
  
  openxlsx::addStyle(wb = wb,
                     sheet = sheet,
                     cols = 1:length(var_names),
                     rows = startRow,
                     style = style)
  
  lapply(seq_along(var_names), function(i){
    openxlsx::writeData(x = var_names[i],
                        wb = wb,
                        sheet = sheet,
                        startCol = i,
                        startRow = startRow)
  })
  
  
  return(wb)
  
}

create_varnames_style_only <- function(wb, sheet,var_names, halign = "left", startRow) {
  
  style <- readRDS("varname_normal_10_left.rds")  
  
  if (halign == "right") {
    style[["halign"]] <- "right"
  }
  if (halign == "center") {
    style[["halign"]] <- "center"
  }
  
  openxlsx::addStyle(wb = wb,
                     sheet = sheet,
                     cols = 1:length(var_names),
                     rows = startRow,
                     style = style)
  
  return(wb)
  
}

create_singlevar_style <- function(wb, sheet,var_name, halign = "left", row, col) {

  
  style <- readRDS("varname_normal_10_left.rds")  
  
  if (halign == "right") {
    style[["halign"]] <- "right"
  }
  if (halign == "center") {
    style[["halign"]] <- "center"
  }
  
  openxlsx::addStyle(wb = wb,
                     sheet = sheet,
                     cols = col,
                     rows = row,
                     style = style)
  
    openxlsx::writeData(x = var_name,
                        wb = wb,
                        sheet = sheet,
                        startCol = col,
                        startRow = row)
  
  
  return(wb)
  
}

create_nested_header_style <-  function(wb, sheet, var_level1, var_level2, nesting = "even", startRow,...){
  
  lv1 <- length(var_level1)
  lv2 <- length(var_level2)
  
  print(lv1)
  print(lv2)
  
  modulu_slots <- lv2 %% lv1
  
  if (modulu_slots != 0 & nesting == "even") {
    stop("Number of var_level1 must be dividable by the length of var_level2. Even distribution is not possible.")
  }
  
  if (nesting == "even") {
    "No nesting vector given. Variables will be evenly distributed"
  }
  
  if (!is.character(nesting)) {
    nesting_sum <- sum(nesting)
    nesting_length <- length(nesting)
    
    if (nesting_sum != lv2) {
      stop(paste0("The given nesting units do not match the length of var_level2 (",nesting_sum,"!=",lv2,")"))
    }
    
    if (nesting_length != lv1) {
      stop(paste0("The number of given nesting units does not fit the length of var_level1 (",nesting_length,"!=",lv1,")"))
    }
    
    
  }
  
  wb <- create_varnames_style(wb = wb, sheet=sheet, startRow = startRow+1,var_names = var_level2,...)
  
  
  if (nesting=="even") {
    merge_size <- lv2/lv1
    start = 1
    for (i in 1: lv1) {
      mergeCells(wb,sheet,cols = start:(start+merge_size-1), rows = startRow )
      # wb <- create_singlevar_style(wb,sheet,var_name = var_level1[i], col = start, row=startRow)
      
      openxlsx::writeData(x = var_level1[i],
                          wb = wb,
                          sheet = sheet,
                          startCol = start,
                          startRow = startRow)
      
      start = start + merge_size
      
      
    }
  } else {
    nesting = unlist(nesting)
    start = 1
    for (i in 1:length(nesting)) {
      merge_size <- nesting[i]
      mergeCells(wb,sheet,cols = start:(start+merge_size-1), rows = startRow)
      # wb <- create_singlevar_style(wb,sheet,var_name = var_level1[i], col = start, row=startRow)
      openxlsx::writeData(x = var_level1[i],
                          wb = wb,
                          sheet = sheet,
                          startCol = start,
                          startRow = startRow)
      
      start <- merge_size+start
      
    }

  }
  
  wb <- create_varnames_style_only(wb = wb, sheet=sheet, startRow = startRow,var_names = var_level2,...)
  
  
  return(wb)
  
}


create_data_style <- function(wb,sheet,startRow, date = NULL,year = NULL, startCol=1, data, datenquelle = NULL) {
  
  style_number <- readRDS("number_normal_10_right.rds")
  style_year <- readRDS("jahr_normal_10_right.rds")
  
  openxlsx::writeData(x = data,
                      wb = wb,
                      sheet = sheet,
                      startCol = startCol,
                      startRow = startRow,
                      colNames = FALSE)
  
  a <- lapply(1:ncol(data), function(x) {
    
    if (year == x) {
      openxlsx::addStyle(wb = wb,
                         sheet = sheet,
                         cols = startCol+x-1,
                         rows = startRow:(startRow+nrow(data)-1),
                         style = style_year)
    } else {
      openxlsx::addStyle(wb = wb,
                         sheet = sheet,
                         cols = startCol+x-1,
                         rows = startRow:(startRow+nrow(data)-1),
                         style = style_number)
    }
    
    
    
  })
  
  if (!is.null(datenquelle)) {
    datenquelle_italic_9_left <- readRDS("datenquelle_italic_9_left.rds")
    
    openxlsx::writeData(x = paste0("Datenquelle: ",datenquelle),
                        wb = wb,
                        sheet = sheet,
                        startCol = startCol,
                        startRow = startRow+nrow(data)+2,
                        colNames = FALSE)
    
    openxlsx::addStyle(wb = wb,
                       sheet = sheet,
                       cols = startCol,
                       rows = startRow+nrow(data)+2,
                       style = datenquelle_italic_9_left)
    
  }
}



