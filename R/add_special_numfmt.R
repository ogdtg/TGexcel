#' add_special_numfmt
#' 
#' Function writes data to workbook and adds special numFmt for matching cells
#'
#' @param wb a workbook object
#' @param sheet name of the sheet
#' @param data a data.frame. should only conatin the numbers you want to write to the excel but no leading column with string information
#' @param pattern a regular expression that should be looked for in the data
#' @param startCol where the data insert should start
#' @param dataStart row in which the numeric data in the file starts (e.g. first row after header)
#' @param position_prefix if `TRUE` the pattern is inserted as prefix, otherwiese as suffix (default is prefix)

#'
#' @export
#'
add_special_numfmt <- function(wb,sheet,data, pattern,startCol,dataStart,position_prefix=TRUE) {
  num_styles_pure <- c("#####0","###'##0","###0.0")
  
  if(position_prefix){
    num_styles_pattern <- paste0(pattern,num_styles_pure)
  } else {
    num_styles_pattern <- paste0(num_styles_pure,pattern)
  }
  
  
  num_styles_vec <- c(num_styles_pure,num_styles_pattern)
  
  
  styles <- lapply(num_styles_vec, function(x){
    createStyle(numFmt = x)
  })
  
  names(styles) <- c("below1000","above1000","decimal","below1000_pat","above1000_pat","decimal_pat")
  
  index_df <- data.frame(rows = rep(1:nrow(data),times = ncol(data)),
                         cols = rep(1:ncol(data),each = nrow(data)))
  
  index_df$val <- sapply(seq_along(index_df$rows), function(i){
    data[index_df$rows[i],index_df$cols[i]]
  })
  
  index_df <- index_df %>% 
    mutate(real_col = cols + startCol-1,
           real_row = rows + dataStart-1,
           pure_val = as.numeric(str_remove(val,pattern))) %>% 
    mutate(cell_style = case_when(
      str_detect(val,pattern) & floor(pure_val) == pure_val & abs(pure_val)<1000 ~ "below1000_pat",
      str_detect(val,pattern) & floor(pure_val) != pure_val  ~ "decimal_pat",
      str_detect(val,pattern) & floor(pure_val) == pure_val & abs(pure_val)>=1000 ~ "above1000_pat",
      !str_detect(val,pattern) & floor(pure_val) == pure_val & abs(pure_val)<1000 ~ "below1000",
      !str_detect(val,pattern) & floor(pure_val) != pure_val  ~ "decimal",
      !str_detect(val,pattern) & floor(pure_val) == pure_val & abs(pure_val)>=1000 ~ "above1000",
      TRUE~NA
    ))
  
  data <- data %>% dplyr::mutate_all(.funs = function(x) {
    str_remove_all(x, pattern)
  }) %>% dplyr::mutate_all(.funs = as.numeric)
  openxlsx::writeData(wb = wb, sheet = sheet, x = data, startRow = dataStart, 
                      startCol = startCol, colNames = FALSE)
  
  
  temp <- lapply(seq_along(index_df$rows), function(i){
    addStyle(wb,sheet,
                       rows = index_df$real_row[i],
                       cols = index_df$real_col[i],
                       style = styles[[index_df$cell_style[i]]],stack = T)
  })
  
}
