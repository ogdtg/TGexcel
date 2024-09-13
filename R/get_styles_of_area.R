#' get_styles_of_area
#' 
#' Returns a list of styles present in a user defined area
#'
#' @param wb a workbook
#' @param sheet sheet name
#' @param rows rows that should be looked at
#' @param cols columns that should be looked at
#'
#' @return list of styles including cols and rows of the style
#' @export
#'
get_styles_of_area <- function(wb, sheet, rows, cols) {
  new_styles <- lapply(seq_along(wb$styleObjects), function(i){
    
    if (wb$styleObjects[[i]]$sheet == sheet){
      existing_cols <- sum(cols %in% wb$styleObjects[[i]]$cols)
      existing_rows <- sum(rows %in% wb$styleObjects[[i]]$rows)  
      
      if (existing_rows>0 & existing_cols >0){
        df <- data.frame(wbrows = wb$styleObjects[[i]]$rows, wbcols = wb$styleObjects[[i]]$cols) %>% 
          dplyr::filter(wbrows %in% rows & wbcols %in% cols)
        if (nrow(df)>0) {
          return(list(cols = df$wbcols, rows = df$wbrows, style = wb$styleObjects[[i]]$style))
          
        }
      }
    }
  })
  new_styles <- new_styles[lengths(new_styles) != 0]
  return(new_styles)
  
}