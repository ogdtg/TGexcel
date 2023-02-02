#' create_style_df
#' 
#' Helper Function to delete NAs
#'
#' @param wb a workbook
#'
#' @return data.frame
#'
create_style_df <- function(wb){
  
  sheetnames <- wb$sheet_names
  
  data_list <- lapply(seq_along(sheetnames), function(i) {
    c <- wb$worksheets[i][[1]][["sheet_data"]][[".->cols"]]
    r <- wb$worksheets[i][[1]][["sheet_data"]][[".->rows"]]
    t <- wb$worksheets[i][[1]][["sheet_data"]][[".->t"]]
    v <- wb$worksheets[i][[1]][["sheet_data"]][[".->v"]]
    sheet <- sheetnames[i]
    
    return(data.frame(sheet,c,r,t,v))
    
  })
  data_df <- data_list %>%dplyr::bind_rows()
  
  style_list <- lapply( wb$styleObjects, function(style){
    fill_fg <- style$style$fill$fillFg
    fill_bg <- style$style$fill$fillBg
    
    if (is.null(fill_bg)){
      fill_bg <- NA
    }
    
    if (is.null(fill_fg)){
      fill_fg <- NA
    }
    
    cols <- style$cols
    rows <- style$rows
    sheet <- style$sheet
    return(data.frame(sheet,c=cols,r=rows,fill_fg,fill_bg, row.names = NULL))
  })
  
  style_df <- style_list %>%  dplyr::bind_rows()
  
  result <- data_df %>% 
    dplyr::left_join(style_df) %>% 
    dplyr::filter(v==1 & t==1)
  
  return(result)
}
