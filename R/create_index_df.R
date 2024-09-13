#' create_index_df
#' 
#' Function returns a dataset with the indeces of those cells that are matching the selected pattern
#'
#' @param data a data.frame
#' @param pattern a regular expression that should be looked for in the data
#'
#' @return a data.frame with row and col index of matching cells
#' @export
#'
create_index_df <- function(data,pattern = "\\*\\*") {
  data = dplyr::as_tibble(data)
  index_list <- list()
  for (i in 1:ncol(data)){
    for (k in 1:nrow(data[,i])){
      if(stringr::str_detect(data[k,i],pattern)) {
        index_list <-
          append(list(
            list(
              colName = names(data)[i],
              col = i,
              row = k,
              value = dplyr::pull(data[k, i])
            )
          ), index_list)
      }
    }
  }
  index_df <- index_list %>% dplyr::bind_rows()
  index_df <- index_df %>% 
    dplyr::mutate(val_num = stringr::str_remove(value,"\\*\\*"),
           val_num = as.numeric(val_num)) %>%
    dplyr::mutate(integer = ifelse(val_num-floor(val_num)==0,TRUE,FALSE))
  
  return(index_df)
}
