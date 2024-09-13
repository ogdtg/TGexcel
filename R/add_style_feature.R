#' add_style_feature
#' 
#' A Function to change specific parameters of a cellstyle without overwriting it.
#'
#' @param wb a workbook object
#' @param sheet sheet name
#' @param rows rows the style should be added to
#' @param cols columns the style should be added to
#' @param feature_names a list including the names of the features you wish to change (see [openxlsx::createStyle()])
#' @param feature_values a list of values for the  corresponding features from `feature_names`
#'
#' @export
#'
add_style_feature <- function(wb,sheet, rows,cols,feature_names, feature_values) {
  
  illegal_names <- feature_names[!feature_names %in% featureNames]
  feature_names_mod <- feature_names[feature_names %in% featureNames]
  feature_values <- feature_values[which(!feature_names %in% illegal_names)]
  
  if (length(illegal_names>0)) {
    message(paste0(illegal_names," is not a valid feature and will be ignored\n"))
  }
  if (length(feature_names_mod)==0) {
    stop("No valid feature name given.")
  }
  area_styles <- get_styles_of_area(wb, sheet, rows, cols)
  
  
  invisible(lapply(area_styles, function(x){
    temp <- x$style$copy()
    for (i in 1:length(feature_names_mod)) {
      
      if (feature_names_mod[i] %in% c("fillBg","fillFg")) {
        temp$fill[[feature_names_mod[[i]]]] <- feature_values[[i]]
        next
      }
      
      temp[[feature_names_mod[[i]]]] <- feature_values[[i]]
    }
    openxlsx::addStyle(wb,sheet = sheet, style = temp, rows = unique(x$rows), cols = unique(x$cols), gridExpand = T)
  }))
  
}