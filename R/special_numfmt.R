#' special_numfmt
#' 
#' Function applys specific numfmt on selected cells. Works only in cobination with `create_index_df`
#'
#' @param wb a workbook object
#' @param sheet name of the sheet
#' @param index_df result of `create_index_df`
#' @param dataStart index of the row where data starts in the excel file
#' @param colShift number of leading columns before the numeric data starts (difference in columns in excel and columns in data)
#' @param below1000 numFmt for integer numbers below 1000 (default =  #####)
#' @param from1000 numFmt for integer numbers from 1000 onwards (default = ###'##0)
#' @param decimal numFmt for decimal numbers (default = ###0.0)
#' @param prefix character string that will be put before the number. Every character value has to start with "\\". For Example if you wish to have to stars in front of the numer (e.g. `**100`), you need to write the suffix as follows. `\\*\\*`
#' @param suffix character string that will be put after the number. Same style as prefix
#'
special_numFmt <- function(wb,
                           sheet,
                           index_df,
                           dataStart,
                           colShift = 0,
                           below1000 = "#####0",
                           from1000 = "###'##0",
                           decimal = "###0.0",
                           prefix = NULL,
                           suffix = NULL) {
  if (is.null(suffix) &
      is.null(prefix) &
      below1000 == "#####0" & 
      from1000 == "###'##0" & 
      decimal == "###0.0") {
    warning("No special format given. Therefore no style will be applied")
  } else {
    below1000 = paste0(prefix, below1000, suffix)
    from1000 = paste0(prefix, from1000, suffix)
    decimal = paste0(prefix, decimal, suffix)

    for (i in 1:nrow(index_df)) {
      if (index_df$integer[i] == TRUE) {
        # wenn integer dann anderes Zahlenformat
        if (index_df$val_num[i] >= 1000) {
          # Wenn kleiner 1000 dann keine Tausendertrenner
          add_style_feature(
            wb = wb,
            sheet = sheet,
            rows = index_df$row[i] + dataStart - 1,
            cols = index_df$col[i] + colShift,
            feature_names = list("numFmt"),
            feature_values = list(list(
              numFmtId = 165, formatCode = from1000
            ))
          )
        } else {
          add_style_feature(
            wb = wb,
            sheet = sheet,
            rows = index_df$row[i] + dataStart - 1,
            cols = index_df$col[i] + colShift,
            feature_names = list("numFmt"),
            feature_values = list(list(
              numFmtId = 165, formatCode = below1000
            ))
          )
        }
        
      } else {
        # Style für dezimal zahlen
        add_style_feature(
          wb = wb,
          sheet = sheet,
          rows = index_df$row[i] + dataStart - 1,
          cols = index_df$col[i] + colShift,
          feature_names = list("numFmt"),
          feature_values = list(list(
            numFmtId = 165, formatCode = decimal
          ))
        )
      }
    }
  }
}


