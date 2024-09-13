#' addSuperSubScriptToCell
#'
#' See https://github.com/awalker89/openxlsx/issues/407#issuecomment-841870240
#'
#' @param wb a workbook object
#' @param sheet the name of the Sheet in the workbook object
#' @param row row index
#' @param col col index
#' @param texto data to write
#' @param style cell style
#' @param colour font colour in hex represenation (default is black -> default = '000000')
#'
addSuperSubScriptToCell <- function(wb,
                                    sheet,
                                    row,
                                    col,
                                    texto,
                                    style,
                                    colour = '000000') {

  font <- style$fontName
  family <- style$fontFamily
  size <- style$fontSize
  fDec <- style$fontDecoration

  placeholderText <- 'This is placeholder text that should not appear anywhere in your document.'

  openxlsx::writeData(wb = wb,
                      sheet = sheet,
                      x = placeholderText,
                      startRow = row,
                      startCol = col)

  #finds the string that you want to update
  stringToUpdate <- which(sapply(wb$sharedStrings,
                                 function(x){
                                   grep(pattern = placeholderText,
                                        x)
                                 }
  )
  == 1)

  #splits the text into normal text, superscript and subcript

  normal_text <- stringr::str_split(texto, "\\[.*\\]|~.*~") %>% purrr::pluck(1) %>% purrr::discard(~ . == "")

  sub_sup_text <- stringr::str_extract_all(texto, "\\[.*\\]|~.*~") %>% purrr::pluck(1)

  if (length(normal_text) > length(sub_sup_text)) {
    sub_sup_text <- c(sub_sup_text, "")
  } else if (length(sub_sup_text) > length(normal_text)) {
    normal_text <- c(normal_text, "")
  }

  #formatting instructions

  sz    <- paste('<sz val =\"',size,'\"/>',
                 sep = '')
  col   <- paste('<color rgb =\"',colour,'\"/>',
                 sep = '')
  rFont <- paste('<rFont val =\"',font,'\"/>',
                 sep = '')
  fam   <- paste('<family val =\"',family,'\"/>',
                 sep = '')


  # this is the separated text which will be used next
  texto_separado <- purrr::map2(normal_text, sub_sup_text, ~ c(.x, .y)) %>%
    purrr::reduce(c) %>%
    purrr::discard(~ . == "")



  #if its sub or sup adds the corresponding xml code
  sub_sup_no <- function(texto) {

    if(stringr::str_detect(texto, "\\[.*\\]")){
      return('<vertAlign val=\"superscript\"/>')
    } else if (stringr::str_detect(texto, "~.*~")) {
      return('<vertAlign val=\"subscript\"/>')
    } else {
      return('')
    }
  }

  #get text from normal text, sub and sup
  get_text_sub_sup <- function(texto) {
    stringr::str_remove_all(texto, "\\[|\\]|~")
  }

  if (length(fDec)>0) {
    #formating
    if(fDec=="BOLD"){
      bld <- '<b/>'
    } else{bld <- ''}

    if(fDec=="ITALIC"){
      itl <- '<i/>'
    } else{itl <- ''}
  } else {
    bld <-  ""
    itl <- ""
  }



  #get all properties from one element of texto_separado

  get_all_properties <- function(texto) {

    paste0('<r><rPr>',
           sub_sup_no(texto),
           sz,
           col,
           rFont,
           fam,
           bld,
           itl,
           '</rPr><t xml:space="preserve">',
           get_text_sub_sup(texto),
           '</t></r>')
  }


  # use above function in texto_separado
  newString <- purrr::map(texto_separado, ~ get_all_properties(.)) %>%
    purrr::reduce(paste, sep = "") %>%
    {c("<si>", ., "</si>")} %>%
    purrr::reduce(paste, sep = "")

  # replace initial text
  wb$sharedStrings[stringToUpdate] <- newString
}
