#' save_tg_workbook
#'
#' @param wb a workbook object
#' @param filename path where the file should be stored. Must have the extension .xlsx
#' @param tg_header if TRUE the header of the Internettabellen will be included (default = FALSE)
#' @param overwrite if TRUE, an existing file will be overwritten
#'
#' @export
#'
save_tg_workbook <- function(wb,filename,tg_header = TRUE, overwrite = FALSE) {
  
  filename <- suppressWarnings(normalizePath(filename))
  
  if (!grepl("\\.xlsx$",filename)) {
    stop("filename must have the extension .xlsx")
  }
  
  #Header entfernen
  
  a <- lapply(wb$sheet_names,function(x){
    openxlsx::setHeaderFooter(wb,
                    sheet = x,
                    header = c("","",""),
                    footer = c("","","")
                    
    )
  })
  
  # Bild und VML löschen
  wb$media <- NULL
  wb$vml <- NULL
  wb$vml_rels <- NULL
  
  # WB ohne header abspeichern
  openxlsx::saveWorkbook(wb,filename,overwrite = overwrite)
  
  vbs_file <- normalizePath(system.file("extdata", "importAndRunModule.vbs", package = "TGexcel"))
  modul1 <- normalizePath(system.file("extdata", "Modul1.bas", package = "TGexcel"))
  imagePath <- normalizePath(system.file("extdata", "image1.png", package = "TGexcel"))

  # 
  # #Neuer Header einfügen
  if (tg_header) {
    # filename <- gsub("\\.xlsx","",filename)
    system_command <- paste("WScript",
                            vbs_file,
                            paste0('"""',filename,'"""'),
                            modul1,
                            # '"AddFooterHeaderImage"',
                            imagePath,
                            sep = " ")
    test <- system(command = system_command,
           wait = TRUE)

    if (test==0) {
      message("File saved")
    } else {
      stop("File could not be saved")
    }
  } else {
    message("File saved without header")
  }
  
}
