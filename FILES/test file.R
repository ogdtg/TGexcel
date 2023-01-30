library(openxlsx)


vars1 <- c("","","Heirat eines Schweizers mit","Heirat eines Ausländers mit")
vars2 <- c("Jahr","Total","einer Schweizerin","einer Ausländerin","einer Schweizerin","einer Ausländerin")

stringi::stri_rand_strings(12, 4, pattern = "[A-Za-z0-9]")

data = tibble(Jahr = 1999:2010,
              bfs_nr = 2000:2011,
              Name = stringi::stri_rand_strings(12, 4, pattern = "[A-Za-z0-9]"),
              alter = 1:12,
              anzahl = 1:12,
              proz = seq(0.01,0.12,0.01),
              test = 1:12,
              tes2t = seq(0.01,0.12,0.01))

test_wb <- createWorkbook()
addWorksheet(test_wb,"test")
test_wb <- create_header_style(wb = test_wb,ncol = length(vars2),text = "Heiraten nach Staatsangehörigkeit der Ehepartner", sheet = "test")
test_wb <- create_subheader_style(wb = test_wb,ncol = length(vars2),text = "Kanton Thurgau, 2000-2021", sheet = "test")

test_wb <-
  create_nested_header_style(
    wb = test_wb,
    sheet = "test",
    var_level1 = vars1,
    var_level2 = vars2,
    nesting = c(1, 1, 2, 2) ,
    halign = "center",
    startRow = 3
  )

# create_data_style(wb = test_wb, sheet ="test", data =data, startRow = 5,year = 1,datenquelle = "Bundesamt für Statistik, BEVNAT" )

setColWidths(test_wb,sheet = "test",cols = 1:length(vars2),widths = 15)




saveWorkbook(test_wb,"functions.xlsx", overwrite = T)



