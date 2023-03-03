# TGexcel

<!-- badges: start -->
[![Lifecycle: experimental](https://img.shields.io/badge/lifecycle-experimental-orange.svg)](https://lifecycle.r-lib.org/articles/stages.html#experimental)
<!-- badges: end -->

R Package um Internettabellen aus R anzulegen und zu formatieren. TGexcel ist eine Erweiterung des openxlsx Package und basiert weitestgehend auf den Funktionen aus openxlsx.

Das Package soll es dem User möglich machen die Erstellung der Internettabellen direkt aus einem R-Prozess heraus durchzuführen. So können Prozesse weiter standardisiert und automatisiert werden.

Das Package macht es möglich. Titel, Subtitel und Variablennamen (auch verschachtelt) im entsprechenden Format anzulegen. Ausserdem können auch die Daten selbst im gewünschten Format eingetragen werden. Auch die Kopfzeile mit Logo kann per Funktion eingefügt werden. Es können sowohl neue Excel Files angelegt erden als auch bereits existierende Files bearbeitet werden.

Da das Package auf `openxlsx` basiert, werden für alle Funktionen workbook-Objekte verwendet. Diese können entweder mit `createWorkbook()` neu erstellt, oder mit `loadWorkbook("path/to/file.xlsx")` bearbeitet werden.

![Aufbau der Excel](https://github.com/ogdtg/TGexcel/blob/main/img/01_aufbau.PNG)


## Installation

Die Development Version des Packages kann wie folgt installiert und genutzt werden:

``` r
devtools::install_github("ogdtg/TGexcel")
library(TGexcel)
```


## Titel (header)

Mit der `create_header_style()` Funktion wird der Titel in das Excel File eingefügt. Neben dem workbook Object, dem Name des Sheets sowie dem Titeltext muss für `ncol` angegeben werden, wie viele Spalten der Titel bzw. die Tabelle umfassen soll. Hier sollte die Anzahl der Spalten, die der Datensatz umfasst eingetragen werden. Per Default ist als Startzeile (`startRow`) die erste Zeile des Sheets angegeben. Dieser Wert kann aber auch angepasst werden. Gleiches gilt für die Argumente `style` und `rowHeight`.

```r
wb <- createWorkbook()
addWorksheet(wn,"Test")

create_header_style(wb = wb, sheet = "Test", ncol = 6, text = "Das ist der Titel")

# Workbook als xlsx speichern
saveWorkbook(wb, "test.xlsx", overwrite = TRUE)

```
Das Ergebnis nach dem Speichern sieht wie folgt aus:

![Excel Header](https://github.com/ogdtg/TGexcel/blob/main/img/02_header.PNG)


## Subtitel (subheader)

Mit `create_subheader_style()` kann ein Untertiel eingefügt werden. Der Funktionsaufbau ist der selbe wie bei `create_header_style()`.

```r
## Subheader
wb <- loadWorkbook("test.xlsx")
create_subheader_style(wb = wb, sheet = "Test", ncol = 6, text = "Das ist der Subtitel")

# Workbook als xlsx speichern
saveWorkbook(wb, "test.xlsx", overwrite = TRUE)
```
![Excel Subheader](https://github.com/ogdtg/TGexcel/blob/main/img/03_subheader.PNG)


## Variablennamen (varnames)

Das Package bietet die Möglichkeit Variablennamen mit oder ohne Verschachtelung einzutragen.

### Unverschachtelt

Bei unverschachtelten Variablennamen muss für `create_varnames_style()` lediglich ein Vektor mit den Spaltennamen angegeben werden. Per default ist `startRow=3`.

```r
# Unverschachtelt
wb <- loadWorkbook("test.xlsx")
create_varnames_style(wb = wb, sheet = "Test", var_names = c("Var1","Var2","Var3","Var4","Var5", "Var6"))


# Workbook als xlsx speichern
saveWorkbook(wb, "test_01.xlsx", overwrite = TRUE)
```
![Excel Varnames](https://github.com/ogdtg/TGexcel/blob/main/img/04_varnames.PNG)


### Verschachtelt

Bei verschachtelten Variablennamen kann die `create_nested_header_style()` Funktion verwendet werden.
Hierbei müssen zwei Vektoren mit den übergeordneten (`vars_level1`) bzw. untergeordneten (`vars_level2`) Variablennamen als Argument angegeben werden. Des Weiteren muss der `nesting` Parameter angegeben werden. Hier kann wahlweise ein numerischer Vektor mit der gleichen Länge wie `vars_level1` angegeben werden oder der String `"even"`. Wird `"even"` angegeben versucht die Funktion die untergeordneten Variablen gleichmässig auf die übergeordneten Variablen zu verteilen. Dies funktioniert nur, wenn die Anzahl der untergeordneten Variablen durch die Anzahl der übergeordneten Variablen teilbar ist (Modulu = 0).

Besipiel:

```r
wb <- loadWorkbook("test.xlsx")
create_nested_header_style(wb = wb,
                           sheet = "Test",
                           nesting = "even",
                           vars_level1 = c("SVar1","SVar2","SVar3"),
                           vars_level2 = c("Var1","Var2","Var3","Var4","Var5", "Var6"))


# Workbook als xlsx speichern
saveWorkbook(wb, "test_02.xlsx", overwrite = TRUE)

```

![Excel Nested Even](https://github.com/ogdtg/TGexcel/blob/main/img/05__nested_even.PNG)

Es ist auch möglich eine benutzerdefinierte Verschachtelung einzufügen. Dazu muss beim `nesting` Paramteter ein numerischer Vektor mit der selben Länge wie `vars_level1` angegeben werden. Jeder Wert in diesem Vektor gibt an wie viele Variablen der übergeorndeten Variable an dieser Stelle untergeordnet werden sollen.

Ein Beispiel: es ist gegeben
`vars_level1 = c("SVar1","SVar2","SVar3")`
`vars_level2 = c("Var1","Var2","Var3","Var4","Var5", "Var6")` und 
`nesting = c(1,3,2)`

Das bedeutet, dass 
- `SVar1` genau eine untergeordnte Variable besitzt (`Var1`),
- `SVar2` genau 3 untergeordnte Variablen besitzt (`Var2`,`Var3` und `Var4`) und
- `SVar3` genau 2 untergeordnte Variablen besitzt (`Var5` und`Var6`)

```r
wb <- loadWorkbook("test.xlsx")
create_nested_header_style(wb = wb,
                           sheet = "Test",
                           nesting = c(1,3,2),
                           vars_level1 = c("SVar1","SVar2","SVar3"),
                           vars_level2 = c("","Var2","Var3","Var4","Var5", "Var6"))



saveWorkbook(wb, "test_03.xlsx", overwrite = TRUE)

```

![Excel Nested Even](https://github.com/ogdtg/TGexcel/blob/main/img/06__nested_uneven.PNG)

Durch Kombination von mehreren `create_nested_header_style` Funktionen können auch mehrfach verschachtelte Variablennamen erstellt werden.


```r

wb <- createWorkbook()
sheet <- "TEST"
addWorksheet(wb,sheet)

#header erstellen
create_header_style(wb, sheet, ncol = ncol(data), text = "Wohnbevölkerung der Gemeinden nach Nationalität und Geschlecht" )

#subheader erstellen
create_subheader_style(wb,
                       sheet,
                       ncol =  ncol(data),
                       text = "Kanton Thurgau, 31.12.2026, ständige Wohnbevölkerung[1] in Personen und Anteile in %",
                       startRow = 2)


# 1. nested header erstellen
create_nested_header_style(
  wb,
  sheet,
  vars_level1 = c(
    "BFS-Nr. [2]",
    "Gemeinde",
    "Total",
    "Nach Nationalität",
    "Nach Geschlecht",
    "Date"
  ),
  vars_level2 = c("", "", "", "Schweiz", "Ausland", "", "Mann", "Frau", "", ""),
  startRow = 3,
  nesting = c(1, 1, 1, 3, 3, 1)
)

# 2. nested header erstellen
create_nested_header_style(
  wb,
  sheet,
  vars_level1 = c("", "", "", "Schweiz", "Ausland", "Mann", "Frau", ""),
  vars_level2 = c("", "", "", "", "absolut", "in %", "", "absolut", "in%", ""),
  startRow = 4,
  nesting = c(1, 1, 1, 1, 2, 1, 2, 1)
)

saveWorkbook(wb, "test_04.xlsx", overwrite = TRUE)

```

![Excel Nested Uneven](https://github.com/ogdtg/TGexcel/blob/main/img/07_nested_uneven.PNG)


## Daten (data)

Die tatsächlichen Daten müssen vor dem Einfügen bereits in der richtigen Reihenfolge sein, was die Spalten angeht. Die im dataframe gespeicherten Variablennamen werden *NICHT* ins Excel File geschrieben, sondern müssen wie oben beschrieben manuell eingefügt werden.

Mit der `create_data_style` kann ein dataframe mit den gewüschten Spezifikationen ins Excel File geschrieben werden.
Dafür muss dieses Mal explizit eine `startRow` angegeben werden. Ausserdem kann der Parameter `gemeinde_format` auf `TRUE` gesetzt werden, damit die Tabelle im Stil der bisherigen Tabellen mit Gemeindebezug formatiert wird (Bezirke und Gesamtkanton fett). Ist eine Jahresvariable in den Daten enthalten, muss deren Spaltenindex für den Parameter `year` angegeben, um die korrekte Formatierung dieser Variable zu erreichen.

Beispiel:
```r

#datensatz erstellen
data <- gemeinde
data$total <- sample(1000:200000, size = nrow(data))
data$nat_ch <- sample(1000:100000, size = nrow(data))
data$nat_ausl <- sample(1000:100000, size = nrow(data))
data$nat_ausl_perc <- runif(nrow(data))*100
data$sex_male <- sample(1000:100000, size = nrow(data))
data$sex_fem <- sample(1000:100000, size = nrow(data))
data$sex_fem_perc <-runif(nrow(data))*100 
data$date <- seq.Date(Sys.Date(), to = Sys.Date()+nrow(data)-1,by="day")

# Tabelle mit mehrfach verschachteltem Header laden
wb <- loadWorkbook("test_04.xlsx")

# Daten einfügen
create_data_style(wb,sheet,startRow = 6, gemeinde = TRUE, data = data)

saveWorkbook(wb, "test_04.xlsx", overwrite = T)

```

![Excel Data](https://github.com/ogdtg/TGexcel/blob/main/img/08_data.PNG)


## Wrapper Funktion für gesamte Tabelle

Tabellen mit einfachen oder einfach verschachtelten Variablennamen können mit der `create_table_style` Funktion erstellt werden. Die Funktion beinhaltet alle oben beschriebenen Funktionen.

Beispiel:

```r

wb <- createWorkbook()
addWorksheet(wb, "TEST")

data <- data.frame(v1 = LETTERS, 
                   v2 = letters, 
                   v3 = 1:26,
                   v4 = runif(26)*100,
                   v5 = seq.Date(Sys.Date(), to = Sys.Date()+26-1,by="day"),
                   v6 = 1996:2021)

create_table_style (
  wb = wb,
  sheet = "TEST",
  header = "Das ist ein Test[1]",
  subheader = "Test mit Wrapper[2] Funktion",
  varnames = c("", "Var2", "Var3", "Var4", "Var5", "Var6"),
  var_categories = c("SVar1", "SVar2", "SVar3"),
  data = data,
  datenquelle = "BEVNAT Kanton Thurgau",
  footnotes = c("1 Test der Funktion", "2 alle Funktionen zusammengefasst"),
  gemeinde_format = FALSE,
  nesting = c(1, 3, 2),
  year = 6
)

saveWorkbook(wb, "test_05.xlsx", overwrite = T)

```
![Excel Wrapper](https://github.com/ogdtg/TGexcel/blob/main/img/09_wrapper.PNG)


## Kopfzeile einfügen

Um die Kopfzeile mit Logo und Name des Amtes einzufügen, kann die `save_tg_workbook()` verwendet werden.
Diese Funktion führt ein VBA Makro in der Excel aus. Dafür wird die Excel Datei kurz geöffnet und im Anschluss daran wieder geschlossen.

```r

wb <- loadWorkbook("test_05.xlsx")
save_tg_workbook(wb,
                 filename = "test_06.xlsx",
                 tg_header = TRUE,
                 overwrite = TRUE)
```

![Excel Kopfzeile](https://github.com/ogdtg/TGexcel/blob/main/img/10_kopfzeile.PNG)


## Daten einfügen in existierender Tabelle

Manchmal müssen Daten in eine existierende Tabelle eingefügt werden. Mit der `insert_data_row` Funktion können Datenzeilen an einer beliebigen Stelle im Tabellenblatt eingefügt werden. Sämtliche Styles der darüberliegenden Zeile werden auf die neuen Zeilen übertragen.

```r
save_tg_workbook(wb, "test.xlsx", overwrite = T, tg_header =F)



wb <- loadWorkbook("test_11.xlsx")

df <- data.frame(
  x1 = "20XX",
  x2 = "Bezirk Testhausen",
  x3 = 10000,
  x4 = 10000,
  x5 = 10000,
  x6 = 50.0,
  x7 = 10000,
  x8 = 10000,
  x9 = 50
)

insert_data_row(wb,
                sheet = "2022",
                insertRow = 7,
                data = df,
                dataStartRow = 6)

save_tg_workbook(wb, "test_13.xlsx", overwrite = T, tg_header =F)

```
![Excel Insert](https://github.com/ogdtg/TGexcel/blob/main/img/16_insert.PNG)

## Bestehende Files bearbeiten

Mit den Basisfunktionen des `openxlsx` Packages können bestehende Excel Files mit neuen Daten befüllt werden. Dazu muss einfach ein bestehendes File mit `loadWorkbook` geöffnet werden. Anschliessend kann das gewünschte Tabellenblatt mit `cloneWorksheet` kopiert werden. Mit `writeData` können dann neue Daten eingefügt werden. Das alte Tabellenblatt kann mit `removeWorksheet` entfernt werden und das Workbook dann unter einem neuen Namen mit `save_tg_workbook` abgespeichert werden.



```r
# Jahr definieren
jahr <- 2021
# Pfad alte Tabellen
path_excel <- paste0("Y:/SK/SKStat/R/Prozesse/vz_see/Erwerb_Ausbildung/erwerb_ausbildung/output/excel/",jahr-1,"/",jahr+1,"_Arbeitsmarktstatus_ab_2015.xlsx")

#Workbook laden
wb_am_status <- loadWorkbook(path_excel)

# Alle Blätter ausser dem aktuellen löschen
sheets_to_remove <- sheets(wb_am_status)

# Tabellenblatt kopieren
cloneWorksheet(wb_am_status, paste0(jahr,"_Personen"), clonedSheet = paste0(jahr-1,"_Personen"))

# Dummy Datensatz erstellen
rownum <- 6
sample_df <- data.frame(x1 = sample(60000:200000,rownum),
                        x2 = runif(rownum),
                        x3 = sample(60000:200000,rownum),
                        x4 = sample(60000:200000,rownum),
                        x5 = sample(60000:100000,rownum),
                        x6 = runif(rownum),
                        x7 = sample(60000:100000,rownum),
                        x8 = sample(60000:100000,rownum),
                        x9 = sample(60000:100000,rownum),
                        x10 = runif(rownum),
                        x11 = sample(60000:100000,rownum),
                        x12 = sample(60000:100000,rownum))

# Daten einfügen
writeData(wb_am_status, sheet = paste0(jahr,"_Personen"),x = sample_df, startCol = 2, startRow = 7, colNames = FALSE)

# Subheader und Datenquelle bearbeiten
writeData(wb_am_status, sheet = paste0(jahr,"_Personen"),x = paste0("Kanton Thurgau, ",jahr,", in Anzahl Personen"), startCol = 1, startRow = 2, colNames = FALSE)
writeData(wb_am_status, sheet = paste0(jahr,"_Personen"),x = paste0("Datenquelle: Bundesamt für Statistik, Strukturerhebung ",jahr), startCol = 1, startRow = 22, colNames = FALSE)

# Alle nicht gebrauchten Sheets löschen
lapply(sheets_to_remove, function(x){
  removeWorksheet(wb_am_status,sheet = x)
})


```


**Vorher:**
![Original](https://github.com/ogdtg/TGexcel/blob/main/img/17_origedit.PNG)

**Nachher:**
![Original](https://github.com/ogdtg/TGexcel/blob/main/img/18_newedit.PNG)


## Cellstyles erstellen und bearbeiten

### Komplette Styles

Im Package sind bereits mehrere Cellstyles enthalten. Diese können mit der `customize_style` Funktion rudimentär bearbeitet werden.

```r

wb <- loadWorkbook("test_05.xlsx")

# Style bearbeiten
varname_bold_left <- customize_style(template = "varname",
                                     halign = "center",
                                     decoration = "bold",
                                     fontName = "Cambria",
                                     fontSize = 14)

# Style auf Zelle anwenden
addStyle(wb = wb,
         sheet = "TEST",
         style = varname_bold_left,
         rows = 3,
         cols = 2)


# WB speichern
save_tg_workbook(wb,
                 filename = "test_07.xlsx",
                 tg_header = TRUE,
                 overwrite = TRUE)


```
![Excel Kopfzeile](https://github.com/ogdtg/TGexcel/blob/main/img/11_style.PNG)


### Einzelne Features hinzufügen

Um einzelne Features wie zum Beispiel die Hintergrundfarbe oder die Schriftart zu verändern, ohne aber den Rest des vorhanden Celsstyles in einer Zelle zu überschreiben kann die `add_style_feature` Funktion und die daraus abgeleiteten Funktionen verwendet werden. Neben einem Workbookobject `wb` und dem Namen des Tabellenblatts `sheet` muss ein vector mit den zu verändernden Features, sowie ein Vektor mit den entsprechenden Werten angegeben werden. Hierbei ist es wichtige die korrekte Formatierung der Werte zu beachten, da ansonsten das xlsx File nicht korrekt abgespeichert werden kann. Es ist daher ratsam, die von der Funktion abgeleitetn Funktionen zu verwenden, die die korrekte Formatierung gewährleisten. Mehr Informationen zu den möglichen Features finden sich in der Dokumentation von `openxlsx::createStyle()`.


```r
wb <- loadWorkbook("test_11.xlsx")

# Funktion ergänzt eine blaue, horizontale Linie unter der 10. und 15. Zeile, 
# die sich über 9 Spalten erstreckt
add_style_feature(
  wb,
  sheet,
  rows = c(10,15),
  cols =  c(1:9),
  c("borderBottom", "borderBottomColour"),
  c("thin", list(rgb = "FF0000FF"))
)

save_tg_workbook(wb, "test_12.xlsx", overwrite = T, tg_header = F)

```
#### Horizontale Linie

Das oben Beschriebene kann auch durch `add_hline` erreicht werden, welche von `add_style_feature` abgeleitet wurde und nutzerfreundlicher ist.

```r
wb <- loadWorkbook("test_11.xlsx")

# Funktion ergänzt eine blaue, horizontale Linie unter der 10. und 15. Zeile, 
# die sich über 9 Spalten erstreckt
add_hline(wb,
          sheet,
          rows = c(10,15),
          ncol = 9,
          style = "thin",
          color = "blue",
          below = TRUE)


save_tg_workbook(wb, "test_12.xlsx", overwrite = T, tg_header = F)
```
#### Font fett, kursiv etc.

Mit `add_decoration` kann der Text in einer oder mehrere Zellen fett, kursiv o.ä. gemacht werden (vgl. `openxlsx::createStyle()` für verfügbare Style Methoden)

```r

wb <- loadWorkbook("test_11.xlsx")

# Macht den Text jeweils in der ersten, dritten und neunten Spalte
# der zehnten, 20. und 30. Reihe fett
add_decoration(
  wb,
  sheet,
  rows = c(10, 20, 30),
  cols = c(1, 3, 9),
  decoration = "BOLD"
)

save_tg_workbook(wb, "test_12.xlsx", overwrite = T, tg_header = F)

```
#### Füllfarbe der Zelle

Die Füllung einer Zelle kann mit `add_bg` verändert werden.

```r
wb <- loadWorkbook("test_11.xlsx")

# Setzt die Füllung in der zwölften Reihe für die
# ersten neun Spalten auf himmelblau (wie in Internettabellen verwendet) 
add_bg(wb,
       sheet,
       rows = 12,
       cols = c(1:9),
       color = "#c5d9f1"
)

save_tg_workbook(wb, "test_12.xlsx", overwrite = T, tg_header = F)

```


## Weitere Funktionen

### Spezielle Zahlenformate einfügen

Existieren in einer Tabelle Zahlen die auf eine bestimmte Art gekennzeichnet werden sollen, z.B. mit `**` weil es sich um weniger als x Beochatungen handelt so kann dies mithilfe eines speziellen Zahlenformats in Excel bewerkstelligt werden ohne, dass die Zahl ihre Eigenschaft als Zahl verliert. Sofern in R ein Datensatz besteht, indem die entsprechenden Zellen bereits gekennzeichnet sind, kann die `add_special_numFmt` Funktion verwendet werden.

Die Funktion sucht in einem gegebenen Datensatz nach einem bestimmten `pattern`, merkt sich die Position der matchenden Zellen, schreibt die Daten in das Workbook und wendet am Ende das gegebene Zahlenformat auf die enstprechenden Zellen an.

**ACHTUNG**: Die Funktion kann NICHT auf Zellen angewendet werden, die keinerlei Styling enthalten.

**Beispiel:**

Wir haben folgenden Datensatz namens `test`:

![Dataset Stars](https://github.com/ogdtg/TGexcel/blob/main/img/19_df_stars.PNG)

Diesen möchten wir nun in ein Excel File schreiben. Anstelle der Sternchen möchten wir eckige Klammern um die Zahlen herum. In der Excel soll jedoch kein Text eingetragen werden, sondern das Zahlenformat soll beibehalten werden. Angenommen, die entsprechenden Zellen in der Excel Datei enthalten mindestens ein Style Attribut, dann gehen wir wie folgt vor:



````r
add_special_numfmt(wb=wb,
                   sheet = "test",
                   data = test,
                   pattern = "\\*\\*",
                   dataStart = 1,
                   startCol = 1,
                   prefix = "\\[",
                   suffix = "\\]")

```

Wichtig sind die anführenden `\\`, welche vor *jedem einzelnen* character eingefügt werden müssen.

Das Ergebnis sieht dann wie folgt aus:

![Dataset Stars](https://github.com/ogdtg/TGexcel/blob/main/img/20_excel_stars.PNG)


### Fussnoten und Datenquelle

Mit `add_footnotes` können Fussnoten, mit `add_datenquelle` die Datenquelle hinzugefügt werden

```r

wb <- loadWorkbook("test_05.xlsx")

add_footnotes(
  wb = wb,
  sheet = "TEST",
  startRow = 33,
  footnotes = c("1 Neue Footnote", "2 Neue Footnote")
)

add_datenquelle(
  wb = wb,
  sheet = "TEST",
  startRow = 35,
  datenquelle =  "Neue Datenquelle"
)

# WB speichern
save_tg_workbook(wb,
                 filename = "test_08.xlsx",
                 tg_header = TRUE,
                 overwrite = TRUE)

```
![Excel Fussnote](https://github.com/ogdtg/TGexcel/blob/main/img/12_footnote.PNG)

### Hochgestellte und tiefgestellte Ziffern

Hoch- und tiefgestellte Zahlen können mit`[Zahl]` bzw. `~Zahl~` erstellt werden.

```r
wb <- createWorkbook()
addWorksheet(wb,"Test")

create_header_style(wb = wb, sheet = "Test", ncol = 6, text = "Hochgestellt[1] Tiefgestellt~2~")

saveWorkbook(wb, "test_09.xlsx", overwrite = TRUE)

```

![Excel Fussnote](https://github.com/ogdtg/TGexcel/blob/main/img/13_hochtief.PNG)

### NAs entfernen

Beim Abspeichern kann es sein, dass Leere Zellen mit ungewollten NAs befüllt werden.

```r
# Pfad definieren
path <- "Y:\\SK\\SKStat\\Internet\\1_Statistik\\1_Themen und Daten\\1_Bevölkerung und Haushalte\\1_01_Bevölkerungsstand\\Tabellen\\"

# Workbook laden
wb <- loadWorkbook(paste0(path,"2022_2015_BevGmd_Geschl_Nat.xlsx"))

save_tg_workbook(wb,"test_10.xlsx")

```
![Excel NA](https://github.com/ogdtg/TGexcel/blob/main/img/14_na.PNG)


Zu diesem Zwecke kann die `replace_nas_in_wb` verwendet werden.

```r
# Pfad definieren
path <- "Y:\\SK\\SKStat\\Internet\\1_Statistik\\1_Themen und Daten\\1_Bevölkerung und Haushalte\\1_01_Bevölkerungsstand\\Tabellen\\"

# Workbook laden
wb <- loadWorkbook(paste0(path,"2022_2015_BevGmd_Geschl_Nat.xlsx"))

# NAs ersetzen
replace_nas_in_wb(wb)

save_tg_workbook(wb,"test_10.xlsx")

```

![Excel noNA](https://github.com/ogdtg/TGexcel/blob/main/img/15_nona.PNG)

