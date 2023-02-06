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
  data,
  datenquelle = "BEVNAT Kanton Thurgau",
  footnotes = c("1 Test der Funktion", "2 alle Funktionen zusammengefasst"),
  gemeinde_format = FALSE,
  nesting = c(1, 3, 2),
  year = 6
)

saveWorkbook(wb, "test_05.xlsx", overwrite = T)

```
![Excel Wrapper](https://github.com/ogdtg/TGexcel/blob/main/img/09_wrapper.PNG)