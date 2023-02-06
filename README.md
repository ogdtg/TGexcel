# TGexcel

<!-- badges: start -->
[![Lifecycle: experimental](https://img.shields.io/badge/lifecycle-experimental-orange.svg)](https://lifecycle.r-lib.org/articles/stages.html#experimental)
<!-- badges: end -->

R Package um Internettabellen aus R anzulegen und zu formatieren. TGexcel ist eine Erweiterung des openxlsx Package und basiert weitestgehend auf den Funktionen aus openxlsx.

Das Package soll es dem User möglich machen die Erstellung der Internettabellen direkt aus einem R-Prozess heraus durchzuführen. So können Prozesse weiter standardisiert und automatisiert werden.

Das Package macht es möglich. Titel, Subtitel und Variablennamen (auch verschachtelt) im entsprechenden Format anzulegen. Ausserdem können auch die Daten selbst im gewünschten Format eingetragen werden. Auch die Kopfzeile mit Logo kann per Funktion eingefügt werden. Es können sowohl neue Excel Files angelegt erden als auch bereits existierende Files bearbeitet werden.

Da das Package auf `openxlsx` basiert, werden für alle Funktionen workbook-Objekte verwendet. Diese können entweder mit `createWorkbook()` neu erstellt, oder mit `loadWorkbook("path/to/file.xlsx")` bearbeitet werden.

![alt text](http://url/to/img.png)


## Installation

Die Development Version des Packages kann wie folgt installiert und genutzt werden:

``` r
devtools::install_github("ogdtg/TGexcel")
library(TGexcel)
```


## Titel (header)

Mit der 

