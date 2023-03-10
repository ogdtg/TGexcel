% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/copy_body_data.R
\name{copy_body_data}
\alias{copy_body_data}
\title{copy_body_data}
\usage{
copy_body_data(wb, sheet, newSheet, dataStartRow, year)
}
\arguments{
\item{wb}{a workbook object}

\item{sheet}{the name of the Sheet in the workbook object from which to copy the data}

\item{newSheet}{new sheet where the data should be copied}

\item{dataStartRow}{row where the body data starts}

\item{year}{index of the column where a year variable is located}
}
\description{
Copy the data body from one sheet to another
}
