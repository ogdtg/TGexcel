% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/insert_column.R
\name{insert_column}
\alias{insert_column}
\title{insert_column}
\usage{
insert_column(
  wb,
  sheet,
  patternCol = NULL,
  col,
  endCol = NULL,
  startRow = 1,
  endRow = NULL,
  dataStart,
  data,
  colNames = FALSE
)
}
\arguments{
\item{wb}{a workbook object}

\item{sheet}{sheet name}

\item{patternCol}{column from which the styles will be copied (default is NULL -> last column of table)}

\item{col}{column index where to insert data}

\item{endCol}{column where to end style copying (default = 200)}

\item{startRow}{row index where to start copying the styles (default = 1)}

\item{endRow}{row index where to end copying the styles (default is NULL -> endrow = nrow(data)+dataStart-1)}

\item{dataStart}{row index where the data starts}

\item{data}{data.frame (ncol must be 1) to insert}

\item{colNames}{if TRUE, colNames of data will be included (default is FALSE)}
}
\description{
Function to insert a data column somewhere in the table.
}
