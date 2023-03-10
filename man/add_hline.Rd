% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/add_hline.R
\name{add_hline}
\alias{add_hline}
\title{add_hline}
\usage{
add_hline(
  wb,
  sheet,
  rows,
  ncol,
  color,
  style = "thin",
  below = TRUE,
  startCol = 1
)
}
\arguments{
\item{wb}{a workbook object}

\item{sheet}{sheet name}

\item{rows}{rows the the line should be placed in}

\item{ncol}{number of columns the line should embrace}

\item{color}{color of the line (Rcolor or HEX Value)}

\item{style}{see [openxlsx::createStyle()]}

\item{below}{logical whether line should be placed below or above selected rows (default is `TRUE`)}

\item{startCol}{column where the line should start (default is 1)}
}
\description{
Adds a horizontal line by setting cell borders
}
