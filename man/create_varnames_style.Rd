% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/create_varnames_style.R
\name{create_varnames_style}
\alias{create_varnames_style}
\title{create_varnames_style}
\usage{
create_varnames_style(
  wb,
  sheet,
  var_names = NULL,
  startRow = 3,
  style = varname_normal_10_left,
  ...
)
}
\arguments{
\item{wb}{a workbook object}

\item{sheet}{the name of the Sheet in the workbook object}

\item{var_names}{vector of variable names}

\item{startRow}{row in which the subheader should be placed (default = 2)}

\item{style}{an openxlsx::Style element. Default is current "Internettabellen" variable names style}

\item{ncol}{only necessary if var_names are NULL (style_only). Provide a positive integer.}
}
\description{
Function to insert and style variable names in the style of the Internettabellen
}
