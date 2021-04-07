mschart R package
================

<!-- README.md is generated from README.Rmd. Please edit that file -->

[![R build
status](https://github.com/ardata-fr/mschart/workflows/R-CMD-check/badge.svg)](https://github.com/ardata-fr/mschart/actions)
[![version](https://www.r-pkg.org/badges/version/mschart)](https://CRAN.R-project.org/package=mschart)
![cranlogs](https://cranlogs.r-pkg.org/badges/mschart)
![Active](http://www.repostatus.org/badges/latest/active.svg)

The `mschart` package provides a framework for easily create charts for
‘Microsoft PowerPoint’ documents. It has to be used with package
[`officer`](https://davidgohel.github.io/officer/) that will produce the
charts in new or existing PowerPoint or Word documents.

![](https://www.ardata.fr/img/illustrations/ms_barchart.png)

## Example

This is a basic example which shows you how to create a line chart.

``` r
library(mschart)
library(officer)

linec <- ms_linechart(data = iris, x = "Sepal.Length",
                      y = "Sepal.Width", group = "Species")
linec <- chart_ax_y(linec, num_fmt = "0.00", rotation = -90)
```

Then use package `officer` to send the object as a chart.

``` r
doc <- read_pptx()
doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
doc <- ph_with(doc, value = linec, location = ph_location_type(type = "body"))

print(doc, target = "example.pptx")
```

At any moment, you can type `print(your_chart, preview = TRUE)` to
preview the chart in a temporary PowerPoint file. This requires to have
a PowerPoint Viewer installed on the machine.

## Installation

You can get the development version from GitHub:

``` r
devtools::install_github("ardata-fr/mschart")
```

Or the latest version on CRAN:

``` r
install.packages("mschart")
```

## Contributing to the package

### Bug reports

When you file a [bug
report](https://github.com/ardata-fr/mschart/issues), please spend some
time making it easy for me to follow and reproduce. The more time you
spend on making the bug report coherent, the more time I can dedicate to
investigate the bug as opposed to the bug report.

### Contributing to the package development

A great way to start is to contribute an example or improve the
documentation.

If you want to submit a Pull Request to integrate functions of yours,
please provide:

  - the new function(s) with code and roxygen tags (with examples)
  - a new section in the appropriate vignette that describes how to use
    the new function
  - add corresponding tests in directory `tests/testthat`.

By using rhub (run `rhub::check_for_cran()`), you will see if everything
is ok. When submitted, the PR will be evaluated automatically on travis
and appveyor and you will be able to see if something broke.
