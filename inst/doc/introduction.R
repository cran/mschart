## ----'setup', echo = FALSE, message=FALSE, warning=FALSE-----------------
dir.create("assets/pptx", recursive = TRUE, showWarnings = FALSE)
dir.create("assets/docx", recursive = TRUE, showWarnings = FALSE)
office_doc_link <- function(url){
  stopifnot(requireNamespace("htmltools", quietly = TRUE))
  htmltools::tags$p(  htmltools::tags$span("Download file "),
    htmltools::tags$a(basename(url), href = url), 
    htmltools::tags$span(" - view with"),
    htmltools::tags$a("office web viewer", target="_blank", 
      href = paste0("https://view.officeapps.live.com/op/view.aspx?src=", url)
      ), 
    style="text-align:center;font-style:italic;color:gray;"
    )
}

## ------------------------------------------------------------------------
library(mschart)
library(officer)
library(magrittr)

## ------------------------------------------------------------------------
my_barchart <- ms_barchart(data = browser_data, 
  x = "browser", y = "value", group = "serie")

## ----results='hide'------------------------------------------------------
library(officer)
doc <- read_pptx()
doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
doc <- ph_with_chart(doc, chart = my_barchart)
print(doc, target = "assets/pptx/barchart_01_stacked.pptx")

## ----echo=FALSE----------------------------------------------------------
office_doc_link( url = paste0( "https://ardata-fr.github.io/mschart/articles/", "assets/pptx/barchart_01_stacked.pptx" ) )

## ----results='hide'------------------------------------------------------
doc <- read_docx()
doc <- body_add_chart(doc, chart = my_barchart, style = "centered")
print(doc, target = "assets/docx/barchart_01_stacked.docx")

## ----echo=FALSE----------------------------------------------------------
office_doc_link( url = paste0( "https://ardata-fr.github.io/mschart/articles/", "assets/docx/barchart_01_stacked.docx" ) )

## ------------------------------------------------------------------------
my_barchart <- chart_settings( my_barchart, grouping = "stacked", gap_width = 50, overlap = 100 )

## ------------------------------------------------------------------------
my_barchart <- chart_ax_x(my_barchart, cross_between = 'between', 
  major_tick_mark = "in", minor_tick_mark = "none")
my_barchart <- chart_ax_y(my_barchart, num_fmt = "0.00", rotation = -90)

## ------------------------------------------------------------------------
my_barchart <- chart_labels(my_barchart, title = "A main title", 
  xlab = "x axis title", ylab = "y axis title")

## ------------------------------------------------------------------------
my_barchart <- chart_data_fill(my_barchart,
  values = c(serie1 = "#003C63", serie2 = "#ED1F24", serie3 = "#F2AA00") )
my_barchart <- chart_data_stroke(my_barchart, values = "transparent" )

## ------------------------------------------------------------------------
title_style <- fp_text(color = "#175FAA", font.size = 14, bold = TRUE)
mytheme <- mschart_theme(main_title = update( title_style, font.size = 18),
  axis_title_x = title_style, axis_title_y = title_style,
  grid_major_line_y = fp_border(width = 1, color = "orange"),
  axis_ticks_y = fp_border(width = 1, color = "orange") )
my_barchart <- set_theme(my_barchart, mytheme)

## ----results='hide'------------------------------------------------------
library(officer)
doc <- read_pptx()
doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
doc <- ph_with_chart(doc, chart = my_barchart)
print(doc, target = "assets/pptx/barchart_02_stacked.pptx")

## ----echo=FALSE----------------------------------------------------------
office_doc_link( url = paste0( "https://ardata-fr.github.io/mschart/articles/", "assets/pptx/barchart_02_stacked.pptx" ) )

## ------------------------------------------------------------------------
gen_pptx <- function( chart, file ){
  doc <- read_pptx()
  doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
  doc <- ph_with_chart(doc, chart = chart)
  print(doc, target = file)
  office_doc_link( url = paste0( "https://ardata-fr.github.io/mschart/articles/", file ) )
}


## ------------------------------------------------------------------------
my_barchart_01 <- ms_barchart(
  data = browser_data, x = "browser",
  y = "value", group = "serie"
)
gen_pptx(my_barchart_01, file = "assets/pptx/gallery_bar_01.pptx")

## ------------------------------------------------------------------------
my_barchart_02 <- as_bar_stack(my_barchart_01)
gen_pptx(my_barchart_02, file = "assets/pptx/gallery_bar_02.pptx")

## ------------------------------------------------------------------------
my_barchart_03 <- as_bar_stack(my_barchart_01, dir = "horizontal")
gen_pptx(my_barchart_03, file = "assets/pptx/gallery_bar_03.pptx")

## ------------------------------------------------------------------------
my_barchart_04 <- as_bar_stack(my_barchart_01, dir = "horizontal", 
  percent = TRUE)
gen_pptx(my_barchart_04, file = "assets/pptx/gallery_bar_04.pptx")

## ------------------------------------------------------------------------
data <- structure(
  list(
    supp = structure( c(1L, 1L, 1L, 2L, 2L, 2L), 
                      .Label = c("OJ", "VC"), class = "factor" ),
    dose = c(0.5, 1, 2, 0.5, 1, 2),
    length = c(13.23, 22.7, 26.06, 7.98, 16.77, 26.14)
  ),
  .Names = c("supp", "dose", "length"),
  class = "data.frame", row.names = c(NA,-6L)
)

lc_01 <- ms_linechart(data = data, x = "dose",
                      y = "length", group = "supp")
gen_pptx(lc_01, file = "assets/pptx/gallery_line_01.pptx")

## ------------------------------------------------------------------------
lc_02 <- ms_linechart(data = browser_ts, x = "date",
                      y = "freq", group = "browser")
lc_02 <- chart_ax_x(lc_02, num_fmt = "m/d/yy")
gen_pptx(lc_02, file = "assets/pptx/gallery_line_02.pptx")

## ------------------------------------------------------------------------
lc_03 <- chart_settings(lc_02, grouping = "percentStacked")
gen_pptx(lc_03, file = "assets/pptx/gallery_line_03.pptx")

## ------------------------------------------------------------------------
ac_01 <- ms_areachart(data = data, x = "dose",
                      y = "length", group = "supp")
gen_pptx(ac_01, file = "assets/pptx/gallery_area_01.pptx")

ac_02 <- ms_areachart(data = browser_ts, x = "date",
                      y = "freq", group = "browser")
ac_02 <- chart_ax_y(ac_02, cross_between = "between", num_fmt = "General")
ac_02 <- chart_ax_x(ac_02, cross_between = "midCat", num_fmt = "m/d/yy")
gen_pptx(ac_02, file = "assets/pptx/gallery_area_02.pptx")

ac_03 <- chart_settings(ac_02, grouping = "percentStacked")
gen_pptx(ac_03, file = "assets/pptx/gallery_area_03.pptx")

## ------------------------------------------------------------------------
sc_01 <- ms_scatterchart(data = mtcars, x = "disp",
                         y = "drat", group = "gear")
gen_pptx(sc_01, file = "assets/pptx/gallery_scatter_01.pptx")

sc_02 <- ms_scatterchart(data = mtcars, x = "disp",
                         y = "drat", group = "gear")
gen_pptx(sc_02, file = "assets/pptx/gallery_scatter_02.pptx")


