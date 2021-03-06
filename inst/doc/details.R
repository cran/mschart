## ----'setup', include=FALSE, echo=FALSE, message=FALSE, warning=FALSE----
knitr::opts_chunk$set(
  collapse = TRUE,
  comment = "#>"
)


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

base_url <- "https://ardata-fr.github.io/mschart/articles/"

## ------------------------------------------------------------------------
library(mschart)
library(officer)
library(magrittr)

## ------------------------------------------------------------------------
gen_pptx <- function( chart, file ){
  doc <- read_pptx()
  doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
  doc <- ph_with(doc, value = chart, location = ph_location_type(type = "body"))
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
linec <- ms_linechart(data = iris, x = "Sepal.Length",
                      y = "Sepal.Width", group = "Species")
linec_smooth <- chart_data_smooth(linec,
                values = c(virginica = 1, versicolor = 0, setosa = 0) )
gen_pptx(linec_smooth, file = "assets/pptx/gallery_line_04.pptx")

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


## ------------------------------------------------------------------------
my_bc_01 <- ms_barchart(
    data = browser_data, x = "browser", y = "value", group = "serie") %>% 
  chart_settings( dir = "vertical", grouping = "stacked", overlap = 100 )

my_bc_02 <- ms_barchart(
    data = browser_data, x = "browser", y = "value", group = "serie") %>% 
  chart_settings( dir = "vertical", grouping = "clustered", 
                  gap_width = 400, overlap = -100 )

my_sc_01 <- ms_scatterchart(
  data = mtcars, x = "disp", y = "drat") %>% 
  chart_settings(scatterstyle = "marker")

my_sc_02 <- ms_scatterchart(
  data = mtcars, x = "disp", y = "drat") %>% 
  chart_settings(scatterstyle = "lineMarker")

## ----results='hide'------------------------------------------------------
layout <- "Title and Content"
master <- "Office Theme"
read_pptx() %>% 
  add_slide(layout, master) %>% 
  ph_with(my_bc_01, location = ph_location_fullsize()) %>% 
  add_slide(layout, master) %>% 
  ph_with(my_bc_02, location = ph_location_fullsize()) %>% 
  add_slide(layout, master) %>% 
  ph_with(my_sc_01, location = ph_location_fullsize()) %>% 
  add_slide(layout, master) %>% 
  ph_with(my_sc_02, location = ph_location_fullsize()) %>% 
  print(target = "assets/pptx/chart_settings_01.pptx")

## ----echo=FALSE----------------------------------------------------------
office_doc_link( url = paste0( base_url, "assets/pptx/theme_01.pptx" ) )

## ------------------------------------------------------------------------
my_bc <- ms_barchart(
  data = browser_data, x = "browser", y = "value", group = "serie") %>% 
  as_bar_stack( ) %>% 
  chart_labels(title = "Title example", xlab = "x label", ylab = "y label")

## ----results='hide'------------------------------------------------------
gen_pptx(my_bc, file = "assets/pptx/theme_00.pptx")

## ----echo=FALSE----------------------------------------------------------
do.call(htmltools::tags$ul, 
lapply( sort(formalArgs(mschart_theme)), 
        function( x ) {
          htmltools::tags$li(x)
        } ) )

## ------------------------------------------------------------------------
mytheme <- mschart_theme(
  axis_title = fp_text(color = "#0D6797", font.size = 20, italic = TRUE),
  main_title = fp_text(color = "#0D6797", font.size = 24, bold = TRUE),
  grid_major_line = fp_border(color = "#AA9961", style = "dashed"),
  axis_ticks = fp_border(color = "#AA9961") 
  )
my_bc <- set_theme(my_bc, mytheme)
gen_pptx(my_bc, file = "assets/pptx/theme_01.pptx")

## ------------------------------------------------------------------------
my_bc <- chart_theme( x = my_bc, 
  axis_title_x = fp_text(color = "red", font.size = 11),
  main_title = fp_text(color = "red", font.size = 15),
  legend_position = "t" )
gen_pptx(my_bc, file = "assets/pptx/theme_02.pptx")

