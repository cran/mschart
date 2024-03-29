library(mschart)
library(officer)
library(xml2)

test <- browser_data
test[test$browser=='Opera', 1] <- '\U0001f47d'

chart_01 <- ms_barchart(data = test, x = "browser", y = "value", group = "serie")
chart_01 <- chart_labels(chart_01, title = "\U0001f47d", xlab = "\U0001f47d")

xml <- format(
  chart_01,
  sheetname = "sheet1",
  id_x = "64451212",
  id_y = "64453248"
)

chart <- read_xml(xml)
txt <- xml_find_all(chart, "//c:title/c:tx/c:rich/a:p/a:r/a:t")

expect_equivalent(xml_text(txt), c("\U0001f47d", "\U0001f47d"))


txt_data <- xml_find_all(chart, "//c:ser[1]/c:cat/c:strRef/c:strCache/c:pt")
expect_equivalent(xml_text(txt_data), c("Android", "Chrome", "Firefox", "IE", "Safari", "\U0001f47d"))
