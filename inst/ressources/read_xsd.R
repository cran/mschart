library(xml2)
library(magrittr)
library(tidyverse)
library(purrr)
library(stringr)
chart_spec <- read_xml("inst/tools/dml-chart.xsd")


# simple types -----
elts <- xml_find_all(chart_spec, "//xsd:simpleType[xsd:restriction[@base='xsd:string']/xsd:enumeration]")
r6_names <- xml_attr( elts, "name" ) %>% str_to_lower()

str <- "%s <- c(%s)\n"

choices <- sapply( elts, function(x) {
  values <- xml_attr( xml_children(xml_child(x, "xsd:restriction")), "value")
  paste0(shQuote(values), collapse = ", ")
})

# cat( sprintf(str, r6_names, choices), sep = "\n", file = "R/simple_types.R" )




# complex types ------
ct <- xml_find_all(chart_spec, "//xsd:complexType")
has_sequence <- map_lgl(ct, function(x) "sequence" %in% xml_name(xml_children(x)) )
has_attribute <- map_lgl(ct, function(x) "attribute" %in% xml_name(xml_children(x)) )
r6_names <- xml_attr( ct, "name" ) %>% str_to_lower()
ct[has_attribute]

xml_find_first(chart_spec, "/xsd:schema/xsd:group[@name='EG_AxShared']") %>% as.character()





sequence_df <- function(id, chart_spec){
  xpath_ <- sprintf("//*[@name='%s']/xsd:sequence/xsd:element", id)
  elts <- xml_find_all(chart_spec, xpath_)
  tibble( name = elts %>% xml_attr("name"),
          node_type = rep("node", length(elts)),
          mandatory = (elts %>% xml_attr("minOccurs")) %in% "1",
          max = as.integer(elts %>% xml_attr("maxOccurs")),
          type = elts %>% xml_attr("type") )
}

attributes_df <- function(id, chart_spec){
  xpath_ <- sprintf("//*[@name='%s']/xsd:attribute", id)
  elts <- xml_find_all(chart_spec, xpath_)
  tibble( name = elts %>% xml_attr("name"),
          node_type = rep( "attribute", length(elts)),
          mandatory = ifelse( (elts %>% xml_attr("use")) %in% c("require"), TRUE, FALSE ),
          default = elts %>% xml_attr("default"),
          type = elts %>% xml_attr("type") )
}

EG_AxShared <- sequence_df("EG_AxShared", chart_spec)

attributes_df("CT_AxPos", chart_spec)
sequence_df("EG_AxShared", chart_spec)

as_df <- function(id, chart_spec){
  nodes_df <- sequence_df(id, chart_spec)
  attrs_df <- attributes_df(id, chart_spec)
  bind_rows(nodes_df, attrs_df)
}

types <- unique(EG_AxShared$type)
types <- set_names(types, types)
level1_par <- map_df(types, as_df, chart_spec, .id = "st_type")

types_ <- unique(level1_par$type)
types_ <- set_names(types_, types_)
map_df(types_, as_df, chart_spec, .id = "st_type")

