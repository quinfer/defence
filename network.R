rm(list = ls())
library(GGally)
library(network)
library(sna)
library(RColorBrewer)
library(igraph)
library(docxtractr)
library(tidyverse)
static <- readRDS("data/static.rds")
static$rtn95 %>% 
  filter(!(From %in% c("Net","To Others","Including Own","From Others"))) %>%
  filter(!(To %in% c("Net","To Others","Including Own","From Others"))) %>%
  drop_na()-> links
links$spillover <- as.numeric(links$spillover)
static$rtn95 %>% 
  filter(From =="Including Own") %>%
  rename(name=To,spillover_including_own=spillover) %>%
  select(!From) %>%
  left_join(
static$rtn95 %>% 
  filter(To =="From Others") %>%
  rename(name=From,spillover_from_others=spillover) %>%
  select(!To),by="name") %>%
  mutate(spillover=as.numeric(spillover_including_own)+as.numeric(spillover_from_others)) -> nodes
docx<-read_docx("draft.docx")
docx_extract_all(docx)->all_tbls
 all_tbls[[11]] %>% assign_colnames(row=1) -> info
 bind_cols(
  nodes %>% arrange(name),
  info %>% arrange(`Company Name`)
) %>%
  mutate(`Company Name`=str_remove_all(`Company Name`," Corp| Co| SE| Inc| PLC| SA| Ltd| AG"))->nodes
nodes$name <- nodes$`Company Name`
links %>%
  left_join(nodes %>% select(name,`Company Name`),
            by=c("From"="name")) %>%
  mutate(From=`Company Name`) %>%
  select(!`Company Name`) %>%
  left_join(nodes %>% select(name,`Company Name`),
            by=c("To"="name")) %>%
  mutate(To=`Company Name`) %>%
  select(!`Company Name`) -> links
# links %>% filter
net <- graph_from_data_frame(d=links, vertices=nodes, directed=T) 
net <- igraph::simplify(net, remove.multiple = F, remove.loops = T) 
# Generate colors based on media type:
unique(nodes$Country)
colrs <- c("gray50", "tomato", "gold", "red","brown","green")
V(net)$color <- colrs[V(net)$Country]
V(net)$size <- V(net)$spillover
E(net)$width <- E(net)$spillover
V(net)$deg
ggnet2(net,label = T,label.size = 2,edge.alpha = 0.4, color = "Country",size = "spillover")
