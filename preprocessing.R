library(tidyverse)
library(readxl)
library(readODS)
library(stringr)
excel_sheets("data/Defense data.xlsx") 
raw_dat<-list(prices=read_excel("data/Defense data.xlsx",sheet = "Sheet2"),
              mkt_cap=read_excel("data/Defense data.xlsx",sheet = "Sheet2"))
saveRDS(raw_dat,file = "data/raw_dat.rds")

excel_sheets("data/All figures return spillovers.xlsx")->sheets
spillover_rtns<-vector("list",10)
names(spillover_rtns)<-sheets
for (i in sheets) {
print(i)
  if (i=="TCI") {
    spillover_rtns[[i]] <- read_excel("data/All figures return spillovers.xlsx",sheet = i) %>%
      pivot_longer(!Date,names_to = 'Stock', values_to = 'Spillover')
  } else {
spillover_rtns[[i]] <- read_excel("data/All figures return spillovers.xlsx",sheet = i) %>%
  rename(Date=`...1`) %>%
  pivot_longer(!Date,names_to = 'Stock', values_to = 'Spillover')
}
}
saveRDS(spillover_rtns,file = "data/spillover_rtns.rds")

excel_sheets("data/All figures volatility spillovers.xlsx")->sheets
spillover_vols<-vector("list",10)
names(spillover_vols)<-sheets
for (i in sheets) {
  print(i)
  if (i=="TCI") {
    spillover_vols[[i]] <- read_excel("data/All figures volatility spillovers.xlsx",sheet = i) %>%
      pivot_longer(!Date,names_to = 'Stock', values_to = 'Spillover')
  } else {
    
  spillover_vols[[i]] <- read_excel("data/All figures volatility spillovers.xlsx",sheet = i) %>%
    rename(Date=`...1`) %>%
    pivot_longer(!Date,names_to = 'Stock', values_to = 'Spillover')
  }
}
saveRDS(spillover_vols,file = "data/spillover_vols.rds")

# library(plotly)
# plot_ly(data = spillover_rtns_long,mode="lines") %>%
#  add_trace(x=~Date,y=~Spillover,color=~Stock)

read_excel("data/All figures return spillovers.xlsx",sheet = "TCI") %>%
  pivot_longer(!Date,names_to = 'Stock', values_to = 'Spillover')

library(docxtractr)
docx<-read_docx("draft.docx")
docx_extract_all(docx)->all_tbls

all_tbls[c(5:10)] %>% 
  map(~assign_colnames(.x,row = 1)) %>%
  map(~(.x %>% 
          rename(From='') %>%
          pivot_longer(!From,names_to = "To") %>%
          rename(spillover=value) %>%
          mutate(From=str_to_title(From),
                 To=str_to_title(To)))) -> static
names(static) <- c("rtn50","rtn95","rtn5","vol50","vol95","vol5")
saveRDS(static, "data/static.rds")

library(igraph)
static$rtn95 %>% 
  filter(!(From %in% c("Net","To Others","Including Own","From Others"))) %>%
  filter(!(To %in% c("Net","To Others","Including Own","From Others"))) -> links
static$rtn95 %>% 
  filter(From %in% c("Net","To Others","Including Own","From Others")) %>%
  filter(To %in% c("Net","To Others","Including Own","From Others")) -> nodes
    graph_from_data_frame(directed = T) -
# net <- simplify(net, remove.multiple = F, remove.loops = T) 
net<-simplify(net, edge.attr.comb=list(Weight="sum","ignore"))
ggnet2(net,label = T,label.size = 2,edge.alpha = 0.4)

library(GGally)
library(network)
library(sna)
library(RColorBrewer)

# random graph
networ
net = rgraph(10, mode = "graph", tprob = 0.5)
net = network(net, directed = FALSE)

# vertex names
network.vertex.names(net) = letters[1:10]
