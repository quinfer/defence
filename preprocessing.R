library(tidyverse)
library(readxl)
library(readODS)
library(stringr)
library(docxtractr)
rm(list = ls())
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


docx<-read_docx("draft.docx")
docx_extract_all(docx)->all_tbls
all_tbls[[11]] %>% assign_colnames(row=1)
nodes <-readRDS("data/index.rds")
all_tbls[c(5:10)] %>% 
  map(~assign_colnames(.x,row = 1)) %>%
  map(~(.x %>% 
          rename(From='') %>%
          pivot_longer(!From,names_to = "To") %>%
          rename(spillover=value) %>%
          mutate(From=str_to_title(From),
                 To=str_to_title(To)) %>%
          left_join(nodes %>% select(name,`Company Name`),
                    by=c("From"="name")) %>%
          mutate(From=ifelse(!is.na(`Company Name`),`Company Name`,From)) %>%
          select(!`Company Name`) %>%
          left_join(nodes %>% select(name,`Company Name`),
                    by=c("To"="name")) %>%
          mutate(To=ifelse(!is.na(`Company Name`),`Company Name`,To)) %>%
          select(!`Company Name`))) -> static
names(static) <- c("rtn50","rtn95","rtn5","vol50","vol95","vol5")
unique(static$rtn95$To)
saveRDS(static, "data/static.rds")

## Add names and country
ocx<-read_docx("draft.docx")
docx_extract_all(docx)->all_tbls
all_tbls[[11]] %>% assign_colnames(row=1) -> info
spill_rtn$to50 %>% distinct(Stock) %>% 
  arrange(Stock) %>%
  bind_cols(info %>% arrange(`Company Name`)) %>%
  rename(name=`Company Name`) %>%
  mutate(name=str_remove_all(name,
                             " Corp| Co| SE| Inc| PLC| SA| Ltd| AG"))-> index
spill_rtn %>% map(~left_join(.x,index,by="Stock")) -> spillover_rtns
spill_vol %>% map(~left_join(.x,index,by="Stock")) -> spillover_vols
saveRDS(spillover_vols,file = "data/spillover_vols.rds")
saveRDS(spillover_rtns,file = "data/spillover_rtns.rds")
