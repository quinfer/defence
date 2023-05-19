library(rio)
library(tidyverse)
dat<-import('data/refinitiv22.2.23.xlsx')
dat<-dat[,c(1:3,5:7)]
names(dat)<-c("RIC","Name","MktCap","RevenueFY0","RevenueFY1","CountryHQ")
dat %>%
  mutate(total_mc=sum(MktCap,na.rm = T),
         mkt_cap_share=MktCap/total_mc,
         global_mktcap=118388540621399,
         total_rev_fy0=sum(RevenueFY0,na.rm = T),
         total_rev_fy1=sum(RevenueFY1,na.rm = T),
         rev_share=RevenueFY0/total_rev_fy0,
         rev_growth=(RevenueFY0-RevenueFY1)/RevenueFY1,
         total_rev_growth=(total_rev_fy0-total_rev_fy1)/total_rev_fy1)->dat
saveRDS(dat,"data/refinitiv.rds")
dat %>% arrange(desc(mkt_cap_share)) %>% slice(1:25) %>% {sum(.$rev_share)}
skimr::skim(dat)
docxtractr::read_docx(path = "spilloverV3.docx")-> spill_doc
spill_doc %>% docxtractr::docx_extract_all_tbls()
head(dat,n = 21) 
dat %>% 
  ggplot(aes(area=RevenueFY0,fill=RIC)) + geom_treemap()
library(treemapify)

