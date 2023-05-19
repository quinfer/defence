
#|label: setup
#|include: false
library(tidyverse);library(lubridate);library(kableExtra);library(corrr)
raw <- readRDS("data/raw_dat.rds")
spill_vol <- readRDS("data/spillover_vols.rds")
spill_rtn <- readRDS("data/spillover_rtns.rds")
static <- readRDS("data/static.rds")
width=9
height=6
spill_rtn$to50 %>% {unique(.$Stock)}->stock_names
stock_names<-str_remove_all(stock_names,"_") 
raw$prices[-1,] %>% 
  map_df(as.numeric) %>%
  mutate(Date=as.Date(Name,origin = "1899-12-30")) %>%
  select(-Name) %>%
  pivot_longer(!Date,values_to = "price",names_to = "stock") %>%
  filter(Date>=dmy("23-08-2010")) %>%
  mutate(stock_match_name=str_remove_all(stock,"\\'|\\(|\\)|\\.|-"),
         stock_match_name=str_remove_all(stock_match_name," "),
         stock=str_to_title(stock)) %>%
  filter(stock_match_name %in% stock_names) %>%
  arrange(stock,Date) %>%
  group_by(stock) %>%
  mutate(return=log(price/lag(price)),
         volatility=return^2) -> dat

country_of_origin<-tibble(stock=unique(dat$stock),country=
                            c("China","France","China","China","UK","US","France","US","US","US","US","US","Germany","US","US","UK","France","Singapore","US","France","US"))
dat %>% left_join(country_of_origin,by="stock")->dat



#| label: fig-prices
#| fig-cap: "Price series levels"
dat %>%
  ggplot(aes(x=Date,y=price,color=country)) + geom_line() + 
  facet_wrap(~stock,scales = "free_y") +
  theme(text = element_text(size=7),
        legend.position = "top",
        legend.title = element_blank(),
        axis.text.x = element_text(angle = 45),
        legend.text = element_text(size = 12)) +
  guides(color = guide_legend(nrow = 1)) +
  labs(x="")
ggsave("plots/fig-prices.png",width = 9,height = 12)

#| label: fig-rtns
#| fig-cap: "Daily returns"
dat %>%
  ggplot(aes(x=Date,y=return,color=country)) + geom_line() + 
  facet_wrap(~stock) +
  scale_color_brewer(type='qual',palette = "Set2") +
  theme(text = element_text(size=7),
        legend.position = "top",
        legend.title = element_blank(),
        axis.text.x = element_text(angle = 45),
        legend.text = element_text(size = 12)) +
  guides(color = guide_legend(nrow = 1)) +
  labs(x="")
ggsave("plots/fig-rtns.png",width = 9,height = 12)


#| label: fig-vols
#| fig-cap: "Daily Volatilities"
dat %>%
  ggplot(aes(x=Date,y=volatility,color=country)) + geom_line() + 
  facet_wrap(~stock) +
  scale_color_brewer(type='qual',palette = "Set2") +
  theme(text = element_text(size=7),
        legend.position = "top",
        legend.title = element_blank(),
        axis.text.x = element_text(angle = 45),
        legend.text = element_text(size = 12)) +
  guides(color = guide_legend(nrow = 1)) +
  labs(x="")
ggsave("plots/fig-vol.png",width = 9,height = 12)


#| label: tbl-sumrtn
#| tbl-cap: "Summary statistics of daily returns"
docxtractr::read_docx(path = "spilloverV3.docx")-> spill_doc
spill_doc |>
  docxtractr::docx_extract_all_tbls() %>%
  {.[[3]]} -> tblsumrtns
spill_doc |>
  docxtractr::docx_extract_all_tbls() %>%
  {.[[4]]} -> tblsumvol
names(tblsumrtns)[1]<-"A&D Stock"
tblsumrtns |> kbl(format = "latex",booktabs = TRUE) |> kable_styling(latex_options = c("scale_down","HOLD_position"))


#| label: tbl-sumvol
#| tbl-cap: "Summary statistics of daily volatilies"
names(tblsumvol)[1]<-"A&D Stock"
tblsumvol |> kbl(format = "latex",booktabs = TRUE) |> kable_styling(latex_options = c("scale_down","HOLD_position"))


#| label: tbl-Xtremes
#| tbl-cap: "Extremely votality events"
dat %>% filter(volatility>0.06) %>% select(!stock_match_name&!price) %>% kbl(digits = 2,format = "latex",booktabs = TRUE) |>
  kable_styling(latex_options="HOLD_position")


#| label: fig-cor
#| fig-cap: "Correlation matrix of daily returns"
dat %>% 
  select(stock,Date,return)%>% 
  drop_na() %>%
  pivot_wider(names_from = "stock",values_from = "return") %>%
  select(!Date) %>%
  correlate() %>% 
  rplot(print_cor = TRUE) +
  theme(axis.text.x = element_text(angle = 45,hjust = 1))
ggsave("plots/fig-cor.png",width = 9,height = 12)



#| label: fig-rtn50
#| fig-cap: "Network topology of static results for returns at the 50th percentile"
knitr::include_graphics("plots/fig-rtn50.png",dpi = 400)


#| label: fig-rtn95
#| fig-cap: "Network topology of static results for returns at the 95th percentile"
knitr::include_graphics("plots/fig-rtn95.png",dpi = 400)


#| label: fig-rtn5
#| fig-cap: "Network topology of static results for returns at the 5th percentile"
knitr::include_graphics("plots/fig-rtn5.png",dpi = 400)


#| label: fig-vol50
#| fig-cap: "Network topology of static results for volatility at the 50th percentile"
knitr::include_graphics("plots/fig-vol50.png",dpi=400)


#| label: fig-vol95
#| fig-cap: "Network topology of static results for volatility at the 95th percentile"
knitr::include_graphics("plots/fig-vol95.png",dpi=400)


#| label: fig-vol5
#| fig-cap: "Network topology of static results for volatility at the 5th percentile"
knitr::include_graphics("plots/fig-vol5.png",dpi=400)



list("Russia began annexation of Crimea"="February 20,2014",
     "Start of war in Donbas by pro-Russian activists"= "April 7, 2014",
     "October 2014 flash crash"="Oct 15,2014",
     "Brexit referendum"= "June 23, 2016", 
     "Federal Reserve raises interest rates"= "December 14, 2016", 
     "the United Kingdom invokes article 50 of the Lisbon Treaty"="March 29, 2017", 
     "snap election held in the United Kingdom"="June 8, 2017",
     "Dash for cash crisis in bond market peaks"="March 18, 2020",
     "Russia initiated a special military operation in Donbas"="Feb 24,2022")->imp_dates

labels_dat<-tibble(
  label=letters[1:length(imp_dates)],
xvalues=parse_date_time(imp_dates,"mdy"),
description=names(imp_dates))



#| label: tbl-dates
#| tbl-cap: "Important Dates"
labels_dat %>% rename(date=xvalues) %>% mutate(date=ymd(date)) %>% kbl(format = "latex",booktabs = TRUE) %>% kable_styling(latex_options = "HOLD_position")


#| label: tbl-reg1
#| tbl-cap: "Drivers of return spillovers across A&D companies for the full sample period"
spill_doc <- docxtractr::read_docx("spilloverV3.docx")
spill_doc %>% docxtractr::docx_extract_all_tbls() %>% {.[[27]]}->tbl_reg_1
spill_doc %>% docxtractr::docx_extract_all_tbls() %>% {.[[28]]}->tbl_reg_2
names<-tbl_reg_1[1,] %>% unlist(use.names = F)
library(kableExtra)
tbl_reg_1[-1,] |>
  kbl(col.names = names,format = "latex",booktabs = TRUE) %>%
  add_header_above(header = c(" ","Middle quantile"=2,"Upper quantile"=2,"Lower quantile"=2)) |> add_footnote("This table presents the estimated coefficients of the regression model in Equation (7) based on a covariance estimator that accounts for the presence of heteroscedasticityand autocorrelation (HAC). The sample period is 23 August 2010 –July 1, 2022.") |> kable_styling(latex_options = c("scale_down","HOLD_position"))


#| label: tbl-reg2
#| tbl-cap: "Drivers of volatility spillovers across A&D companies for the full sample period"
tbl_reg_2[-1,] |>
  kbl(col.names = names,format = "latex",booktabs = TRUE) %>%
  add_header_above(header = c(" ","Middle quantile"=2,"Upper quantile"=2,"Lower quantile"=2)) |> add_footnote("This table presents the estimated coefficients of the regression model in Equation (7) based on a covariance estimator that accounts for the presence of heteroscedasticityand autocorrelation (HAC). The sample period is 23 August 2010 –July 1, 2022.") |> kable_styling(latex_options = c("scale_down","HOLD_position"))


#| label: tbl-ref
#| tbl-cap: "Size information for our study's sample"
refinitiv<-readRDS("data/refinitiv.rds")
docxtractr::read_docx(path = "spilloverV3.docx")-> spill_doc
spill_doc %>% docxtractr::docx_extract_all_tbls() %>% {.[[29]]}->sample_dat
refinitiv %>% 
  filter(RIC %in% sample_dat$Identifier.RIC) %>%
  select(Name,MktCap,RevenueFY0,RevenueFY1) %>%
   kbl(format = "latex",col.names = c("Company Name","Market Capitalistation","Total Revenue 2022","Total Revenue 2021"),booktabs = TRUE,format.args = list(big.mark = ",")) |> 
  kable_styling(latex_options = c("scale_down","HOLD_position")) |>
  footnote("This data is source from Refinitiv Eikon and all values are in US Dollars.  The Market Capitalisation is a weight average of the 2022 daily values")


#| label: fig-mktcapshare
#| fig-cap: "Treemap of market capitalisation of our study's sample"
library(treemapify)
refinitiv %>% 
  mutate(Our_Study=if_else(RIC %in% sample_dat$Identifier.RIC,"This Study","Rest of A&D industry")) %>%
  ggplot(aes(area=MktCap,fill=Our_Study)) + geom_treemap()

