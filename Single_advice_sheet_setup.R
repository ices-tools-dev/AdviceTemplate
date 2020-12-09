fileList <- get_filelist(2019)

#Write the name of the WG

WG<- fileList %>% filter(ExpertGroup=="WGCSE")

#for nephrops of WGCSE

# WG <- WG %>% filter(AdviceDraftingGroup == "ADGNEPH")


#Check that the number of elements in this list corresponds to those in the excel
# Change the number form 1 to n number of names in the WG list

rm(fileName)
rm(stock_sd)


stock_name <- WG$StockKeyLabel[2]

# stock_name  <- "ane.27.8"

# Check the correct doi in the excel

doi <- "10.17895/ices.advice.4825"

fulldoi<- paste0("https://doi.org/",doi)

dat <- cat1.l[[1]]
source('D:/Adriana/Documents/R_Projects/AdviceTemplate/single_extract2020.R')
dat <- cat1.l[[2]]
source('D:/Adriana/Documents/R_Projects/AdviceTemplate/single_extract2020.R')
dat <- cat1.l[[10]]
source('D:/Adriana/Documents/R_Projects/AdviceTemplate/single_extract2020.R')
dat <- cat1.l[[1]]
source('D:/Adriana/Documents/R_Projects/AdviceTemplate/single_extract2020.R')


rm(list = ls.str(mode = c('character')))
rm(tab_heads, tab_names, table_1, table_cells, table_header_name, tabs)
dat <- cat1.l[[2]]
thewholething()
rm(list = ls.str(mode = 'character'))
rm(tab_heads, tab_names, table_1, table_cells, table_header_name, tabs, fig_heads)
dat <- cat1.l[[3]]
thewholething()
rm(list = ls.str(mode = 'character'))	
rm(tab_heads, tab_names, table_1, table_cells, table_header_name, tabs, fig_heads)
dat <- cat1.l[[4]]
thewholething()
rm(list = ls.str(mode = 'character'))
rm(tab_heads, tab_names, table_1, table_cells, table_header_name, tabs, fig_heads)
dat <- cat1.l[[5]]
thewholething()
rm(list = ls.str(mode = 'character'))
n=6
thewholething()
rm(list = ls.str(mode = 'character'))	
n=7
thewholething()
rm(list = ls.str(mode = 'character'))	
n=8
thewholething()
rm(list = ls.str(mode = 'character'))
n=9
thewholething()
rm(list = ls.str(mode = 'character'))
n=1
thewholething()
rm(list = ls.str(mode = 'character'))	
n=10
thewholething()
rm(list = ls.str(mode = 'character'))	
n=1
thewholething()
rm(list = ls.str(mode = 'character'))	
n=11
thewholething()
rm(list = ls.str(mode = 'character'))	
thewholething()
rm(list = ls.str(mode = 'character'))	
n=12
thewholething()
rm(list = ls.str(mode = 'character'))	
n=13
thewholething()
rm(list = ls.str(mode = 'character'))	
n=14
thewholething()
rm(list = ls.str(mode = 'character'))	
n=15
thewholething()
rm(list = ls.str(mode = 'character'))	
n=16
thewholething()
rm(list = ls.str(mode = 'character'))	
n=17
thewholething()
rm(list = ls.str(mode = 'character'))	
thewholething()
n=18
rm(list = ls.str(mode = 'character'))	
thewholething()
n=19
rm(list = ls.str(mode = 'character'))	
thewholething()
n=20
rm(list = ls.str(mode = 'character'))	
thewholething()
n=21
rm(list = ls.str(mode = 'character'))	
thewholething()
n=22
rm(list = ls.str(mode = 'character'))	
thewholething()
n=23
rm(list = ls.str(mode = 'character'))	
thewholething()
n=24
rm(list = ls.str(mode = 'character'))	
thewholething()
n=25
rm(list = ls.str(mode = 'character'))	
thewholething()
n=26
rm(list = ls.str(mode = 'character'))	
thewholething()
n=27
rm(list = ls.str(mode = 'character'))	
thewholething()
n=28
rm(list = ls.str(mode = 'character'))	
thewholething()
n=29
rm(list = ls.str(mode = 'character'))	
thewholething()
n=30
rm(list = ls.str(mode = 'character'))	
thewholething()

n=23	
thewholething()
n=24	
thewholething()
n=25	
thewholething()
n=26	
thewholething()
n=27	
thewholething()
n=28	
thewholething()
n=29	
thewholething()
n=30	
thewholething()
n=31	
thewholething()
n=32	
thewholething()
n=33	
thewholething()
n=34	
thewholething()
n=35	
thewholething()
n=36	
thewholething()
n=37	
thewholething()
n=38	
thewholething()
n=39	
thewholething()
n=40	
thewholething()
n=41	
thewholething()
n=42	
thewholething()
n=43	
thewholething()
n=44	
thewholething()
n=45	
thewholething()
n=46	
thewholething()
n=47	
thewholething()
n=48	
thewholething()
n=49	
thewholething()
n=50	
thewholething()
n=51	
thewholething()
n=52	
thewholething()
n=53	
thewholething()
n=54	
thewholething()
n=55	
thewholething()
n=56	
thewholething()
n=57	
thewholething()
n=58	
thewholething()
n=59	
thewholething()
n=60	
thewholething()
n=61	
thewholething()
n=62	
thewholething()
n=63	
thewholething()
n=64	
thewholething()
n=65	
thewholething()
n=66	
thewholething()
n=67	
thewholething()
n=68	
thewholething()
n=69	
thewholething()
n=70	
thewholething()
n=71	
thewholething()
n=72	
thewholething()
n=73	
thewholething()
n=74	
thewholething()
n=75	
thewholething()
n=76	
thewholething()
n=77	
thewholething()
n=78	
thewholething()
n=79	
thewholething()
n=80	
thewholething()
n=81	
thewholething()
n=82	
thewholething()
n=83	
thewholething()
n=84	thewholething()
n=85	thewholething()
n=86	thewholething()
n=87	thewholething()
n=88	thewholething()
n=89	thewholething()
n=90	thewholething()
n=91	thewholething()
n=92	thewholething()
n=93	thewholething()
n=94	thewholething()
n=95	thewholething()
n=96	thewholething()



