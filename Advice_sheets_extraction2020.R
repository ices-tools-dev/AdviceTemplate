# Loop extraction as much as possible
#

library(purrr)

getwd()

# setwd("D:/Adriana/Documents/R_Projects/AdviceTemplate")

# I will use the modified file list as starting point, it has the dates and dois

x2020_Advice_Overview <- read.csv("2020_Advice_Overview.csv")

#Where you want to save the files
setwd("D:/Adriana/Documents/AdviceSheets2020")

# fileList <- x2020_Advice_Overview

fileList <- get_filelist(2020)

# cat1 <- X2020_Advice_Overview %>% filter(DataCategory == 1)
List <- split(x2020_Advice_Overview,x2020_Advice_Overview$StockKeyLabel)

# List <- split(fileList,fileList$StockKeyLabel)

# Set n from 1 to the number of elements in List

for(n in 1:184){
  dat <- List[[n]]
  tryCatch({source('D:/Adriana/Documents/R_Projects/AdviceTemplate/single_extract2020.R')},error=function(e){}) 
}

list.files()

# 45 out of 184

ok <- list.files() 
ok <- unlist(strsplit(ok, split='.docx', fixed=TRUE))

library(operators)
fileList2 <- dplyr::filter(x2020_Advice_Overview, StockKeyLabel %!in% ok)
detach("package:operators", unload=TRUE)
#Here

fileList2 <- fileList2 %>% drop_na(filepath)

List2 <- split(fileList2,fileList2$StockKeyLabel, drop = TRUE)

for(n in 1:128){
  dat <- List2[[n]]
  tryCatch({source('D:/Adriana/Documents/R_Projects/AdviceTemplate/single_extract2020.R')},error=function(e){}) 
}





n = 1
dat <- List2[[n]]
n = n+1
dat <- List2[[n]]
source('D:/Adriana/Documents/R_Projects/AdviceTemplate/single_extract2020.R')

list.files("D:/Adriana/Documents/AdviceSheets2020/ready")
x <-try(blast(n))
x
