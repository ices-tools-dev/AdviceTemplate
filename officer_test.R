rm(list = ls())
#############################
# Install and load packages #
#############################

library(docxtractr)
library(officer)
library(jsonlite)
library(stringr)
library(dplyr)

# Note: You must log in to SharePoint and have this drive mapped
sharePoint <- "//community.ices.dk/DavWWWRoot/"
if(!dir.exists(paste0(sharePoint, "Advice/Advice2017/"))) {
  stop("Note: You must be on the ICES network and log in to SharePoint and have this drive mapped")
}

rawsd <- fromJSON("http://sd.ices.dk/services/odata3/StockListDWs3")$value %>% 
  filter(ActiveYear == 2017,
         YearOfNextAssessment == 2018)

advice_sheet <-  sprintf("//community.ices.dk/DavWWWRoot/Advice/Advice2017/NorthSea/Released_Advice/%s.docx",
                         grep("cod.27.47d20", rawsd$StockKeyLabel, value = TRUE))

doc <- read_docx(advice_sheet)

content <- docx_summary(doc)
