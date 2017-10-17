rm(list = ls())
#############################
# Install and load packages #
#############################

library(docxtractr)
library(officer)
library(magrittr)
library(jsonlite)
library(stringr)
library(dplyr)
library(flextable)

# Note: You must log in to SharePoint and have this drive mapped
sharePoint <- "//community.ices.dk/DavWWWRoot/"
if(!dir.exists(paste0(sharePoint, "Advice/Advice2017/"))) {
  stop("Note: You must be on the ICES network and log in to SharePoint and have this drive mapped")
}

#all advice to be done in 2018

rawsd <- fromJSON("http://sd.ices.dk/services/odata3/StockListDWs3")$value %>% 
  filter(ActiveYear == 2017,
         YearOfNextAssessment == 2018)


#bianual advice to be done in 2018

#rawsd <- fromJSON("http://sd.ices.dk/services/odata3/StockListDWs3")$value %>% 
 # filter(ActiveYear < 2017,
#         YearOfNextAssessment == 2018)

#anual advice to be done in 2018, advice sheets should be easier to fit?

#rawsd <- fromJSON("http://sd.ices.dk/services/odata3/StockListDWs3")$value %>% 
#  filter(ActiveYear == 2017,
#         YearOfNextAssessment == 2018,
#         YearOfLastAssessment == 2017)

# Load Table functions #

#source("table_functions.R")


#load a word file of 2017 advice

advice_sheet <-  sprintf("//community.ices.dk/DavWWWRoot/Advice/Advice2017/NorthSea/Released_Advice/%s.docx",
                         grep("cod.27.47d20", rawsd$StockKeyLabel, value = TRUE))


doc <- read_docx(advice_sheet)

#some exploration

content <- docx_summary(doc)
content

docx_bookmarks(doc)

#number of Tables present

docx_tbl_count(doc)

#does not work, as tables are unequal
#tbls <- data.frame(docx_extract_all_tbls(doc, guess_header = TRUE, trim = TRUE))

#tbls

#Extract one at the time and convert to flextable(the only
#way I found)

tbl1 <- flextable(data.frame(docx_extract_tbl(doc, tbl_number = 1,header = TRUE, trim = TRUE)))
tbl2 <- flextable(data.frame(docx_extract_tbl(doc, tbl_number = 2,header = TRUE, trim = TRUE)))
tbl3 <- flextable(data.frame(docx_extract_tbl(doc, tbl_number = 3,header = TRUE, trim = TRUE)))
tbl4 <- flextable(data.frame(docx_extract_tbl(doc, tbl_number = 4,header = TRUE, trim = TRUE)))
tbl5 <- flextable(data.frame(docx_extract_tbl(doc, tbl_number = 5,header = TRUE, trim = TRUE)))
tbl6 <- flextable(data.frame(docx_extract_tbl(doc, tbl_number = 6,header = TRUE, trim = TRUE)))
tbl7 <- flextable(data.frame(docx_extract_tbl(doc, tbl_number = 7,header = TRUE, trim = TRUE)))
tbl8 <- flextable(data.frame(docx_extract_tbl(doc, tbl_number = 8,header = TRUE, trim = TRUE)))
tbl9 <- flextable(data.frame(docx_extract_tbl(doc, tbl_number = 9,header = TRUE, trim = TRUE)))
tbl10 <- flextable(data.frame(docx_extract_tbl(doc, tbl_number = 10,header = TRUE, trim = TRUE)))
tbl11 <- flextable(data.frame(docx_extract_tbl(doc, tbl_number = 11,header = TRUE, trim = TRUE)))
tbl12 <- flextable(data.frame(docx_extract_tbl(doc, tbl_number = 12,header = TRUE, trim = TRUE)))
tbl13 <- flextable(data.frame(docx_extract_tbl(doc, tbl_number = 13,header = TRUE, trim = TRUE)))

tbl2
tbl2<-void(tbl2, j = 2:10)
tbl2



catchOptionsData <- docx_extract_tbl(doc, header = TRUE)

colnames(catchOptionsData) <- c("DESCRIPTION", "VALUE", "BLANK")
catchOptionsData <- apply(catchOptionsData, 2, function(x) as.character(gsub("[[:digit:]]", "X",  x)))

catchOptionsTable <- FlexTable(catchOptionsData, header.columns = FALSE,
                               body.cell.props = cellProperties(padding.left = 2, padding.right = 2, padding.bottom = 0),
                               body.text.props = textProperties(font.family = "Calibri", font.weight = "normal", font.size = 9)
)
catchOptionsTable[, 1] = cellProperties(background.color = "#E8EAEA")
catchOptionsTable[, 2] = parProperties(text.align = "right")

catchOptionsTable[catchOptionsData[,c(1)] %in% c("Precautionary buffer", "Uncertainty cap"), 2] = parProperties(text.align = "left")
catchOptionsTable[catchOptionsData[,c(1)] %in% c("Precautionary buffer", "Uncertainty cap"), 3] = parProperties(text.align = "right")

spanFlexTableColumns(catchOptionsTable, 
                     i = !catchOptionsData[,c(1)] %in% c("Precautionary buffer", "Uncertainty cap"),
                     from = 2,
                     to = 3)
setFlexTableWidths(catchOptionsTable, c((6.42/2.54), (5.07/2.54), (6.52/2.54)))

