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


#bianual advice to be done in 2018, advice sheets need to be updated from scratch

rawsd <- fromJSON("http://sd.ices.dk/services/odata3/StockListDWs3")$value %>% 
  filter(ActiveYear < 2017,
          YearOfNextAssessment == 2018)

#anual advice to be done in 2018, advice sheets should be easier to fit?

rawsd <- fromJSON("http://sd.ices.dk/services/odata3/StockListDWs3")$value %>% 
  filter(ActiveYear == 2017,
         YearOfNextAssessment == 2018,
        YearOfLastAssessment == 2017)

# Load Table functions #

source("table_functions.R")


#load a word file of 2017 advice

doc <- docxtractr::read_docx("//community.ices.dk/DavWWWRoot/Advice/Advice2017/NorthSea/Released_Advice/cod.27.47d20.docx")



#some exploration, I will call it docx for officer and doc for docxtractr

content <- docx_summary(docx)
content

docx_bookmarks(docx)

#number of Tables present

docxtractr::docx_tbl_count(doc)

#does not work, as tables are unequal
#tbls <- data.frame(docx_extract_all_tbls(doc, guess_header = TRUE, trim = TRUE))

#tbls

#Extract one at the time and convert to flextable(the only
#way I found)

tbl1 <- flextable(data.frame(docx_extract_tbl(doc, tbl_number = 1,header = TRUE, trim = TRUE)))
tbl2 <- flextable(data.frame(docx_extract_tbl(doc, tbl_number = 2,header = TRUE, trim = TRUE)))
# tbl3 <- flextable(data.frame(docx_extract_tbl(doc, tbl_number = 3,header = TRUE, trim = TRUE)))
tbl4 <- flextable(data.frame(docx_extract_tbl(doc, tbl_number = 4,header = TRUE, trim = TRUE)))
tbl5 <- flextable(data.frame(docx_extract_tbl(doc, tbl_number = 5,header = TRUE, trim = TRUE)))
#tbl6 <- flextable(data.frame(docx_extract_tbl(doc, tbl_number = 6,header = TRUE, trim = TRUE)))
tbl7 <- flextable(data.frame(docx_extract_tbl(doc, tbl_number = 7,header = TRUE, trim = TRUE)))
tbl8 <- flextable(data.frame(docx_extract_tbl(doc, tbl_number = 8,header = TRUE, trim = TRUE)))
tbl9 <- flextable(data.frame(docx_extract_tbl(doc, tbl_number = 9,header = TRUE, trim = TRUE)))
tbl10 <- flextable(data.frame(docx_extract_tbl(doc, tbl_number = 10,header = TRUE, trim = TRUE)))
tbl11 <- flextable(data.frame(docx_extract_tbl(doc, tbl_number = 11,header = TRUE, trim = TRUE)))
tbl12 <- flextable(data.frame(docx_extract_tbl(doc, tbl_number = 12,header = TRUE, trim = TRUE)))
tbl13 <- flextable(data.frame(docx_extract_tbl(doc, tbl_number = 13,header = TRUE, trim = TRUE)))

####Formatting work for each table, 

tbl1 #"The basis for the catch options."
catchOptionsTable<-void(tbl1, j = 2)
catchOptionsTable

#coloured header

catchOptionsTable <- catchOptionsTable %>%
                                bg(bg= "#E8EAEA", part= "header")

font<- fp_text(color = "black", font.size = 9,font.family = "Calibri", bold=FALSE)

style(catchOptionsTable, i = NULL, j = NULL, pr_t = font, pr_p = NULL, pr_c = NULL,
      part = "all")

align(catchOptionsTable, i = 1, j = NULL, align = "center", part = "header")
align(catchOptionsTable, i = NULL, j = c(1,4), align = "left", part = "body")
align(catchOptionsTable, i = NULL, j = 2, align = "right", part = "body")
align(catchOptionsTable, i = NULL, j = 3, align = "center", part = "body")

#autofit and the fix a bit

autofit(catchOptionsTable) #bohh...
autofit(catchOptionsTable, add_w = 2.6, add_h = -2)

#rebooohh, this way should be finer, but I really can't see it

catchOptionsTable <- height(catchOptionsTable, height= 2.54)
catchOptionsTable <- width(catchOptionsTable, j =c(1,3) , width = 6.42) 
catchOptionsTable <- width(catchOptionsTable, j =2 , width = 5.07) 


##I dont see much differences in the viewer, hope is fine

#padding think is not needed, messes uo the size 
#  padding(catchOptionsTable, i = NULL, j = NULL, 
        #padding.bottom = 0, padding.left = 2, padding.right =2,
        #part = "body")




tbl2   ###"Annual catch options. All weights are in tonnes."
adviceBasisTable<-void(tbl2, j = 2:10)

bg(adviceBasisTable, i= 1, bg= "#E8EAEA", part= "all")
bg(adviceBasisTable, i= 3, bg= "#E8EAEA", part= "body")
bg(adviceBasisTable, i= 22, bg= "#E8EAEA", part= "body")
 

font<- fp_text(color = "black", font.size = 9,font.family = "Calibri", bold=FALSE)

style(adviceBasisTable, i = NULL, j = NULL, pr_t = font, pr_p = NULL, pr_c = NULL,
      part = "all")

#not working, cant modify the dots, should be easier when importing the table
#adviceBasisTable <- gsub("\\.", " ", adviceBasisTable)

align(adviceBasisTable, i = NULL, j = 1, align = "left", part = "all")
align(adviceBasisTable, i = NULL, j = -1, align = "right", part = "body")
align(adviceBasisTable, i = 1, j = -1, align = "center", part = "header") 
#can't see much changing in the viewer!!

#flextable does not allow to merge body and header cells!!

adviceBasisTable <- merge_h(adviceBasisTable, i = c(1,3,22),  part="body")

#this works but I don't like it

adviceBasisTable <- border(adviceBasisTable, i = c(1,3), j=1,border.right = fp_border(color="#E8EAEA"))


#autofit and the fix a bit, not happy with this autofit neither

autofit(adviceBasisTable)

#can't see if its working, I'm using measures from advicetemplate.R 

adviceBasisTable <- height(adviceBasisTable, height= 2.54)
adviceBasisTable <- width(adviceBasisTable, j =c(2:10) , width = 3.89) 
adviceBasisTable <- width(adviceBasisTable, j =1 , width = 13.86) 

############################
tbl4                 ##"and the FMSY range was updated as follows:"         
tbl4<-void(tbl4, j = 2)
tbl4

#coloured header

tbl4 <- tbl4 %>%
  bg(bg= "#E8EAEA", part= "header")

font<- fp_text(color = "black", font.size = 9,font.family = "Calibri", bold=FALSE)

style(tbl4, i = NULL, j = NULL, pr_t = font, pr_p = NULL, pr_c = NULL,
      part = "all")

align(tbl4, i = 1, j = NULL, align = "center", part = "header")
align(tbl4, j = 1, align = "left", part = "body")
align(tbl4, i = NULL, j = 2, align = "right", part = "body")
align(tbl4, i = NULL, j = 3, align = "center", part = "body")

#pending of clarifying the dimensions and units

#autofit(tbl4) #bohh...
#autofit(catchOptionsTable, add_w = 2.6, add_h = -2)

#####################
tbl5 ##"Reference points, values, and their technical basis."

tbl5<-void(tbl5, j = 3)
tbl5

#coloured header
tbl5 <- tbl5 %>%
  bg(bg= "#E8EAEA", part= "header")

font<- fp_text(color = "black", font.size = 9,font.family = "Calibri", bold=FALSE)

style(tbl5, i = NULL, j = NULL, pr_t = font, pr_p = NULL, pr_c = NULL,
      part = "all")

tbl5 <- merge_at(tbl5, i= 1:2, j = 1 )
tbl5 <- merge_at(tbl5, i= 3:6, j = 1 )
tbl5 <- merge_at(tbl5, i= 7:10, j = 1 ) #if exists, I have to add something like that
tbl5 <- merge_at(tbl5, i= 7:10, j = 5 )

#first row has no source, is it missing merging or is it ok?

align(tbl5, i = 1, j = NULL, align = "center", part = "header")
align(tbl5, j = 1:2, align = "left", part = "body")
align(tbl5, i = NULL, j = 3, align = "right", part = "body")
align(tbl5, i = NULL, j = 4, align = "left", part = "body")
align(tbl5, i = NULL, j = 5, align = "center", part = "body")

### tbl7, tbl8 and tbl9 are identical, is it corrected with the initial grepl in the new_advice_prep.R?
#"are presented to the nearest thousand tonnes"

histTable <- tbl7

#add row



df <- data.frame(histTable, stringsAsFactors = FALSE)





