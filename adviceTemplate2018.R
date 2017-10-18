rm(list = ls())

#############################
# Install and load packages #
#############################

# devtools::install_github('davidgohel/ReporteRsjars', args = "--no-multiarch")
# devtools::install_github("davidgohel/rvg", args = "--no-multiarch")
# devtools::install_github("davidgohel/ReporteRs@master", args = "--no-multiarch")

library(docxtractr)
library(officer)
library(icesSD)
library(icesSAG)
library(dplyr)
library(tidyr)
library(flextable)

##################
# Load functions #
##################

# Function to format a caption label
addCaption <- function(type = c("Figure", "Table"), number, text) {
  
  base_text_prop <- textProperties(font.family = "Calibri")
  fig_bold_text_prop <- chprop(base_text_prop,
                               font.size = 9,
                               font.weight = "bold",
                               text.align = "justify")
  fig_base_text_prop <- chprop(base_text_prop,
                               font.size = 9,
                               # font.style = "italic",
                               text.align = "justify")
  
  fig_pot <- pot(paste0(type, " ", stockDat$SECTION.NUMBER, ".", number, "\t"), 
                 format = fig_bold_text_prop) + 
    pot(paste0(stockDat$CAPTION.NAME, ". ", text), 
        format = fig_base_text_prop)
  return(fig_pot)
}

addHeading <- function(text) {
  
  base_text_prop <- textProperties(font.family = "Calibri")
  head_bold_text_prop <- chprop(base_text_prop,
                               font.size = 11,
                               font.weight = "bold",
                               text.align = "justify")
  head_italic_text_prop <- chprop(base_text_prop,
                               font.size = 11,
                               font.style = "italic",
                               text.align = "justify")
  
  fig_pot <- pot(paste0(stockDat$SECTION.NUMBER, ".", number, "\t"), 
                 format = fig_bold_text_prop) + 
    pot(paste0(stockDat$CAPTION.NAME, ". ", text), 
        format = fig_base_text_prop)
  return(fig_pot)
}

#############
# Load data #
#############

#Only stocks which last advice was in 2016, so the advice sheet is not updated

sharePoint <- "//community.ices.dk@SSL/DavWWWRoot/admin/AdvisoryProgramme/Personal folders/Scott/Advice_Template_Testing/"

slbi <- fromJSON("http://sd.ices.dk/services/odata3/StockListDWs3")$value %>% 
  filter(ActiveYear < 2017,
         YearOfNextAssessment == 2018)

#########################################################
# Create Draft given the Stock List Database and tables #
#########################################################
#
#Define some formats #this is not working as in officer there is no chprop function
base_text_prop <- textProperties(font.family = "Calibri",
                                 font.size = 10)
base_par_prop <- parProperties(text.align = "left")

header_text_prop <- chprop(base_text_prop,
                           font.size = 9,
                           font.style = "italic")

heading_text_prop <- chprop(base_text_prop,
                            font.size = 11,
                            font.weight = "bold")

heading_par_prop <- chprop(base_par_prop,
                           text.align = "justify")

fig_bold_text_prop <- chprop(base_text_prop,
                             font.size = 9,
                             font.weight = "bold")

fig_bold_par_prop <- chprop(base_par_prop,
                            text.align = "justify")

fig_base_text_prop <- chprop(base_text_prop,
                             font.size = 9)

fig_base_par_prop <- chprop(base_par_prop,
                            text.align = "justify")

head_bold_text_prop <- chprop(base_text_prop,
                              font.size = 11,
                              font.weight = "bold")

head_bold_par_prop <- chprop(base_par_prop,
                             text.align = "justify")

head_italic_text_prop <- chprop(base_text_prop,
                                font.size = 11,
                                font.style = "italic")

head_italic_par_prop <- chprop(base_par_prop,
                               text.align = "justify")
# stock.code = stock.code[1]
createDraft <- function(stock.code, file_path = NULL) {
  pub.date <- fileList$PubDate[fileList$StockCode %in% stock.code]
  ecoregion.name <- fileList$AdviceHeading[fileList$StockCode %in% stock.code]
  draft.url <- fileList$DraftURL[fileList$StockCode %in% stock.code]
  stock.name <- fileList$Description[fileList$StockCode %in% stock.code]
  caption.name <-  fileList$CaptionName[fileList$StockCode %in% stock.code]
  common.name <- fileList$SpeciesCommonName[fileList$StockCode %in% stock.code]
  data.category <- fileList$DataCategory[fileList$StockCode %in% stock.code]
  url.name <- fileList$URL[fileList$StockCode %in% stock.code]
  expert.url <- fileList$ExpertURL[fileList$StockCode %in% stock.code]
  expert.name <- fileList$ExpertGroup[fileList$StockCode %in% stock.code]
  adg.name <- fileList$AdviceDraftingGroup[fileList$StockCode %in% stock.code]
  
  
  if(is.null(file_path)){
    file_path <- paste0(draft.url, stock.code, ".docx")
  }
  
  if(file.exists(file_path)) {
    stop("WATCH OUT! You are trying to over-write a file!!!")
  }
  

####################################
## Start function to write sheets ##
####################################

# adviceTemplate <- function(stockID = "had-346a", output = "~/git/ices-dk/AdviceTemplate/tester.docx") {
  
# Setup data


## TO DO ## Also filter out the stocks for advice in 2018



# Category 3
cat3 <- sl[sl$DataCategory == 3 &
             complete.cases(sl),]

# Category 4
cat4 <- sl[sl$DataCategory == 4 &
             complete.cases(sl),]

# Category 5
cat5 <- sl[sl$DataCategory == 5 &
             complete.cases(sl),]

# Category 6
cat6 <- sl[sl$DataCategory == 6 &
             complete.cases(sl),]

#ADRIANA i don't think this is needed
#cat3$STOCK.CODE[1]
#stockID <- "gug-347d"

#stockDat <- cat3 %>%
#  filter(STOCK.CODE == stockID)
  # 
#  doc <- read_docx(paste0(sharePoint, stockID, ".docx"))

####################################################################
###the templates already exist, maybe would be easier to change dates 
#### and whatever needed in those files instead of doing them again.


  # Category 3 Catch Options Table
  catchOptionsData <- docx_extract_tbl(doc, tbl_number = 2, header = FALSE)
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

  # Category 3 advice Basis Table
  adviceBasisData <- docx_extract_tbl(doc, tbl_number = 3, header = FALSE)
  colnames(adviceBasisData) <- c("DESCRIPTION", "VALUE")
  
  adviceBasisTable <-  FlexTable(adviceBasisData, header.columns = FALSE,
                                     body.cell.props = cellProperties(padding.left = 2, padding.right = 2, padding.bottom = 0),
                                     body.text.props = textProperties(font.family = "Calibri", font.weight = "normal", font.size = 9)
  )
  adviceBasisTable[, 1] = cellProperties(background.color = "#E8EAEA")
  adviceBasisTable[, c(1:2)] = parProperties(text.align = "left")
  setFlexTableWidths(adviceBasisTable, c((3.89/2.54), (13.86/2.54)))
  
  # Category 3 assessment basis Table
  assessmentBasisData <- docx_extract_tbl(doc, tbl_number = 4, header = FALSE)
  colnames(assessmentBasisData) <- c("DESCRIPTION", "VALUE")
  
  assessmentBasisTable <-  FlexTable(assessmentBasisData, header.columns = FALSE,
                                 body.cell.props = cellProperties(padding.left = 2, padding.right = 2, padding.bottom = 0),
                                 body.text.props = textProperties(font.family = "Calibri", font.weight = "normal", font.size = 9)
  )
  assessmentBasisTable[, 1] = cellProperties(background.color = "#E8EAEA")
  assessmentBasisTable[, c(1:2)] = parProperties(text.align = "left")
  setFlexTableWidths(assessmentBasisTable, c((3.55/2.54), (14.45/2.54)))
  
  # Category 3 history of advice Table
  historyOfAdviceData <- docx_extract_tbl(doc, tbl_number = 5, header = TRUE)
  #
  temprow <- data.frame(matrix(c(rep.int(NA,length(historyOfAdviceData))),
                    nrow = 1,
                    ncol = length(historyOfAdviceData)))

  colnames(temprow) <- colnames(historyOfAdviceData)
  #change?
  temprow$Year <- 2018
  historyOfAdviceData <- rbind(historyOfAdviceData,
                               temprow)
  #
  historyOfAdviceTable <-  FlexTable(historyOfAdviceData , header.columns = TRUE,
                                     body.cell.props = cellProperties(padding.left = 2, padding.right = 2, padding.bottom = 0),
                                     body.par.props =  parProperties(text.align = "right", padding = 1),
                                     body.text.props = textProperties(font.family = "Calibri", font.weight = "normal", font.size = 9),
                                     header.cell.props = cellProperties(background.color = "#E8EAEA"),
                                     header.par.props =  parProperties(text.align = "center", padding = 1),
                                     header.text.props = textProperties(font.family = "Calibri", font.weight = "normal", font.size = 9)
  )

  historyOfAdviceTable[, "Year", to = "body"] = parProperties(text.align = "center", padding = 1)
  historyOfAdviceTable[, "ICES advice", to = "body"] = parProperties(text.align = "left", padding = 1)
  setFlexTableWidths(historyOfAdviceTable, rep((17.91/2.54) / ncol(historyOfAdviceData),
                                               ncol(historyOfAdviceData)))
  
  
  # # Category 1 Catch Options Table
  # table2 <- docx_extract_tbl(doc, tbl_number = 2) %>%
  
  #   # mutate()
  #   FlexTable(header.cell.props = cellProperties(background.color = "#E8EAEA"),
  #             header.par.props =  parProperties(text.align = "center", padding = 0),
  #             header.text.props = textProperties(font.family = "Calibri", font.weight = "normal", font.size = 9),
  #             body.text.props = textProperties(font.family = "Calibri", font.weight = "normal", font.size = 9),
  #             body.cell.props = cellProperties(padding.left = 1, padding.right = 5))
  # table2[,2 , to = "body"] = parProperties(text.align = "right")
  # table2[,3 , to = "body"] = parProperties(text.align = "center")
  # 
  
  # table2
  
  #Define some formats #
  base_text_prop <- textProperties(font.family = "Calibri")
  base_par_prop <- parProperties(text.align = "left")
  
  header_text_prop <- chprop(base_text_prop,
                             font.size = 9,
                             font.style = "italic")
  
  heading_text_prop <- chprop(base_text_prop,
                              font.size = 11,
                              font.weight = "bold")
  
  heading_par_prop <- chprop(base_par_prop,
                              text.align = "justify")
  
  headerPot <- pot(stockDat$SECTION.NUMBER,
                   format = heading_text_prop) + 
               "\t" +
               pot(stockDat$STOCK.NAME,
                   format = heading_text_prop)
  
  footerPot <- pot("ICES Advice 2017",
                    format = header_text_prop) +
    pot(", Book ",
        format = header_text_prop) + 
    pot(gsub("^.*\\.","", stockDat$SECTION.NUMBER),
        format = header_text_prop)
  
  headerPot2 <- pot("ICES Advice on fishing opportunities, catch, and effort\t\t",
                    format = header_text_prop) +
                pot("Published 30 June 2017",
                    format = header_text_prop)
    
  
##############################################################################  

  # draftDoc <- docx(template = paste0(sharePoint, "advice_template.docx"))
  draftDoc <- docx(template = "~/git/ices-dk/AdviceTemplate/advice_template_cat3.docx")
  
  # styles(draftDoc)
  draftDoc <- draftDoc %>%
    # Add ecoregion and release date, Is it a random date?
    addParagraph(value = pot(paste0(stockDat$ADVICE.HEADING, "\t\tPublished 30 June 2017"),
                             format = header_text_prop),
                 par.properties = base_par_prop,
                 # stylename = "AdviceHeader",
                 bookmark = "ECOREGION_HEADER") %>%
    # Add Section number and full stock name
    addParagraph(value = headerPot,
      # value = pot(paste0(stockDat$SECTION.NUMBER, "\t", stockDat$STOCK.NAME),
      #                        format = heading_text_prop),
                 par.properties = heading_par_prop,
                 # stylename = "AdviceHeading",
                 bookmark = "ADVICE_HEADING") %>%
    # Add Section number and full stock name
    addParagraph(value = addCaption(type = "Figure",
                                    number = 1, 
                                    text = "Summary of stock assessment."), 
                 bookmark = "Figure_1") %>%
    # Add footer
    addParagraph(value = footerPot,
                 par.properties = heading_par_prop,
                 bookmark = "BOOK_FOOTER") %>%
    # Add second page header
    addParagraph(value = headerPot2,
                 par.properties = heading_par_prop,
                 bookmark = "SECOND_HEADER") %>%
    addParagraph(value = footerPot,
                 par.properties = heading_par_prop,
                 bookmark = "BOOK_FOOTER2") %>%
    # Add Section number and full stock name
    addParagraph(value = addCaption(type = "Table",
                                    number = 1,
                                    text = "State of the stock and fishery, relative to reference points."), 
                 bookmark = "TABLE_1") %>%
    # Add Section number and full stock name
    addParagraph(value = addCaption(type = "Table",
                                    number = 2,
                                    text = "For stocks in ICES data categories 3–6, one catch option is provided."), 
                 bookmark = "TABLE_2") %>%
    addFlexTable(flextable = catchOptionsTable,
                 bookmark = "CatchOptions") %>%
    deleteBookmark(bookmark = "CatchOptions") %>%
    # Add Section number and full stock name
    addParagraph(value = addCaption(type = "Table",
                                    number = 3,
                                    text = "The basis of the advice."), 
                 bookmark = "TABLE_3") %>%
    addFlexTable(flextable = adviceBasisTable,
                 bookmark = "AdviceBasis") %>%
    deleteBookmark(bookmark = "AdviceBasis") %>%
    # Add Section number and full stock name
    addParagraph(value = addCaption(type = "Table",
                                    number = 4,
                                    text = "The basis of the assessment."), 
                 bookmark = "TABLE_4") %>%
    addFlexTable(flextable = assessmentBasisTable,
                 bookmark = "AssessmentBasis") %>%
    deleteBookmark(bookmark = "AssessmentBasis") %>%
    # Add Section number and full stock name
    addParagraph(value = addCaption(type = "Table",
                                    number = 5,
                                    text = "History of ICES advice, the agreed TAC, and ICES estimates of landings. All weights in thousand tonnes."),
                 bookmark = "TABLE_5") %>%
    addFlexTable(flextable = historyOfAdviceTable,
                 bookmark = "AdviceHistory") %>%
    deleteBookmark(bookmark = "AdviceHistory") %>%
    # Add Section number and full stock name
    addParagraph(value = addCaption(type = "Table",
                                    number = 6,
                                    text = "Catch distribution by fleet in 2016 as estimated by ICES."), 
                 bookmark = "TABLE_6") %>%
    # Add Section number and full stock name
    addParagraph(value = addCaption(type = "Table",
                                    number = 7,
                                    text = "History of commercial landings; official values presented by area for each country participating in the fishery."), 
                 bookmark = "TABLE_7") %>%
    # Add Section number and full stock name
    addParagraph(value = addCaption(type = "Table",
                                    number = 7,
                                    text = "Assessment summary (weights in tonnes)."), 
                 bookmark = "TABLE_8") %>%
    
    writeDoc(file = "~/git/ices-dk/AdviceTemplate/tester_3.docx")

  
  
    
 #From sharepointwrangling: 
  #########################################################
  # Create Draft given the Stock List Database and tables #
  #########################################################
  #
  #Define some formats #
  base_text_prop <- textProperties(font.family = "Calibri",
                                   font.size = 10)
  base_par_prop <- parProperties(text.align = "left")
  
  header_text_prop <- chprop(base_text_prop,
                             font.size = 9,
                             font.style = "italic")
  
  heading_text_prop <- chprop(base_text_prop,
                              font.size = 11,
                              font.weight = "bold")
  
  heading_par_prop <- chprop(base_par_prop,
                             text.align = "justify")
  
  fig_bold_text_prop <- chprop(base_text_prop,
                               font.size = 9,
                               font.weight = "bold")
  
  fig_bold_par_prop <- chprop(base_par_prop,
                              text.align = "justify")
  
  fig_base_text_prop <- chprop(base_text_prop,
                               font.size = 9)
  
  fig_base_par_prop <- chprop(base_par_prop,
                              text.align = "justify")
  
  head_bold_text_prop <- chprop(base_text_prop,
                                font.size = 11,
                                font.weight = "bold")
  
  head_bold_par_prop <- chprop(base_par_prop,
                               text.align = "justify")
  
  head_italic_text_prop <- chprop(base_text_prop,
                                  font.size = 11,
                                  font.style = "italic")
  
  head_italic_par_prop <- chprop(base_par_prop,
                                 text.align = "justify")
  
  