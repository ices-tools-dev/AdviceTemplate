rm(list = ls())
#############################
# Install and load packages #
#############################
library(readxl) # not necessary once SLD is up and running
library(tidyverse)
library(docxtractr)
library(ReporteRs)

#############################################
# Organizing stock list and section numbers #
#############################################
# Note: will not be necessary when SLD is operational
sectionNum <- read_excel("~/git/ices-dk/AdviceTemplate/data/section numbers 2017.xlsx") %>%
  filter(`New advice in 2017` == "x") %>%
  select(STOCK.CODE = `Stock lable_2017`,
         OLD.CODE = `Stock code_pre 2017`,
         SECTION.NUMBER = `Section number 2016`) %>%
  mutate(OLD.CODE = tolower(OLD.CODE))

stockList <- read_excel("~/git/ices-dk/AdviceTemplate/data/Stock list for 2016.xlsx",
                          sheet = "Stock Overview")

stockList <- stockList[ , !duplicated(colnames(stockList))]
ecoregions <- colnames(stockList)[grepl("^.+(Ecoregion)$", colnames(stockList))]

dots <- lapply(c("New advice in 2017", "Stock lable_2017", "Stock code_pre 2017",
                 "Full name_2017", ecoregions, 
                 "ICES Data Category 2016", "Species Common Name", "Species Scientific Name"),
               as.symbol)

stockList <- stockList %>%
  select_(.dots = dots) %>%
  rename(ADVICE_2017 = `New advice in 2017`,
         STOCK.CODE = `Stock lable_2017`,
         OLD.CODE = `Stock code_pre 2017`,
         STOCK.NAME = `Full name_2017`,
         CATEGORY = `ICES Data Category 2016`,
         COMMON.NAME = `Species Common Name`,
         SPECIES.NAME = `Species Scientific Name`) %>%
  gather(ECOREGION, value, -ADVICE_2017, -STOCK.CODE, -OLD.CODE, -STOCK.NAME, -CATEGORY, -COMMON.NAME, -SPECIES.NAME) %>%
  filter(!is.na(value),
         ADVICE_2017 == "x") %>%
  select(-ADVICE_2017,
         -value) %>%
  mutate(STOCK.CODE = tolower(STOCK.CODE),
         OLD.CODE = tolower(OLD.CODE),
         CATEGORY = gsub("\\..*$", "", CATEGORY), # Just use the first digit
         ECOREGION = gsub(" Ecoregion", "", as.character(ECOREGION))) %>%
  group_by(STOCK.CODE) %>%
  mutate(count = n()) %>%
  spread(ECOREGION, ECOREGION) %>%
  unite(ADVICE.HEADING, 8:19, sep = ", ", remove = FALSE) %>%
  mutate(STOCK.NAME = gsub(" $","", STOCK.NAME, perl = TRUE), # remove trailing spaces
         ADVICE.HEADING = gsub("NA, ", "", ADVICE.HEADING), # remove NAs
         ADVICE.HEADING = gsub(", NA", "", ADVICE.HEADING), # remove NAs
         ADVICE.HEADING = gsub(",(?=[^,]+$)", " and", ADVICE.HEADING, perl = TRUE), # remove last comma and replace with "and"
         ADVICE.HEADING = ifelse(count == 1,
                                 paste0(ADVICE.HEADING, " Ecoregion"),
                                 paste0(ADVICE.HEADING, " Ecoregions")),
         ADVICE.HEADING = ifelse(count == 12,
                                 "the Northeast Atlantic and adjacent waters",
                                 ADVICE.HEADING), 
         ADVICE.HEADING = ifelse(count/length(ecoregions) >= .70 &
                                   !grepl("Arctic Ocean", ADVICE.HEADING) &
                                   !grepl("Baltic Sea", ADVICE.HEADING),
                                 "Northeast Atlantic",
                                 ADVICE.HEADING),
         ADVICE.HEADING = ifelse(count/length(ecoregions) >= .70 &
                                   grepl("Arctic Ocean", ADVICE.HEADING) &
                                   !grepl("Baltic Sea", ADVICE.HEADING),
                                 "Northeast Atlantic and Arctic Ocean",
                                 ADVICE.HEADING),
         ADVICE.HEADING = ifelse(count/length(ecoregions) >= .70 &
                                   !grepl("Arctic Ocean", ADVICE.HEADING) &
                                   grepl("Baltic Sea", ADVICE.HEADING),
                                 "Northeast Atlantic and Baltic Sea",
                                 ADVICE.HEADING),
         CAPTION.NAME = gsub("\\s*\\([^\\)]+\\)", "", as.character(STOCK.NAME))) %>%
  full_join(sectionNum, by = c("STOCK.CODE", "OLD.CODE")) %>%
  select(STOCK.CODE, OLD.CODE, STOCK.NAME, SECTION.NUMBER, 
         CATEGORY, ADVICE.HEADING, CAPTION.NAME, SPECIES.NAME, 
         COMMON.NAME)

##################################
# Organize SharePoint file paths # 
##################################
# Note: You must log in to SharePoint and have this drive mapped
sharePoint <- "//community.ices.dk@SSL/DavWWWRoot/"

# 2015 (for Cat 3+ stocks with multi-year advice)
Advice2015 <- list.files(paste0(sharePoint, "Advice/Advice2015/"))
folderNames2015 <- c("BalticSea", "BarentsSea", "BayOfBiscay", 
                     "CelticSea", "FaroePlateau", "Iceland",
                     "NASalmon","NorthSea", "Widely")

adviceList2015 <- lapply(Advice2015[Advice2015 %in% folderNames2015],
                                function(x) list.files(paste0(sharePoint, 
                                                              "Advice/Advice2015/", x, "/Released_Advice")))
names(adviceList2015) <- folderNames2015
fileList2015 <- do.call("rbind", lapply(adviceList2015,
                                        data.frame, 
                                        stringsAsFactors = FALSE))
colnames(fileList2015) <- "STOCK.CODE"
fileList2015$file.path <- sprintf(paste0(sharePoint,
                                         "Advice/Advice2015/%s/Released_Advice/%s"), 
                                  gsub("\\..*", "", row.names(fileList2015)),
                                  fileList2015$STOCK.CODE)
fileList2015$STOCK.CODE <- tolower(gsub("\\.docx*", "", fileList2015$STOCK.CODE))

# 2016 (for stocks with annual advice)
Advice2016 <- list.files(paste0(sharePoint, "Advice/Advice2016/"))
folderNames2016 <- c("BalticSea", "BarentsSea", "Biscay", 
                     "CelticSea", "Faroes", "Iceland",
                     "NorthSea", "Salmon", "Widely")
adviceList2016 <- lapply(Advice2016[Advice2016 %in% c("BalticSea", "BarentsSea", "Biscay",
                                                      "CelticSea", "Faroes", "Iceland",
                                                      "NorthSea", "Salmon", "Widely")],
       function(x) list.files(paste0(sharePoint,
                                     "Advice/Advice2016/", x, "/Released_Advice")))
names(adviceList2016) <- folderNames2016
fileList2016 <- do.call("rbind", lapply(adviceList2016,
                                        data.frame, 
                                        stringsAsFactors = FALSE))
colnames(fileList2016) <- "STOCK.CODE"
fileList2016$file.path <- sprintf(paste0(sharePoint, 
                                         "Advice/Advice2016/%s/Released_Advice/%s"), 
                                  gsub("\\..*", "", row.names(fileList2016)),
                                  fileList2016$STOCK.CODE)
fileList2016$STOCK.CODE <- tolower(gsub("\\.docx*", "", fileList2016$STOCK.CODE))

# Do a little clean up of inconsistent file names
stockList$OLD.CODE <- gsub("pan", "pand", stockList$OLD.CODE)

# Join with the SLD so we can have file paths for each stock
fileStock2016 <- stockList %>%
  left_join(fileList2016, by = c("OLD.CODE" = "STOCK.CODE")) 

fileStock2015 <- fileStock2016 %>%
  filter(is.na(file.path)) %>%
  select(-file.path) %>%
  left_join(fileList2015, by = c("OLD.CODE" = "STOCK.CODE")) 

fileStock2015$file.path[fileStock2015$OLD.CODE == "nep-oth-6a"] <-  "//community.ices.dk@SSL/DavWWWRoot/Advice/Advice2015/CelticSea/Released_Advice/nep-oth-6a_FOR AUTUMN.docx"
fileStock2015$file.path[fileStock2015$OLD.CODE == "nep-oth-7"] <-  "//community.ices.dk@SSL/DavWWWRoot/Advice/Advice2015/CelticSea/Released_Advice/nep-oth-7_FOR AUTUMN.docx"
fileStock2015$file.path[fileStock2015$OLD.CODE == "nep-oth-4"] <-  "//community.ices.dk@SSL/DavWWWRoot/Advice/Advice2015/NorthSea/Released_Advice/nep-oth.docx"

fileList <- rbind(fileStock2015, fileStock2016[!is.na(fileStock2016$file.path),])

outList <- fileList[is.na(fileList$file.path),]
fileList <- fileList[!is.na(fileList$file.path),]

####################################
# Wrap it up and write some drafts #
####################################

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

#########################################################
# Create Draft given the Stock List Database and tables #
#########################################################
#
stock.code <- "mur.27.67a-ce-k89a"
# draftDir <- paste0(sharePoint, "/admin/AdvisoryProgramme/Personal folders/Scott/Advice_Template_Testing/drafts/")
# draftDir <-"~/git/ices-dk/AdviceTemplate/"

createDraft <- function(stock.code, output_dir = draftDir) {

  ########################
  # Load Table functions #
  ########################
  source("~/git/ices-dk/AdviceTemplate/table_functions.R")
  
 # Load up the proper template
  if(stockList$CATEGORY[stockList$STOCK.CODE == stock.code] %in% c(1,2)) {
    template <- paste0(sharePoint, "Advice/Advice2017/Advice documents/Draft advice 2017/advice_template_2017_cat12.docx")
  }
  if(stockList$CATEGORY[stockList$STOCK.CODE == stock.code] %in% c(3,4)) {
    template <- paste0(sharePoint, "Advice/Advice2017/Advice documents/Draft advice 2017/advice_template_2017_cat34.docx")  
  }
  if(stockList$CATEGORY[stockList$STOCK.CODE == stock.code] %in% c(5,6)) {
    # template <- paste0(sharePoint, "Advice/Advice2017/Advice documents/Draft advice 2017/advice_template_2017_cat56.docx")
    template <- "~/git/ices-dk/AdviceTemplate/advice_template_2017_cat56.docx"
  }
  if(!file.exists(template)) {
    stop(paste0("Check your file path to make sure the category ", stockList$CATEGORY[stockList$STOCK.CODE == stock.code],
                " template is available.\nCurrent file path: ",
                template))
  } else {
    draftDoc <- docx(template = template,
                     title = stock.code)
  }
  
  
  # Build up the stock pot
  # stockPot <- createStockPot()

  # Add the heading, headers, and footers
  draftDoc <- addParagraph(draftDoc, 
                           value = headerPot(stock.code),
                           par.properties = base_par_prop,
                           bookmark = "TITLE_PAGE_HEADER")

  draftDoc <- addParagraph(draftDoc, 
                           value = headingPot(stock.code),
                           par.properties = base_par_prop,
                           bookmark = "TITLE_HEADING")

  draftDoc <- addParagraph(draftDoc, 
                           value = footerPot(stock.code),
                           par.properties = base_par_prop,
                           bookmark = "FIRST_PAGE_FOOTER")
  
  draftDoc <- addParagraph(draftDoc, 
                           value = headerPot(stock.code),
                           par.properties = base_par_prop,
                           bookmark = "SECOND_PAGE_HEADER")
  
  draftDoc <- addParagraph(draftDoc, 
                           value = footerPot(stock.code),
                           par.properties = base_par_prop,
                           bookmark = "SECOND_PAGE_FOOTER")
  #
  draftDoc <- addParagraph(draftDoc, 
                           value = captionPot(stock.code, 
                                              type = "Figure",
                                              number = 2, 
                                              text = "TEXT"),
                           par.properties = base_par_prop,
                           bookmark = "STOCK_SUMMARY_FIGURE_CAPTION")
  #
  draftDoc <- addParagraph(draftDoc, 
                           value = captionPot(stock.code, 
                                              type = "Table",
                                              number = 1, 
                                              text = "TEXT"),
                           par.properties = base_par_prop,
                           bookmark = "STOCK_STATUS_TABLE_CAPTION")
  #
  draftDoc <- addParagraph(draftDoc, 
                           value = captionPot(stock.code,
                                              type = "Table", 
                                              number = 1, 
                                              text = "CATCH DISTRIBUTION TABLE."),
                           par.properties = base_par_prop,
                           bookmark = "CATCH_DISTRIBUTION_TABLE_CAPTION")
  #
  draftDoc <- addParagraph(draftDoc, 
                           value = captionPot(stock.code, type = "Table", number = 1, text = "CATCH DISTRIBUTION TABLE."),
                           par.properties = base_par_prop,
                           bookmark = "CATCH_HISTORY_TABLE_CAPTION")
  
  draftDoc <- addParagraph(draftDoc, 
                           value = captionPot(stock.code, type = "Table", number = 1, text = "CATCH ADVICE_BASIS_TABLE_CAPTION."),
                           par.properties = base_par_prop,
                           bookmark = "ADVICE_BASIS_TABLE_CAPTION")
  
  draftDoc <- addParagraph(draftDoc, 
                           value = captionPot(stock.code, type = "Table", number = 1, text = "ADVICE_HISTORY_TABLE_CAPTION"),
                           par.properties = base_par_prop,
                           bookmark = "ADVICE_HISTORY_TABLE_CAPTION")
  
  draftDoc <- addParagraph(draftDoc, 
                           value = stockList$COMMON.NAME[stockList$STOCK.CODE == stock.code],
                           par.properties = base_par_prop,
                           bookmark = "COMMON_NAME")

  # Add tables
  adviceHistoryList <- advice_history_table(stock.code)
  adviceHistoryNames <- names(adviceHistoryList)

  for(tn in 1:length(adviceHistoryNames)) {
    bm <- ifelse(tn >= 2,
                 paste0("ADVICE_HISTORY_TABLE_", tn),
                 "ADVICE_HISTORY_TABLE")  
    draftDoc <- addFlexTable(draftDoc,
                             flextable = adviceHistoryList[[tn]],
                             bookmark = bm)
  }
  
  # names()
  assessmentBasisList <- assessment_basis_table(stock.code)
  assessmentBasisNames <- names(assessmentBasisList)
  
  for(tn in 1:length(assessmentBasisNames)) {
    bm <- ifelse(tn >= 2,
                 paste0("ADVICE_HISTORY_TABLE_", tn),
                 "ADVICE_HISTORY_TABLE")
  draftDoc <- addFlexTable(draftDoc,
                           flextable = assessmentBasisList[[tn]],
                           bookmark = bm)
  }
  
  if(stockList$CATEGORY[stockList$STOCK.CODE == stock.code] %in% c(1,2)) {

    draftDoc <- addParagraph(draftDoc, 
                             value = captionPot(stock.code, type = "Table", number = 1, text = "CATCH OPTIONS TABLE"),
                             par.properties = base_par_prop,
                             bookmark = "CATCH_OPTIONS_BASIS_TABLE_CAPTION")
    
    draftDoc <- addFlexTable(draftDoc,
                 flextable = catch_options_basis_table(stock.code),
                 bookmark = "CATCH_OPTIONS_BASIS_TABLE")
  }
  
  writeDoc(doc = draftDoc,
           file = "~/git/ices-dk/AdviceTemplate/tester_4.docx")
  
  writeDoc(doc = draftDoc,
           file = paste0(draftDir, stock.code, "_TEST.docx"))
  
  # stockPot$TITLE_HEADING <-      headingPot(stock.code)
  # stockPot$FIRST_PAGE_FOOTER  <- footerPot(stock.code) # BOOK_FOOTER 
  # stockPot$SECOND_PAGE_HEADER <- headerPot2(stock.code) # SECOND_HEADER 
  # stockPot$SECOND_PAGE_FOOTER <- footerPot(stock.code) # BOOK_FOOTER2 
  # stockPot$STOCK_SUMMARY_FIGURE_CAPTION  <- captionPot(stock.code, type = "Figure", number = 2, text = "Summary of stock assessment")
  # stockPot$STOCK_SUMMARY_FIGURE  <- NULL
  # stockPot$STOCK_STATUS_TABLE_CAPTION  <-  captionPot(stock.code, type = "Table", number = 2, text = "BBBBBBBBB")
  # stockPot$STOCK_STATUS_TABLE  <- NULL
  stockPot$CATCH_OPTIONS_BASIS_TABLE_CAPTION   <- captionPot(stock.code, type = "Table", number = 3, text = "BBBBBBBBB")
  if(stockList$CATEGORY[stockList$STOCK.CODE == stock.code] %in% c(1,2)) {
    stockPot$CATCH_OPTIONS_BASIS_TABLE  <- catch_options_basis_table(stock.code)
  } else {
    stockPot$CATCH_OPTIONS_BASIS_TABLE  <- NULL
  }
  stockPot$CATCH_OPTIONS_TABLE_CAPTION  <-  captionPot(stock.code, type = "Table", number = 4, text = "BBBBBBBBB")
  stockPot$CATCH_OPTIONS_TABLE  <- NULL
  stockPot$ADVICE_BASIS_TABLE_CAPTION   <- captionPot(stock.code, type = "Table", number = 5, text = "BBBBBBBBB") 
  stockPot$ADVICE_BASIS_TABLE   <- advice_basis_table(stock.code)

  if(stockList$CATEGORY[stockList$STOCK.CODE == stock.code] %in% c(1,2)) {
    stockPot$ASSESSMENT_HISTORY_FIGURE_CAPTION  <- captionPot(stock.code, type = "Figure", number = 4, text = "BBBBBBBBB") 
    stockPot$ASSESSMENT_HISTORY_FIGURE  <- NULL
  } else {
    stockPot$ASSESSMENT_HISTORY_FIGURE_CAPTION  <- NULL
    stockPot$ASSESSMENT_HISTORY_FIGURE  <- NULL
  }
  
  stockPot$REFERENCE_POINTS_TABLE_CAPTION  <- captionPot(stock.code, type = "Table", number = 5, text = "BBBBBBBBB") 
  stockPot$REFERENCE_POINTS_TABLE  <- reference_points_table(stock.code)
  stockPot$ASSESSMENT_BASIS_TABLE_CAPTION <- captionPot(stock.code, type = "Table", number = 5, text = "BBBBBBBBB") 
  stockPot$ASSESSMENT_BASIS_TABLE   <- assessment_basis_table(stock.code)
  stockPot$ADVICE_HISTORY_TABLE_CAPTION <- captionPot(stock.code, type = "Table", number = 5, text = "BBBBBBBBB") 
  stockPot$ADVICE_HISTORY_TABLE  <- advice_history_table(stock.code)
  stockPot$CATCH_DISTRIBUTION_TABLE_CAPTION <- captionPot(stock.code, type = "Table", number = 5, text = "BBBBBBBBB") 
  stockPot$CATCH_DISTRIBUTION_TABLE  <- NULL
  stockPot$CATCH_HISTORY_TABLE_CAPTION <- captionPot(stock.code, type = "Table", number = 5, text = "BBBBBBBBB") 
  stockPot$CATCH_HISTORY_TABLE <- NULL
  stockPot$COMMON_NAME <- NULL
  

  addParagraph(draftDoc, 
               value = stockPot$TITLE_PAGE_HEADER,
               par.properties = base_par_prop,
               bookmark = "TITLE_PAGE_HEADER")
  
  addParagraph(draftDoc, 
               value = stockPot$TITLE_HEADING,
               par.properties = base_par_prop,
               bookmark = "TITLE_HEADING")
  
  addParagraph(draftDoc,
               value = stockPot$FIRST_PAGE_FOOTER,
               par.properties = base_par_prop,
               bookmark = "FIRST_PAGE_FOOTER")
  
  addParagraph(draftDoc, 
               value = stockPot$SECOND_PAGE_HEADER,
               par.properties = base_par_prop,
               bookmark = "SECOND_PAGE_HEADER")
  
  # addParagraph(draftDoc, 
  #              value = stockPot$ADVICE_HISTORY_TABLE$T7,
  #              par.properties = base_par_prop,
  #              bookmark = "ADVICE_HISTORY_TABLE")
  
  addFlexTable(draftDoc,
               flextable = stockPot$ADVICE_HISTORY_TABLE$T7,
               bookmark = "ADVICE_HISTORY_TABLE")
  
  # addFlexTable(draftDoc,
  #              flextable = stockPot$REFERENCE_POINTS_TABLE$T5,
  #              bookmark = "REFERENCE_POINTS_TABLE")

  addFlexTable(draftDoc,
               flextable = stockPot$ADVICE_BASIS_TABLE$T4,
               bookmark = "ADVICE_BASIS_TABLE")
  
  
    
  addParagraph(draftDoc, 
               value = stockPot$STOCK_SUMMARY_FIGURE_CAPTION,
               par.properties = base_par_prop,
               bookmark = "STOCK_SUMMARY_FIGURE_CAPTION")
  
  addParagraph(draftDoc, 
               value = stockPot$STOCK_STATUS_TABLE_CAPTION,
               par.properties = base_par_prop,
               bookmark = "STOCK_STATUS_TABLE_CAPTION")
  
  addParagraph(draftDoc, 
               value = stockPot$CATCH_OPTIONS_BASIS_TABLE_CAPTION,
               par.properties = base_par_prop,
               bookmark = "CATCH_OPTIONS_BASIS_TABLE_CAPTION")
  


writeDoc(doc = draftDoc,
         file = paste0(draftFolder, stock.code, "_TEST.docx"))




stockPot <- list("TITLE_PAGE_HEADER", # "ECOREGION_HEADER"
                 "TITLE_HEADING",  # "ADVICE_HEADING"
                 "FIRST_PAGE_FOOTER", #"BOOK_FOOTER"
                 "SECOND_PAGE_HEADER", #"SECOND_HEADER"
                 "SECOND_PAGE_FOOTER", #"BOOK_FOOTER2"
                 "STOCK_SUMMARY_FIGURE_CAPTION",
                 #"STOCK_SUMMARY_FIGURE", 
                 "STOCK_STATUS_TABLE_CAPTION",
                 #"STOCK_STATUS_TABLE",
                 "CATCH_OPTIONS_BASIS_TABLE_CAPTION", 
                 "CATCH_OPTIONS_BASIS_TABLE", #catch_options_basis_table()
                 "CATCH_OPTIONS_TABLE_CAPTION",
                 #"CATCH_OPTIONS_TABLE",
                 "ADVICE_BASIS_TABLE_CAPTION", # 
                 "ADVICE_BASIS_TABLE", #advice_basis_table()
                 "ASSESSMENT_HISTORY_FIGURE_CAPTION",
                 #"ASSESSMENT_HISTORY_FIGURE",
                 "REFERENCE_POINTS_TABLE_CAPTION",
                 "REFERENCE_POINTS_TABLE",
                 "ASSESSMENT_BASIS_TABLE_CAPTION",
                 "ASSESSMENT_BASIS_TABLE", #assessment_basis_table()
                 "ADVICE_HISTORY_TABLE_CAPTION",
                 "ADVICE_HISTORY_TABLE", #assessment_history_table()
                 "CATCH_DISTRIBUTION_TABLE_CAPTION",
                 #"CATCH_DISTRIBUTION_TABLE",
                 "CATCH_HISTORY_TABLE_CAPTION"
                 #,"CATCH_HISTORY_TABLE"
)



# draft_from_template <- function(stock.code) {
#   if()
  
  
  
  
  
  
}





grabTables <- function(stock.code) {
  
  # Table 1: State of the stock and fishery, relative to reference points
  # stockExploitationData - EG to add from SAG
  # stockExploitationTable - EG to add from SAG
  
  
  
  # Category 1
  if(fileList$CATEGORY[fileList$STOCK.CODE == stock.code] == 1)  {
  }
  # Table 2: The basis for the catch options.
  # catchOptionsBasisData
  # catchOptionsBasisTable
  
  
  if(grepl("nep", stock.code)) {
  }
  # Table 3: The catch options. All weights are in thousand tonnes.
  # catchOptionsData
  # catchOptionsTable
  
if(grepl("nep", stock.code)) {
}

  # Table 4: The basis of advice.
  # adviceBasisData
  # adviceBasisTable
  
  # Table 5: Reference points, values, and their technical basis.
  # referencePointsData
  # referencePointsTable

  # Table 6: The basis of the assessment
  # assessmentBasisData
  # assessmentBasisTable

  # Table 7: History of ICES advice, the agreed TAC, and ICES estimates of catch. All weights are in thousand tonnes.
  # adviceHistoryData
  # adviceHistoryTable

  # Table 8: Catch distribution by fleet in 2016 as estimated by ICES.
  # catchDistributionData
  # catchDistributionTable
  
  # Table 9: 
  # catchHistoryData
  # catchHistoryTable

  # Table 10: 
  # assessmentSummaryData
  # assessmentSummaryTable


  
  # Category 2
  if(fileList$CATEGORY[fileList$STOCK.CODE == stock.code] == 2) {
  }
# stockExploitationData
# stockExploitationTable

# catchOptionsBasisData
# catchOptionsBasisTable

# catchOptionsData
# catchOptionsTable

# adviceBasisData
# adviceBasisTable

# referencePointsData
# referencePointsTable

# assessmentBasisData
# assessmentBasisTable

# adviceHistoryData
# adviceHistoryTable

# catchHistoryData
# catchHistoryTable

# assessmentSummaryData
# assessmentSummaryTable


  # Category 3
  if(fileList$CATEGORY[fileList$STOCK.CODE == stock.code] == 3) {
  }
if(grepl("nep", stock.code)) {
}

# stockExploitationData
# stockExploitationTable

# catchOptionsBasisData
# catchOptionsBasisTable

# catchOptionsData
# catchOptionsTable

# adviceBasisData
# adviceBasisTable

# referencePointsData
# referencePointsTable

# assessmentBasisData
# assessmentBasisTable

# adviceHistoryData
# adviceHistoryTable

# catchHistoryData
# catchHistoryTable

# assessmentSummaryData
# assessmentSummaryTable


  # and nephrops
  grep("Fishing pressure", bigList[["nep.fu.30"]])
  
  # Category 4
  if(fileList$CATEGORY[fileList$STOCK.CODE == stock.code] == 4) {
  }
  
  # stockExploitationData
  # stockExploitationTable
  
  # catchOptionsBasisData
  # catchOptionsBasisTable
  
  # catchOptionsData
  # catchOptionsTable
  
  # adviceBasisData
  # adviceBasisTable
  
  # referencePointsData
  # referencePointsTable
  
  # assessmentBasisData
  # assessmentBasisTable
  
  # adviceHistoryData
  # adviceHistoryTable
  
  # catchHistoryData
  # catchHistoryTable
  
  # assessmentSummaryData
  # assessmentSummaryTable
  
  if(fileList$CATEGORY[fileList$STOCK.CODE == stock.code] %in% c(5, 6)) {
  }
  if(grepl("nep", stock.code)) {
  }

  # stockExploitationData
  # stockExploitationTable
  
  # catchOptionsBasisData
  # catchOptionsBasisTable
  
  # catchOptionsData
  # catchOptionsTable
  
  # adviceBasisData
  # adviceBasisTable
  
  # referencePointsData
  # referencePointsTable
  
  # assessmentBasisData
  # assessmentBasisTable
  
  # adviceHistoryData
  # adviceHistoryTable
  
  # catchHistoryData
  # catchHistoryTable
  
  # assessmentSummaryData
  # assessmentSummaryTable
  
  
  # Category 5 and 6
 
  # if(the table contains these column headers)
  # create the following table
  # if(the table contains these column headers)
  # create the following table
  
  grep("Fishing pressure", bigList[stock.code])
  grep("Stock size", bigList[["pok.27.5b"]])
  grep("Year", bigList[["pok.27.5b"]])
  grep("ICES advice.", bigList[["pok.27.5b"]])
  grep("Advice basis", bigList[["pok.27.5b"]])
  
  grep("Variable", bigList[["pok.27.5b"]])
  grep("Value", bigList[["pok.27.5b"]])
  grep("Framework", bigList[["pok.27.5b"]])
  
  
} # close grabTables function

#to do: add bookmarks to template and see if template of template works. 
