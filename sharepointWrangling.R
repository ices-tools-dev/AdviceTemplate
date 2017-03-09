rm(list = ls())
#############################
# Install and load packages #
#############################
# library(readxl) # not necessary once SLD is up and running
library(tidyverse)
library(docxtractr)
library(ReporteRs)
library(jsonlite)
library(stringr)
# library(icesSharePoint)

# Note: You must log in to SharePoint and have this drive mapped
sharePoint <- "//community.ices.dk@SSL/DavWWWRoot/"
if(!dir.exists(paste0(sharePoint, "Advice/Advice2017/"))) {
  stop("Note: You must be on the ICES network and log in to SharePoint and have this drive mapped")
}

########################
# Load Table functions #
########################
source("~/git/ices-dk/AdviceTemplate/table_functions.R")


#############################################
# Organizing stock list and section numbers #
#############################################
rawsd <- fromJSON("http://sd.ices.dk/services/odata3/StockListDWs3?$filter=ActiveYear%20eq%202017")$value

sl_ecoregion <- rawsd %>%
  select(StockCode, OldStockCode, Description, SpeciesCommonName,
         SpeciesScientificName, 
         SectionNumber,
         DataCategory, 
         AdviceHeading = EcoRegion,
         AdviceReleaseDate,
         ExpertGroup, 
         AdviceDraftingGroup) %>%
  mutate(count = str_count(AdviceHeading, ",") + 1,
         AdviceReleaseDate = as.Date(AdviceReleaseDate, format = "%d/%m/%y"),
         PubDate = format(AdviceReleaseDate, "%e %B %Y"),
         DraftURL = gsub("(.*?)(\\..*)", "\\1", SectionNumber),
         ExpertURL = paste0("http://www.ices.dk/community/groups/Pages/", ExpertGroup, ".aspx"),
         DraftURL = recode(DraftURL, "2" = paste0(sharePoint, "Advice/Advice2017/Iceland/Draft_advice/")),
         DraftURL = recode(DraftURL, "3" = paste0(sharePoint, "Advice/Advice2017/BarentsSea/Draft_advice/")),
         DraftURL = recode(DraftURL, "4" = paste0(sharePoint, "Advice/Advice2017/Faroes/Draft_advice/")),
         DraftURL = recode(DraftURL, "5" = paste0(sharePoint, "Advice/Advice2017/CelticSea/Draft_advice/")),
         DraftURL = recode(DraftURL, "6" = paste0(sharePoint, "Advice/Advice2017/NorthSea/Draft_advice/")),
         DraftURL = recode(DraftURL, "7" = paste0(sharePoint, "Advice/Advice2017/BayOfBiscay/Draft_advice/")),
         DraftURL = recode(DraftURL, "8" = paste0(sharePoint, "Advice/Advice2017/BalticSea/Draft_advice/")),
         DraftURL = recode(DraftURL, "9" = paste0(sharePoint, "Advice/Advice2017/Widely/Draft_advice/")),
         DraftURL = recode(DraftURL, "10" = paste0(sharePoint, "Advice/Advice2017/Salmon/Draft_advice/")))

sl <- sl_ecoregion %>%
  mutate(Description = gsub(" $","", Description, perl = TRUE), # remove trailing spaces
         DataCategory = gsub("\\..*$", "", DataCategory), # Just use the first digit
         AdviceHeading = gsub("Ecoregion", "", AdviceHeading), # remove Ecoregion
         AdviceHeading = gsub(",(?=[^,]+$)", " and", AdviceHeading, perl = TRUE), # remove last comma and replace with "and"
         AdviceHeading = ifelse(count == 1,
                                 paste0(AdviceHeading, "Ecoregion"),
                                 paste0(AdviceHeading, "Ecoregions")),
         AdviceHeading = ifelse(count == 12,
                                 "the Northeast Atlantic and adjacent waters",
                                 AdviceHeading), 
         AdviceHeading = ifelse(count/12 >= .70 &
                                   !grepl("Arctic Ocean", AdviceHeading) &
                                   !grepl("Baltic Sea", AdviceHeading),
                                 "Northeast Atlantic",
                                 AdviceHeading),
         AdviceHeading = ifelse(count/12 >= .70 &
                                   grepl("Arctic Ocean", AdviceHeading) &
                                   !grepl("Baltic Sea", AdviceHeading),
                                 "Northeast Atlantic and Arctic Ocean",
                                 AdviceHeading),
         AdviceHeading = ifelse(count/12 >= .70 &
                                   !grepl("Arctic Ocean", AdviceHeading) &
                                   grepl("Baltic Sea", AdviceHeading),
                                 "Northeast Atlantic and Baltic Sea",
                                 AdviceHeading),
         CaptionName = gsub("\\s*\\([^\\)]+\\)", "", as.character(Description))) %>%
  select(StockCode, Description,
         PubDate, DataCategory, AdviceHeading, CaptionName, 
         SpeciesScientificName, OldStockCode, ExpertGroup,
         SpeciesCommonName, ExpertURL, DraftURL,
         AdviceDraftingGroup)

##################################
# Organize SharePoint file paths # 
##################################
# This is mostly manual and custom built to the existing folder structure. More rigid definitions would make this more clear and
# less apt to breaking
#
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
colnames(fileList2015) <- "StockCode"
fileList2015$file.path <- sprintf(paste0(sharePoint,
                                         "Advice/Advice2015/%s/Released_Advice/%s"), 
                                  gsub("\\..*", "", row.names(fileList2015)),
                                  fileList2015$StockCode)
fileList2015$StockCode <- tolower(gsub("\\.docx*", "", fileList2015$StockCode))

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
colnames(fileList2016) <- "StockCode"
fileList2016$file.path <- sprintf(paste0(sharePoint, 
                                         "Advice/Advice2016/%s/Released_Advice/%s"), 
                                  gsub("\\..*", "", row.names(fileList2016)),
                                  fileList2016$StockCode)
fileList2016$StockCode <- tolower(gsub("\\.docx*", "", fileList2016$StockCode))

# Do a little clean up of inconsistent file names
sl$OldStockCode <- gsub("pan", "pand", sl$OldStockCode)

# Join with the SLD so we can have file paths for each stock
fileStock2016 <- sl %>%
  left_join(fileList2016, by = c("OldStockCode" = "StockCode")) 

fileStock2015 <- fileStock2016 %>%
  filter(is.na(file.path)) %>%
  select(-file.path) %>%
  left_join(fileList2015, by = c("OldStockCode" = "StockCode")) 

fileStock2015$file.path[fileStock2015$OldStockCode == "nep-oth-6a"] <-  "//community.ices.dk@SSL/DavWWWRoot/Advice/Advice2015/CelticSea/Released_Advice/nep-oth-6a_FOR AUTUMN.docx"
fileStock2015$file.path[fileStock2015$OldStockCode == "nep-oth-7"] <-  "//community.ices.dk@SSL/DavWWWRoot/Advice/Advice2015/CelticSea/Released_Advice/nep-oth-7_FOR AUTUMN.docx"
fileStock2015$file.path[fileStock2015$OldStockCode == "nep-oth-4"] <-  "//community.ices.dk@SSL/DavWWWRoot/Advice/Advice2015/NorthSea/Released_Advice/nep-oth.docx"

fileList <- rbind(fileStock2015, fileStock2016[!is.na(fileStock2016$file.path),])

outList <- fileList[is.na(fileList$file.path),]
fileList <- fileList[!is.na(fileList$file.path),]
fileList$URL <- paste0(gsub("//community.ices.dk@SSL/DavWWWRoot/", "https://community.ices.dk/", fileList$file.path), 
                       "?Web=1")

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


createDraft <- function(stock.code) {

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
  
  
  if(file.exists(paste0(draft.url, stock.code, ".docx"))) {
    stop("WATCH OUT! You are trying to over-write a file!!!")
  }
  ####################################
  # Wrap it up and write some drafts #
  ####################################
  
  
 # Load up the proper template
  if(data.category %in% c(1,2)) {
    template <- paste0(sharePoint, "Advice/Advice2017/Advice documents/Draft advice 2017 - Secretariat use/advice_template_2017_cat12.docx")
    catchOptionsCaption <- "Annual catch options. All weights are in tonnes."
  }
  if(data.category %in%  c(3,4)) {
    template <- paste0(sharePoint, "Advice/Advice2017/Advice documents/Draft advice 2017 - Secretariat use/advice_template_2017_cat34.docx")  
    catchOptionsCaption <- "For stocks in ICES categories 3-6, one catch option is provided."
  }
  if(data.category %in% c(5,6)) {
    template <- paste0(sharePoint, "Advice/Advice2017/Advice documents/Draft advice 2017 - Secretariat use/advice_template_2017_cat56.docx")
    catchOptionsCaption <- "For stocks in ICES categories 3-6, one catch option is provided."
  }
  if(!file.exists(template)) {
    stop(paste0("Check your file path to make sure the category ", sl$DataCategory[sl$StockCode == stock.code],
                " template is available.\nCurrent file path: ",
                template))
  } else {
    draftDoc <- docx(template = template,
                     title = stock.code)
  }
  
  # Build up the stock pot
  # Add the heading, headers, and footers
  draftDoc <- addParagraph(draftDoc, 
                           value = titlePot(ecoregion.name, pub.date),
                           par.properties = base_par_prop,
                           bookmark = "TITLE_PAGE_HEADER")
  
  draftDoc <- addParagraph(draftDoc, 
                           value = pot(stock.code, format = header_text_prop),
                           par.properties = base_par_prop,
                           bookmark = "STOCK_CODE_HEADER")

  draftDoc <- addParagraph(draftDoc, 
                           value = pot(stock.name, format = heading_text_prop),
                           par.properties = base_par_prop,
                           bookmark = "TITLE_HEADING")
  
  draftDoc <- addParagraph(draftDoc, 
                           value = pot("For reference in drafting the advice, please find previous advice here: ",
                                       format = base_text_prop) +
                             pot(url.name,
                                       format = base_text_prop,
                                       hyperlink = url.name),
                           par.properties = base_par_prop,
                           bookmark = "PREVIOUS_ADVICE_URL")

  draftDoc <- addParagraph(draftDoc, 
                           value = headerPot(stock.code, pub.date),
                           par.properties = base_par_prop,
                           bookmark = "SECOND_PAGE_HEADER")
  
  draftDoc <- addParagraph(draftDoc, 
                           value = pot(stock.code, format = header_text_prop),
                           par.properties = base_par_prop,
                           bookmark = "STOCK_CODE_HEADER_2")

  figCount <- 1
  tabCount <- 1

  # SAG figures
  draftDoc <- addParagraph(draftDoc, 
                           value = captionPot(stock.code, 
                                              type = "Figure",
                                              caption.number = figCount, 
                                              caption.name = caption.name,
                                              caption.text = "Summary of the stock assessment."),
                           par.properties = base_par_prop,
                           bookmark = "STOCK_SUMMARY_FIGURE_CAPTION")
  figCount <- figCount + 1
  # Tick-mark table
  draftDoc <- addParagraph(draftDoc, 
                           value = captionPot(stock.code, 
                                             type = "Table",
                                             caption.number = tabCount, 
                                             caption.name = caption.name,
                                             caption.text = "State of the stock and fishery relative to reference points."),
                           par.properties = base_par_prop,
                           bookmark = "STOCK_STATUS_TABLE_CAPTION")
  tabCount <- tabCount + 1
  
  # Basis of catch option table
    draftDoc <- addParagraph(draftDoc,
                             value = captionPot(stock.code,
                                                caption.number = tabCount,
                                                type = "Table",
                                                caption.name = caption.name,
                                                caption.text = "The basis for the catch options."),
                             par.properties = base_par_prop,
                             bookmark = "CATCH_OPTIONS_BASIS_TABLE_CAPTION")
    tabCount <- tabCount + 1
    
    if(data.category %in% c(1, 2)) {  
    draftDoc <- addParagraph(draftDoc,
                             value = captionPot(stock.code,
                                                caption.number = figCount,
                                                type = "Figure",
                                                caption.name = caption.name,
                                                caption.text = "Historical assessment results."),
                             par.properties = base_par_prop,
                             bookmark = "HISTORICAL_ASSESSMENT_FIGURE_CAPTION")


  # Catch options table
  draftDoc <- addParagraph(draftDoc, 
                           value = captionPot(stock.code, 
                                              type = "Table",
                                              caption.number = tabCount, 
                                              caption.name = caption.name,
                                              caption.text = catchOptionsCaption),
                           par.properties = base_par_prop,
                           bookmark = "CATCH_OPTIONS_TABLE_CAPTION")
  tabCount <- tabCount + 1
  }
  # Basis of the advice
  draftDoc <- addParagraph(draftDoc, 
                           value = captionPot(stock.code, 
                                              type = "Table",
                                              caption.number = tabCount, 
                                              caption.name = caption.name,
                                              caption.text = "The basis of the advice."),
                           par.properties = base_par_prop,
                           bookmark = "ADVICE_BASIS_TABLE_CAPTION")
  tabCount <- tabCount + 1

  # Basis of catch option table
  if(data.category %in% c(1, 2, 3, 4)) {  
    draftDoc <- addParagraph(draftDoc,
                             value = captionPot(stock.code,
                                                type = "Table",
                                                caption.number = tabCount,
                                                caption.name = caption.name,
                                                caption.text = "Reference points, values, and their technical basis."),
                             par.properties = base_par_prop,
                             bookmark = "REFERENCE_POINT_TABLE_CAPTION")
    tabCount <- tabCount + 1
  }
  
  # Basis of the assessment
  draftDoc <- addParagraph(draftDoc,
                           value = captionPot(stock.code,
                                              type = "Table",
                                              caption.number = tabCount,
                                              caption.name = caption.name,
                                              caption.text = "Basis of assessment and advice."),
                           par.properties = base_par_prop,
                           bookmark = "ASSESSMENT_BASIS_TABLE_CAPTION")
  tabCount <- tabCount + 1
  
  
  # Advice history
  draftDoc <- addParagraph(draftDoc,
                           value = captionPot(stock.code,
                                              type = "Table",
                                              caption.number = tabCount,
                                              caption.name = caption.name,
                                              caption.text = "ICES advice and official landings. All weights are in thousand tonnes."),
                           par.properties = base_par_prop,
                           bookmark = "ADVICE_HISTORY_TABLE_CAPTION")
  tabCount <- tabCount + 1
  
  # Catch Distribution
  draftDoc <- addParagraph(draftDoc, 
                           value = captionPot(stock.code, 
                                              type = "Table",
                                              caption.number = tabCount, 
                                              caption.name = caption.name, 
                                              caption.text = "Catch distribution by fleet in 2016 as estimated by ICES."),
                           par.properties = base_par_prop,
                           bookmark = "CATCH_DISTRIBUTION_TABLE_CAPTION")
  tabCount <- tabCount + 1
  
  # Catch history
  draftDoc <- addParagraph(draftDoc, 
                           value = captionPot(stock.code, 
                                              type = "Table",
                                              caption.number = tabCount, 
                                              caption.name = caption.name, 
                                              caption.text = "History of commercial catch and landings; both the official and ICES estimated values are presented by area for each country participating in the fishery. All weights are in tonnes."),
                           par.properties = base_par_prop,
                           bookmark = "CATCH_HISTORY_TABLE_CAPTION")
  
  tabCount <- tabCount + 1
  if(data.category %in% c(1, 2, 3, 4)) {  
    draftDoc <- addParagraph(draftDoc,
                             value = captionPot(stock.code,
                                                type = "Table",
                                                caption.number = tabCount,
                                                caption.name = caption.name,
                                                caption.text = "Assessment summary. Weights are in tonnes."),
                             par.properties = base_par_prop,
                             bookmark = "ASSESSMENT_SUMMARY_TABLE_CAPTION")
    tabCount <- tabCount + 1
  }
  
  if(data.category %in% c(5,6))
  draftDoc <- addParagraph(draftDoc, 
                           value = summaryPot(common.name),
                           par.properties = base_par_prop,
                           bookmark = "COMMON_NAME")

  # Add tables
  # Advice basis
  draftDoc <- addFlexTable(draftDoc,
                           flextable = advice_basis_table(stock.code),
                           bookmark = "ADVICE_BASIS_TABLE")
  # Assessment basis
  draftDoc <- addFlexTable(draftDoc,
                           flextable = assessment_basis_table(stock.code, data.category,
                                                              expert.name, expert.url),
                           bookmark = "ASSESSMENT_BASIS_TABLE")
  # Advice history
  draftDoc <- addFlexTable(draftDoc,
                           flextable = advice_history_table(stock.code, data.category),
                           bookmark = "ADVICE_HISTORY_TABLE")
  # Catch options
  if(data.category %in% c(1, 2)) {
  draftDoc <- addFlexTable(draftDoc,
                           flextable = catch_options_basis_table(stock.code),
                           bookmark = "CATCH_OPTIONS_BASIS_TABLE")
  }  

  # Clean up bookmarks
  delete_tables <-  c("ADVICE_BASIS_TABLE", "ADVICE_HISTORY_TABLE",
                      "ADVICE_HISTORY_TABLE_2", "ADVICE_HISTORY_TABLE_3",
                      "ADVICE_HISTORY_TABLE_4", "ADVICE_HISTORY_TABLE_5",
                      "ADVICE_HISTORY_TABLE_6", "ASSESSMENT_BASIS_TABLE",
                      "CATCH_OPTIONS_BASIS_TABLE")
  
  lapply(delete_tables, function(x) deleteBookmark(draftDoc, x))
  

  writeDoc(doc = draftDoc,
           file = paste0(draft.url, stock.code, ".docx"))
  
cat(paste0("Check-in you new advice draft here: ",
           paste0(draft.url, stock.code, ".docx"), 
           " . \nDon't forget to add the ADG (", 
           adg.name, ") and EG (",
           expert.name,
           ") to the document properties. \n"))
  
}


stock.code <- fileList$StockCode[fileList$ExpertGroup == "HAWG" &
                                   !grepl(pattern = c("February"), fileList$PubDate)]
createDraft(stock.code[1])
