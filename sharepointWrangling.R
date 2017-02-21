rm(list = ls())
#############################
# Install and load packages #
#############################
library(readxl) # not necessary once SLD is up and running
library(tidyverse)
library(docxtractr)
library(ReporteRs)
library(jsonlite)
library(stringr)
library(httr)

# Note: You must log in to SharePoint and have this drive mapped
sharePoint <- "//community.ices.dk@SSL/DavWWWRoot/"
if(!dir.exists(paste0(sharePoint, "Advice/Advice2017/"))) {
  stop("Note: You must be on the ICES network and log in to SharePoint and have this drive mapped")
}

#############################################
# Organizing stock list and section numbers #
#############################################
# codeList <- readxl::read_excel(paste0(sharePoint, "Advice/Advice2017/Advice documents/ICES Stock code reference list.xlsx"))
# codeList <- codeList[,c(1:2)]
# colnames(codeList) <- c("StockCode", "OldCode")

rawsd <- fromJSON("http://sd.ices.dk/services/odata3/StockListDWs3?$filter=ActiveYear%20eq%202017")$value

# ecoregions <- colnames(rawsd)[grepl("^.+(Ecoregion)$", colnames(rawsd))]
# dots <- lapply(c("StockCode", "Description", "SpeciesCommonName",
#                  #SPECIES NAME,
#                  #OLDCODE,
#                  "SectionNumber", "DataCategory", ecoregions),
#                as.symbol)
# 
# all_ecoregions <- data.frame(ecoregion = ecoregions,
#                              stringsAsFactors = FALSE)
sl_ecoregion <- rawsd %>%
  select(StockCode, OldStockCode, Description, SpeciesCommonName,
         SpeciesScientificName, 
         # SectionNumber, 
         DataCategory, 
         AdviceHeading = EcoRegion,
         AdviceReleaseDate) %>%
  mutate(count = str_count(AdviceHeading, ",") + 1,
         AdviceReleaseDate = as.Date(AdviceReleaseDate, format = "%d/%m/%y"),
         PubDate = format(AdviceReleaseDate, "%e %B %Y"),
         SectionNumber = "TO BE ADDED")
# 



# %>%
  # select_(.dots = dots) %>%
  # gather(ecoregion, value, -StockCode, -Description, -DataCategory, #-SpeciesName, #-OldCode
  #        -SpeciesCommonName, -SectionNumber) %>%
  # mutate(ecoregion = recode(ecoregion, "ArcticOceanEcoregion" = "Arctic Ocean"),
  #        ecoregion = recode(ecoregion, "AzoresEcoregion" = "Azores"),
  #        ecoregion = recode(ecoregion, "BayofBiscayandtheIberianCoastEcoregion" = "Bay of Biscay and the Iberian Coast"),
  #        ecoregion = recode(ecoregion, "BarentsSeaEcoregion" = "Barents Sea"),
  #        ecoregion = recode(ecoregion, "BalticSeaEcoregion" = "Baltic Sea"),
  #        ecoregion = recode(ecoregion, "CelticSeasEcoregion" = "Celtic Seas"),
  #        ecoregion = recode(ecoregion, "FaroesEcoregion" = "Faroes"),
  #        ecoregion = recode(ecoregion, "GreenlandSeaEcoregion" = "Greenland Sea"),
  #        ecoregion = recode(ecoregion, "IcelandSeaEcoregion" = "Iceland Sea"),
  #        ecoregion = recode(ecoregion, "GreaterNorthSeaEcoregion" = "Greater North Sea"),
  #        ecoregion = recode(ecoregion, "OceanicNortheastAtlanticEcoregion" = "Oceanic Northeast Atlantic"),
  #        ecoregion = recode(ecoregion, "NorwegianSeaEcoregion" = "Norwegian Sea")) %>%
  # group_by(StockCode) %>%
  # filter(!is.na(value)) %>%
  # mutate(count = n()) %>%
  # full_join(all_ecoregions, by = "ecoregion") %>%
  # spread(ecoregion, ecoregion) %>%
  # filter(!is.na(value)) %>%
  # unite(AdviceHeading, 
  #       ArcticOceanEcoregion:OceanicNortheastAtlanticEcoregion,
  #       sep = ", ", remove = FALSE)

sl <- sl_ecoregion %>%
  mutate(Description = gsub(" $","", Description, perl = TRUE), # remove trailing spaces
         DataCategory = gsub("\\..*$", "", DataCategory), # Just use the first digit
         # AdviceHeading = gsub("NA, ", "", AdviceHeading), # remove NAs
         # AdviceHeading = gsub(", NA", "", AdviceHeading), # remove NAs
         AdviceHeading = gsub("Ecoregion", "", AdviceHeading), # remove Ecoregion
         AdviceHeading = gsub(",(?=[^,]+$)", " and", AdviceHeading, perl = TRUE), # remove last comma and replace with "and"
         AdviceHeading = ifelse(count == 1,
                                 paste0(AdviceHeading, " Ecoregion"),
                                 paste0(AdviceHeading, " Ecoregions")),
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
  select(StockCode, Description, SectionNumber,
         PubDate, DataCategory, AdviceHeading, CaptionName, 
         SpeciesScientificName, OldStockCode,
         SpeciesCommonName)

# %>%
#   ungroup() %>%
#   left_join(codeList, by = "StockCode")

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

####################################
# Wrap it up and write some drafts #
####################################

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

#########################################################
# Create Draft given the Stock List Database and tables #
#########################################################
#

# draftDir <- paste0(sharePoint, "/admin/AdvisoryProgramme/Personal folders/Scott/Advice_Template_Testing/drafts/")
# draftDir <-"~/git/ices-dk/AdviceTemplate/"

createDraft <- function(stock.code, output_dir = draftDir) {

  stock.code <- fileList$StockCode[1]
  pub.date <- fileList$PubDate[fileList$StockCode %in% stock.code]
  ecoregion.name <- fileList$AdviceHeading[fileList$StockCode %in% stock.code]
  section.number <- fileList$SectionNumber[fileList$StockCode %in% stock.code]
  stock.name <- fileList$Description[fileList$StockCode %in% stock.code]
  caption.name <-  fileList$CaptionName[fileList$StockCode %in% stock.code]
  common.name <- fileList$SpeciesCommonName[fileList$StockCode %in% stock.code]
  data.category <- fileList$DataCategory[fileList$StockCode %in% stock.code]
  
  ########################
  # Load Table functions #
  ########################
  source("~/git/ices-dk/AdviceTemplate/table_functions.R")
  
 # Load up the proper template
  if(sl$DataCategory[sl$StockCode %in% stock.code] %in% c(1,2)) {
    template <- paste0(sharePoint, "Advice/Advice2017/Advice documents/Draft advice 2017/advice_template_2017_cat12.docx")
  }
  if(sl$DataCategory[sl$StockCode %in% stock.code] %in% c(3,4)) {
    template <- paste0(sharePoint, "Advice/Advice2017/Advice documents/Draft advice 2017/advice_template_2017_cat34.docx")  
  }
  if(sl$DataCategory[sl$StockCode %in% stock.code] %in% c(5,6)) {
    template <- paste0(sharePoint, "Advice/Advice2017/Advice documents/Draft advice 2017/advice_template_2017_cat56.docx")
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
                           value = headingPot(stock.code, section.number, stock.name),
                           par.properties = base_par_prop,
                           bookmark = "TITLE_HEADING")

  draftDoc <- addParagraph(draftDoc, 
                           value = footerPot(stock.code, section.number),
                           par.properties = base_par_prop,
                           bookmark = "FIRST_PAGE_FOOTER")
  
  draftDoc <- addParagraph(draftDoc, 
                           value = headerPot(stock.code, pub.date),
                           par.properties = base_par_prop,
                           bookmark = "SECOND_PAGE_HEADER")
  
  draftDoc <- addParagraph(draftDoc, 
                           value = footerPot(stock.code, section.number),
                           par.properties = base_par_prop,
                           bookmark = "SECOND_PAGE_FOOTER")
  # SAG figures
  draftDoc <- addParagraph(draftDoc, 
                           value = captionPot(stock.code, 
                                              type = "Figure",
                                              section.number = section.number,
                                              caption.number = 1, 
                                              caption.name = caption.name,
                                              caption.text = "Summary of the stock assessment."),
                           par.properties = base_par_prop,
                           bookmark = "STOCK_SUMMARY_FIGURE_CAPTION")
  # Tick-mark table
  draftDoc <- addParagraph(draftDoc, 
                           value = captionPot(stock.code, 
                                             type = "Table",
                                             section.number = section.number,
                                             caption.number = 1, 
                                             caption.name = caption.name,
                                             caption.text = "State of the stock and fishery relative to reference points."),
                           par.properties = base_par_prop,
                           bookmark = "STOCK_STATUS_TABLE_CAPTION")
  # Catch options table
  draftDoc <- addParagraph(draftDoc, 
                           value = captionPot(stock.code, 
                                              type = "Table",
                                              section.number = section.number,
                                              caption.number = 2, 
                                              caption.name = caption.name,
                                              caption.text = "For stocks in ICES categories 3-6, one catch option is provided"),
                           par.properties = base_par_prop,
                           bookmark = "CATCH_OPTIONS_BASIS_TABLE_CAPTION")
  # Basis of the advice
  draftDoc <- addParagraph(draftDoc, 
                           value = captionPot(stock.code, 
                                              type = "Table",
                                              section.number = section.number,
                                              caption.number = 3, 
                                              caption.name = caption.name,
                                              caption.text = "The basis of the advice."),
                           par.properties = base_par_prop,
                           bookmark = "ADVICE_BASIS_TABLE_CAPTION")
  # Basis of the assessment
  draftDoc <- addParagraph(draftDoc, 
                           value = captionPot(stock.code, 
                                              type = "Table",
                                              section.number = section.number,
                                              caption.number = 4, 
                                              caption.name = caption.name,
                                              caption.text = "The basis of the assessment"),
                           par.properties = base_par_prop,
                           bookmark = "ASSESSMENT_BASIS_TABLE_CAPTION")
  
  # Advice history
  draftDoc <- addParagraph(draftDoc,
                           value = captionPot(stock.code,
                                              type = "Table",
                                              section.number = section.number,
                                              caption.number = 5,
                                              caption.name = caption.name,
                                              caption.text = "ICES advice and official landings. All weights are in thousand tonnes."),
                           par.properties = base_par_prop,
                           bookmark = "ADVICE_HISTORY_TABLE_CAPTION")
  
  # Catch Distribution history
  draftDoc <- addParagraph(draftDoc, 
                           value = captionPot(stock.code, 
                                              type = "Table",
                                              section.number = section.number,
                                              caption.number = 6, 
                                              caption.name = caption.name, 
                                              caption.text = "Catch distribution by fleet in 2017 as estimated by ICES."),
                           par.properties = base_par_prop,
                           bookmark = "CATCH_DISTRIBUTION_TABLE_CAPTION")
  
  # Catch Distribution history
  draftDoc <- addParagraph(draftDoc, 
                           value = captionPot(stock.code, 
                                              type = "Table",
                                              section.number = section.number,
                                              caption.number = 7, 
                                              caption.name = caption.name, 
                                              caption.text = "History of commercial catch and landings; both the official and ICES estimated values are presented by area for each country participating in the fishery. All weights are in tonnes."),
                           par.properties = base_par_prop,
                           bookmark = "CATCH_HISTORY_TABLE_CAPTION")
  
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
                           flextable = assessment_basis_table(stock.code, data.category),
                           bookmark = "ASSESSMENT_BASIS_TABLE")
  # Advice history
  draftDoc <- addFlexTable(draftDoc,
                           flextable = advice_history_table(stock.code, data.category),
                           bookmark = "ADVICE_HISTORY_TABLE")
  
  # Clean up bookmarks
  delete_tables <-  c("ADVICE_BASIS_TABLE", "ADVICE_HISTORY_TABLE",
                      "ADVICE_HISTORY_TABLE_2", "ADVICE_HISTORY_TABLE_3",
                      "ADVICE_HISTORY_TABLE_4", "ADVICE_HISTORY_TABLE_5",
                      "ADVICE_HISTORY_TABLE_6", "ASSESSMENT_BASIS_TABLE")
  
  lapply(delete_tables, function(x) deleteBookmark(draftDoc, x))
  

  writeDoc(doc = draftDoc,
           file = paste0("~/git/ices-dk/AdviceTemplate/", stock.code, ".docx"))
  
# }
  # adviceHistoryList <- advice_history_table(stock.code, data.category)
  # adviceHistoryNames <- grep("T", names(adviceHistoryList), value = TRUE)
  # 
  # for(tn in 1:length(adviceHistoryNames)) {
  #   bm <- ifelse(tn >= 2,
  #                paste0("ADVICE_HISTORY_TABLE_", tn),
  #                "ADVICE_HISTORY_TABLE")  
  #   draftDoc <- addFlexTable(draftDoc,
  #                            flextable = adviceHistoryList[[tn]],
  #                            bookmark = bm)
  # }
  
  # 
  # assessmentBasisList <- assessment_basis_table(stock.code, data.category)
  # assessmentBasisNames <- names(assessmentBasisList)
  # 
  # for(tn in 1:length(assessmentBasisNames)) {
  #   bm <- ifelse(tn >= 2,
  #                paste0("ADVICE_HISTORY_TABLE_", tn),
  #                "ADVICE_HISTORY_TABLE")
  # draftDoc <- addFlexTable(draftDoc,
  #                          flextable = assessmentBasisList[[tn]],
  #                          bookmark = bm)
  # }
  # 
  
  
  if(data.category %in% c(1,2)) {
    
    draftDoc <- addParagraph(draftDoc, 
                             value = captionPot(stock.code,
                                                section.number = section.number,
                                                caption.number = 3,
                                                type = "Table",
                                                caption.name = caption.name,
                                                caption.text = "CATCH OPTIONS TABLE"),
                             par.properties = base_par_prop,
                             bookmark = "CATCH_OPTIONS_BASIS_TABLE_CAPTION")
    
    draftDoc <- addFlexTable(draftDoc,
                 flextable = catch_options_basis_table(stock.code),
                 bookmark = "CATCH_OPTIONS_BASIS_TABLE")
  }
  
  # writeDoc(doc = draftDoc,
  #          file = "~/git/ices-dk/AdviceTemplate/tester_4.docx")
  # 

}
