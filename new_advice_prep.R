rm(list = ls())
#############################
# Install and load packages #
#############################

# library(docxtractr)
library(flextable)
library(officer)
# library(jsonlite)
library(dplyr)
library(tidyr)

# Note: You must log in to SharePoint and have this drive mapped
sharePoint <- "//community.ices.dk/DavWWWRoot/"
if(!dir.exists(paste0(sharePoint, "Advice/Advice2017/"))) {
  stop("Note: You must be on the ICES network and log in to SharePoint and have this drive mapped")
}

## ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ##
## Download and prepare the stock information data ##
## ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ##

rawsd <- jsonlite::fromJSON("http://sd.ices.dk/services/odata3/StockListDWs3")$value %>% 
  filter(ActiveYear == 2017,
         YearOfNextAssessment == 2018) %>% 
  mutate(CaptionName = gsub("\\s*\\([^\\)]+\\)", "", as.character(StockKeyDescription)),
         AdviceReleaseDate = as.Date(AdviceReleaseDate, format = "%d/%m/%y"),
         PubDate = format(AdviceReleaseDate, "%e %B %Y"),
         ExpertURL = paste0("http://www.ices.dk/community/groups/Pages/", ExpertGroup, ".aspx"),
         DataCategory = gsub("\\..*$", "", DataCategory))


## ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ##
## Find most recent released advice on SharePoint ##
## ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ##

advice_file_finder <- function(year){
  
  advice <- list.files(sprintf("%sAdvice/Advice%s/", sharePoint, year))
  
  if(year == 2015){
    folderNames = c("BalticSea", "BarentsSea", "BayOfBiscay", 
                    "CelticSea", "FaroePlateau", "Iceland",
                    "NASalmon","NorthSea", "Widely")
  }
  if(year == 2016){
    folderNames = c("BalticSea", "BarentsSea", "Biscay", 
                    "CelticSea", "Faroes", "Iceland",
                    "NorthSea", "Salmon", "Widely")
  }
  
  if(year == 2017){
    folderNames = c("BalticSea", "BarentsSea", "BayOfBiscay", 
                    "CelticSea", "Faroes", "Iceland",
                    "NorthSea", "Salmon", "Widely")
  }
  
  adviceList <- lapply(advice[advice %in% folderNames],
                       function(x) list.files(sprintf("%sAdvice/Advice%s/%s/Released_Advice", sharePoint, year, x)))
  
  names(adviceList) <- folderNames
  fileList <- do.call("rbind", lapply(adviceList,
                                      data.frame, 
                                      stringsAsFactors = FALSE))
  colnames(fileList) <- "StockCode"
  fileList$filepath <- sprintf("%sAdvice/Advice%s/%s/Released_Advice/%s", 
                               sharePoint,
                               year,
                               gsub("\\..*", "", row.names(fileList)),
                               fileList$StockCode)
  fileList$StockCode <- tolower(gsub("\\.docx*", "", fileList$StockCode))
  return(fileList)
}

fileList <- bind_rows(
  rawsd %>%
    filter(YearOfLastAssessment == 2015) %>% 
      left_join(advice_file_finder(2015), by = c("PreviousStockKeyLabel" = "StockCode")),
  rawsd %>%  
    filter(YearOfLastAssessment == 2016) %>% 
    left_join(advice_file_finder(2016), by = c("PreviousStockKeyLabel" = "StockCode")),
  rawsd %>%
    filter(YearOfLastAssessment == 2017) %>% 
    left_join(advice_file_finder(2017), by = c("StockKeyLabel" = "StockCode"))
  ) %>% 
  mutate(URL = ifelse(is.na(filepath),
                      NA,
                      paste0(gsub("//community.ices.dk/DavWWWRoot/", "https://community.ices.dk/", filepath), "?Web=1")))


## ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ##
## Identify the tables to keep and clean ##
## ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ##
stock_name <- "cod.27.47d20"

# start function here #
stock_sd <- fileList %>% 
  filter(StockKeyLabel == stock_name) 

fileName <- stock_sd %>% 
  pull(filepath)

doc <- officer::read_docx(fileName)

tabs <- docxtractr::read_docx(fileName)

# content <- docx_summary(doc)
columnNames <- lapply(seq(1: docxtractr::docx_tbl_count(tabs)),
                      function(x) colnames(docxtractr::docx_extract_tbl(tabs,
                                                                        tbl_number = x, header = TRUE)))
## 
##  
##

tab_heads <- content %>% 
  filter(content_type %in% "paragraph",
         grepl("Table ", text)) %>% 
  mutate(table_number = gsub("([0-9]+).*$", "\\1", text),
         table_name = case_when(grepl("State of the stock and fishery relative to reference points", text) ~ "stocksummary",
                                grepl("The basis for the catch options", text) ~ "catchoptionsbasis",
                                grepl("Annual catch options", text) ~ "catchoptions",
                                grepl("The basis of the advice", text) ~ "advicebasis",
                                grepl("Reference points, values, and their technical basis", text) ~ "referencepoints",
                                grepl("Basis of assessment and advice", text) ~ "assessmentbasis",
                                grepl("ICES advice and official landings", text) ~ "advice",
                                grepl("Catch distribution by fleet", text) ~ "catchdistribution",
                                grepl("History of commercial catch", text) ~ "catchhistory",
                                grepl("Assessment summary", text) ~ "assessmentsummary",
                                ### Add additional for Nephrops cat 3+ and other special cases ###
                                TRUE ~ NA_character_),
         ### if a table_name is.na match table_name to table_number ###
         table_name = ave(table_name, table_number, FUN = function(x) unique(x[!is.na(x)]))
  ) %>% 
  select(table_number, table_name, is_header, row_id, cell_id, text)


tab_names <- tab_heads %>% 
  select(table_name, table_number) %>% 
  filter(!table_name %in% c("stocksummary"))

table_cells <- content %>%
  filter(content_type %in% "table cell") %>%
  group_by(doc_index) %>%
  # {mutate(ungroup(.), table_number = 1 + group_indices(.))} %>%
  mutate(is_header = case_when(row_id == 1 ~ TRUE,
                               row_id != 1 ~ FALSE),
         # table_number = paste0("Table", table_number),
         table_name = case_when(is_header == TRUE & grepl("^variable$", tolower(text)) ~ "catchoptionsbasis",
                                is_header == TRUE & grepl("^basis$", tolower(text)) ~ "catchoptions",
                                is_header == TRUE & grepl("^advice basis$", tolower(text)) ~ "advicebasis",
                                is_header == TRUE & grepl("^description$", tolower(text)) ~ "msyranges",
                                is_header == TRUE & grepl("^framework$", tolower(text)) ~ "referencepoints",
                                is_header == TRUE & grepl("^ices stock data category$", tolower(text)) ~ "assessmentbasis",
                                is_header == TRUE & grepl("^ices advice$", tolower(text)) ~ "advice",
                                is_header == TRUE & grepl("^catch \\(\\d{4}\\)$", tolower(text)) ~ "catchdistribution",
                                TRUE ~ NA_character_),
         table_name = case_when(is_header == TRUE &  ave(is.na(table_name), doc_index, FUN = all) ~ "REMOVE",
                                TRUE ~ table_name),
         table_name = ave(table_name, doc_index, FUN = function(x) unique(x[!is.na(x)]))) %>%
  full_join(tab_names, by = "table_name") %>% 
  ungroup() %>% 
  filter(table_name != "REMOVE") %>%
  select(table_number, table_name, is_header, row_id, cell_id, text)


table <- "catchoptionsbasis"
advice_flextable <- function(table_name) {
  
  
  if(table == "Table2" &
     all(grepl("variable|value|source|notes", tolower(table_header$text)))) {
    table_body[, 2:ncol(table_body)] <- ""
    table_body[, c(1, 3)] <- gsub("\\s*\\([^\\)]+\\)", " (UPDATE)", as.matrix(table_body[,c(1, 3)]))
  }
  
  if(table == "Table3" &
     "basis" %in% tolower(table_header$text)) {
    table_body[, c(2, 4)] <- ""
    table_body[, c(1, 3)] <- gsub("\\s*\\([^\\)]+\\)", " (UPDATE)", as.matrix(table_body[,c(1, 3)]))
  }
  
  table_header <- table_header %>% 
    spread(cell_id, text) %>% 
    select(-table_number,
           -is_header,
           -row_id) %>% 
    unlist(., use.names = FALSE)
  
  typology <- data.frame(col_keys = colnames(table_body),
                         table_header)
  
  flextable(table_body) %>%
    set_header_df(mapping = typology , key = "col_keys")
}


advice_flextable(table = "Table2")


table <- "catchoptions"
## Write a function to find table and perform necessary actions based on instructions for table_name ##
table_fix <- function(x, table, update_header = c(TRUE, FALSE)[2], 
                      update_rows = NULL, erase_rows = NULL, 
                      add_row = c(TRUE, FALSE)[2], messages = TRUE) {

  if(!table %in% unique(table_cells$table_name)) {
    stop(paste0("table_name = ", table, " is not a table cell in the table_cells tibble"))
  }
  
  tc <- table_cells %>%
    filter(table_name == table)
  
  table_body <- tc %>% 
    filter(!is_header) %>% 
    spread(cell_id, text) %>% 
    select(-table_number,
           -table_name,
           -is_header,
           -row_id)
  
  if(!is.null(update_rows)){
    update_rows_id <- tc$cell_id[tc$text %in% update_rows]
    
    for(i in update_rows_id) {
      j = unlist(table_body[,i])
      
      # If all values are numeric, add a year
      if(all(grepl("\\(\\d{4}\\)", j) == TRUE)) {
        table_body[,i] <- paste0(gsub("\\(\\d{4}\\)", "", j), 
                                 paste0("(", as.numeric(gsub(".*\\((.*)\\).*", "\\1", j)) + 1,
                                        ")"))
      }
      ## If there are characters in the parentheses, just put "(UPDATE)"
      if(all(grepl("\\(\\d{4}\\)", j) == FALSE)) {
        table_body[,i] <- paste0(gsub("\\(.*\\)", "", j), "(UPDATE)")        
      }
    }
  }
  
  if(!is.null(erase_rows)){
    erase_rows_id <- tc$cell_id[tc$text %in% erase_rows]
    table_body[,erase_rows_id] <- ""
  }
  
  if(add_row){
    table_body
  }
  
  table_header <- tc %>% 
    filter(is_header) %>% 
    spread(cell_id, text) %>% 
    select(-table_number,
           -table_name,
           -is_header,
           -row_id) %>% 
    unlist(., use.names = FALSE)

  if(update_header){
    
    # If all values are numeric, add a year
    table_header[grepl("\\(\\d{4}\\)", table_header)] <-sprintf("%s (%s)", gsub("\\(\\d{4}\\)", "",
                                                                                table_header[grepl("\\(\\d{4}\\)", 
                                                                                                   table_header)]), 
                                                                as.numeric(gsub(".*\\((.*)\\).*", "\\1", 
                                                                                table_header[grepl("\\(\\d{4}\\)",
                                                                                                   table_header)])) + 1)
    }
    
  typology <- data.frame(col_keys = colnames(table_body),
                         table_header)
  
  flextable(table_body) %>%
    set_header_df(mapping = typology , key = "col_keys")
  
}


  # stocksummary 
  if(table_name == "stocksummary"){
    if(messages) cat("Please updated Stock Summary Table at sg.ices.dk and upload directly")
  }
  
  # catchoptionsbasis
  if(table_name == "catchoptionsbasis"){
    if(messages) cat("Basis of the catch options table.\n\terasing: 'Value' and 'Notes' columns\n\tupdating: 'Variable' and 'Source' columns")
    
  }
  # catchoptions      
  if(table_name == "catchoptions"){
    if(messages) cat("Catch options table.\n\terasing: all columns\n\tupdating: 'Variable' and 'Source' columns")
    
  }
  
  # advicebasis       
  # referencepoints   
  # assessmentbasis  
  # advice
  # catchdistribution
  # catchhistory
  # assessmentsummary 
  
    cursor_reach(x, keyword = table_name) %>% 
    cursor_forward %>% 
    body_remove
}




unique(table_cells$table_number)

doc %>% 
  cursor_reach(keyword = sprintf("%s. The basis for the catch options.", "Cod in Subarea 4, Division 7.d, and Subdivision 20")) %>% 
  cursor_forward %>% 
  body_remove %>% 
  body_add_flextable(advice_flextable("Table2")) %>% 
  cursor_reach(keyword = sprintf("%s. Annual catch options.", stock_sd$CaptionName)) %>% 
  cursor_forward %>% 
  body_remove %>% 
  body_add_flextable(advice_flextable("Table3")) %>% 
  cursor_reach(keyword = sprintf("%s. The basis of the advice.", stock_sd$CaptionName)) %>% 
  cursor_forward %>% 
  body_remove %>% 
  body_add_flextable(advice_flextable("Table4")) %>% 
  cursor_reach(keyword = sprintf("%s. Reference points, values, and their technical basis.", stock_sd$CaptionName)) %>% 
  cursor_forward %>% 
  body_remove %>% 
  body_add_flextable(advice_flextable("Table5")) %>% 
  cursor_reach(keyword = sprintf("%s. Basis of assessment and advice.", stock_sd$CaptionName)) %>% 
  cursor_forward %>% 
  body_remove %>% 
  body_add_flextable(advice_flextable("Table6")) %>% 
  # This table just needs an extra row added... sometimes there are multiples, so need to be clever
  cursor_reach(keyword = sprintf("%s. ICES advice and official landings.", stock_sd$CaptionName)) %>% 
  cursor_forward %>% 
  body_remove %>% 
  body_add_flextable(advice_flextable("Table7")) %>% 
  cursor_reach(keyword = sprintf("%s. Catch distribution", stock_sd$CaptionName)) %>% 
  body_add_par(value =  sprintf("%s. Catch distribution by fleet in 2017 as estimated by ICES.",
                                stock_sd$CaptionName), pos = "on") %>% 
  cursor_forward %>% 
  body_remove %>% 
  body_add_flextable(advice_flextable("Table8")) %>% 
  cursor_reach(keyword = "Table ") %>%
  cursor_forward %>% 
  body_remove %>% 
  cursor_reach(keyword = "Table ") %>%
  cursor_forward %>% 
  body_remove
  

    

print(doc, target = "cursor.docx")

  # slip_in_text("Bananna", pos = "after", style = "Default Paragraph Font") 
table = "Table3"


docx_bookmarks(doc)

content <- docx_summary(doc)

doc <- docxtractr::read_docx(advice_sheet)
docxtractr::docx_tbl_count(doc)

flextable(data.frame(docxtractr::docx_extract_tbl(tabs, tbl_number = 5,header = TRUE, trim = TRUE)))


