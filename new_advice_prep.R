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
stock_name <- "pok.27.3a46"

# start function here #
stock_sd <- fileList %>% 
  filter(StockKeyLabel == stock_name) 

fileName <- stock_sd %>% 
  pull(filepath)

doc <- officer::read_docx(fileName)

tabs <- docxtractr::read_docx(fileName)

content <- docx_summary(doc)
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
                                grepl("Basis of assessment and advice|Basis of the assessment and advice", text) ~ "assessmentbasis",
                                grepl("ICES advice and official landings", text) ~ "advice",
                                grepl("Catch distribution by fleet", text) ~ "catchdistribution",
                                grepl("History of commercial", text) ~ "catchhistory",
                                grepl("Assessment summary", text) ~ "assessmentsummary",
                                ### Add additional for Nephrops cat 3+ and other special cases ###
                                TRUE ~ NA_character_),
         ### if a table_name is.na match table_name to table_number ###
         table_name = ave(table_name, table_number, FUN = function(x) unique(x[!is.na(x)]))
  ) %>% 
  select(doc_index, table_number, table_name, is_header, row_id, cell_id, text)


tab_names <- tab_heads %>% 
  select(table_name, table_number) %>% 
  filter(!table_name %in% c("stocksummary")) %>% 
  distinct(.keep_all = TRUE)

table_cells <- content %>%
  filter(content_type %in% "table cell") %>%
  group_by(doc_index) %>%
  # {mutate(ungroup(.), table_number = 1 + group_indices(.))} %>%
  mutate(is_header = case_when(row_id == 1 ~ TRUE,
                               row_id != 1 ~ FALSE),
         # table_number = paste0("Table ", table_number),
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
  left_join(tab_names, by = c("table_name")) %>% 
  ungroup() %>%
  filter(table_name != "REMOVE") %>%
  select(doc_index, table_name, is_header, row_id, cell_id, text)

# 
# td <- tab_heads %>% 
#   select(table_name, table_number, caption_text = text) %>% 
#   filter(!table_name %in% c("stocksummary", "catchhistory", "assessmentsummary"))
# 
# tt <- table_cells %>% 
#   filter(is_header,
#          !table_name %in% c("stocksummary", "catchhistory", "assessmentsummary")) %>% 
#   select(doc_index, table_name) %>% 
#   distinct(.keep_all = TRUE) %>% 
#   left_join(td)
# 
# tab_heads %>% 
#   select(table_number, table_name, caption_text = text) %>% 
#   left_join(tt)
# 
# tt <- table_cells %>%
#   select(table_name, table_number, diddy = text, doc_index) %>% 
#   left_join(tab_heads)

# table <- "catchoptionsbasis"
# advice_flextable <- function(table_name) {
#   
#   
#   if(table == "Table2" &
#      all(grepl("variable|value|source|notes", tolower(table_header$text)))) {
#     table_body[, 2:ncol(table_body)] <- ""
#     table_body[, c(1, 3)] <- gsub("\\s*\\([^\\)]+\\)", " (UPDATE)", as.matrix(table_body[,c(1, 3)]))
#   }
#   
#   if(table == "Table3" &
#      "basis" %in% tolower(table_header$text)) {
#     table_body[, c(2, 4)] <- ""
#     table_body[, c(1, 3)] <- gsub("\\s*\\([^\\)]+\\)", " (UPDATE)", as.matrix(table_body[,c(1, 3)]))
#   }
#   
#   table_header <- table_header %>% 
#     spread(cell_id, text) %>% 
#     select(-table_number,
#            -is_header,
#            -row_id) %>% 
#     unlist(., use.names = FALSE)
#   
#   typology <- data.frame(col_keys = colnames(table_body),
#                          table_header)
#   
#   flextable(table_body) %>%
#     set_header_df(mapping = typology , key = "col_keys")
# }


# advice_flextable(table = "Table2")

x = doc
table <- "advice"
update_cols = NULL
update_header = FALSE
erase_cols = NULL
add_year = TRUE

## Write a function to find table and perform necessary actions based on instructions for table_name ##
table_fix <- function(x, table, 
                      update_header = c(TRUE, FALSE)[2], update_cols = NULL,
                      erase_cols = NULL, add_year = c(TRUE, FALSE)[2], 
                      messages = TRUE) {
  
  if(!table %in% unique(table_cells$table_name)) {
    stop(paste0("table_name = ", table, " is not a table cell in the table_cells tibble"))
  }
  
  tc <- table_cells %>%
    filter(table_name == table)

  for(tab_index in 1:length(unique(tc$doc_index))) {
    
    table_index <- unique(tc$doc_index)[tab_index]
    
    table_body <- tc %>% 
      filter(!is_header,
             doc_index == table_index) %>%
      ungroup %>% 
      select(-doc_index) %>%
      spread(cell_id, text) %>% 
      select(-table_name,
             -is_header,
             -row_id)
    
    table_header <- tc %>% 
      filter(is_header,
             doc_index == table_index) %>% 
      ungroup %>% 
      select(-doc_index) %>% 
      spread(cell_id, text) %>% 
      select(-table_name,
             -is_header,
             -row_id) %>% 
      unlist(., use.names = FALSE)
    
    if(!is.null(update_cols)){
      
      if(is.numeric(update_cols)) {
        update_cols_id <- update_cols
      }
      
      if(is.character(update_cols)) {
        if("all" %in% update_cols) {
          update_cols_id <- unique(tc$cell_id)
        } else {
          update_cols_id <- tc$cell_id[tc$text %in% update_cols]
        }
      }
      
      for(i in update_cols_id) {
        j = unlist(table_body[,i])
        
        ## If all values in parentheses are numeric, add a year
        if(all(grepl("\\(\\d{4}\\)", j) == TRUE)) {
          new_vals <- sprintf("(%s)", as.numeric(gsub("\\(|\\)", 
                                                      "",
                                                      stringr::str_extract(j, "\\(\\d{4}\\)"))) + 1)
          
          table_body[,i] <- stringr::str_replace_all(j, "\\(\\d{4}\\)", new_vals)
        }
        
        ## If there are references w/ characters in the parentheses, just put "(UPDATE)"
        if(all(grepl("ICES \\(\\d{4}.*\\)", j) == TRUE)) {
          table_body[,i] <- "ICES (UPDATE REF)"
        }
        
        ## If there are any F years to update
        if(any(grepl("F\\d{4}", j) == TRUE)) {
          f_vals <- sprintf("F%s", as.numeric(gsub("F", 
                                                   "",
                                                   stringr::str_extract(j, "F\\d{4}"))) + 1)
          table_body[,i] <- stringr::str_replace_all(j, "F\\d{4}", f_vals)
        }
        
        ## 
        if(any(grepl("\\d.* [T-t]onnes|\\d+\\.*\\d*%", j) == TRUE)){
          
          table_body[,i] <- stringr::str_replace_all(j, "\\d.* [T-t]onnes|\\d+\\.*\\d*%", "UPDATE VALUE")
        }
      }
    }
    
    
    if(!is.null(erase_cols)){
      
      if(is.numeric(erase_cols)) {
        erase_cols_id <- erase_cols
      } 
      
      if(is.character(erase_cols)) {
        if("all" %in% erase_cols) {
          erase_cols_id <- unique(tc$cell_id)[-1]
        } else {
          erase_cols_id <- tc$cell_id[tc$text %in% erase_cols]
        }
      }
      table_body[ , erase_cols_id] <- ""
    }
    
    
    if(add_year){
      if("year" %in% tolower(table_header)){
        new_years <- stock_sd$YearOfNextAssessment + 1:stock_sd$AssessmentFrequency
        empty_table <- table_body[1:length(new_years),]
        empty_table[] <- ""
        empty_table[, 1] <- as.character(new_years)
        table_body <- bind_rows(table_body, empty_table)
      } else {
        stop("add_year only works on the 'advice' table. Check your table")    
      }
    }
    
    
    if(update_header){
      # If all values are numeric, add a year
      table_header[grepl("\\(\\d{4}\\)", table_header)] <- sprintf("%s (%s)", gsub("\\(\\d{4}\\)", "",
                                                                                   table_header[grepl("\\(\\d{4}\\)", 
                                                                                                      table_header)]), 
                                                                   as.numeric(gsub(".*\\((.*)\\).*", "\\1", 
                                                                                   table_header[grepl("\\(\\d{4}\\)",
                                                                                                      table_header)])) + 1)
    }
    
    typology <- data.frame(col_keys = colnames(table_body),
                           table_header)
    
    ft <- flextable(table_body) %>%
      set_header_df(mapping = typology , key = "col_keys")
    
    ## Figure out where to put the table ##
    cc <- tab_heads %>% 
      filter(table_name == table)
    
    caption_index <- unique(cc$doc_index)[tab_index]
    
    caption_keyword <- tab_heads %>% 
      filter(table_name == table,
             doc_index == caption_index) %>% 
      pull(text)
    
    if(table %in% c("catchhistory", "assessmentsummary")) {
      cursor_reach(x, keyword = caption_keyword) %>% 
        cursor_forward %>% 
        body_remove      
    }
    
    if(!table %in% c("catchhistory", "assessmentsummary")){
      cursor_reach(x, keyword = caption_keyword) %>%
        cursor_forward %>% 
        body_remove %>% 
        body_add_flextable(ft)
    }
  }
    
  # update_header
  # update_cols
  # erase_cols
  # add_year
  
}
# 
# 
# 
#   # stocksummary 
#   if(table_name == "stocksummary"){
#     if(messages) cat("Please updated Stock Summary Table at sg.ices.dk and upload directly")
#   }
#   
#   # catchoptionsbasis
#   if(table_name == "catchoptionsbasis"){
#     if(messages) cat("Basis of the catch options table.\n\terasing: 'Value' and 'Notes' columns\n\tupdating: 'Variable' and 'Source' columns")
#     
#   }
#   # catchoptions      
#   if(table_name == "catchoptions"){
#     if(messages) cat("Catch options table.\n\terasing: all columns\n\tupdating: 'Variable' and 'Source' columns")
#     
#   }

doc %>% 
  table_fix(x = ., table = "catchoptionsbasis", 
            update_header = FALSE, update_cols = c("Variable", "Source"), 
            erase_cols = c("Value", "Notes"),
            add_year = FALSE, 
            messages = FALSE) %>% 
  table_fix(table = "catchoptions", 
            update_header = TRUE, update_cols = "Basis", 
            erase_cols = "all",
            add_year = FALSE, 
            messages = FALSE) %>% 
  table_fix(table = "advicebasis") %>% 
  # table_fix(table = "msyranges") %>% 
  table_fix(table = "referencepoints") %>% 
  # table_fix(table = "assessmentbasis") %>% 
  table_fix(table = "advice") %>% 
  table_fix(table = "catchdistribution") %>% 
  table_fix(table = "catchhistory") %>% 
  table_fix(table = "assessmentsummary")
  
print(doc, target = "cursor.docx")

# 
# 
#   cursor_reach(keyword = sprintf("%s. The basis for the catch options.", "Cod in Subarea 4, Division 7.d, and Subdivision 20")) %>% 
#   cursor_forward %>% 
#   body_remove %>% 
#   body_add_flextable(advice_flextable("Table2")) %>% 
#   cursor_reach(keyword = sprintf("%s. Annual catch options.", stock_sd$CaptionName)) %>% 
#   cursor_forward %>% 
#   body_remove %>% 
#   body_add_flextable(advice_flextable("Table3")) %>% 
#   cursor_reach(keyword = sprintf("%s. The basis of the advice.", stock_sd$CaptionName)) %>% 
#   cursor_forward %>% 
#   body_remove %>% 
#   body_add_flextable(advice_flextable("Table4")) %>% 
#   cursor_reach(keyword = sprintf("%s. Reference points, values, and their technical basis.", stock_sd$CaptionName)) %>% 
#   cursor_forward %>% 
#   body_remove %>% 
#   body_add_flextable(advice_flextable("Table5")) %>% 
#   cursor_reach(keyword = sprintf("%s. Basis of assessment and advice.", stock_sd$CaptionName)) %>% 
#   cursor_forward %>% 
#   body_remove %>% 
#   body_add_flextable(advice_flextable("Table6")) %>% 
#   # This table just needs an extra row added... sometimes there are multiples, so need to be clever
#   cursor_reach(keyword = sprintf("%s. ICES advice and official landings.", stock_sd$CaptionName)) %>% 
#   cursor_forward %>% 
#   body_remove %>% 
#   body_add_flextable(advice_flextable("Table7")) %>% 
#   cursor_reach(keyword = sprintf("%s. Catch distribution", stock_sd$CaptionName)) %>% 
#   body_add_par(value =  sprintf("%s. Catch distribution by fleet in 2017 as estimated by ICES.",
#                                 stock_sd$CaptionName), pos = "on") %>% 
#   cursor_forward %>% 
#   body_remove %>% 
#   body_add_flextable(advice_flextable("Table8")) %>% 
#   cursor_reach(keyword = "Table ") %>%
#   cursor_forward %>% 
#   body_remove %>% 
#   cursor_reach(keyword = "Table ") %>%
#   cursor_forward %>% 
#   body_remove
#   
