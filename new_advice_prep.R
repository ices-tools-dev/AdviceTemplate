#############################
# Install and load packages #
#############################

library(flextable)
library(officer)
library(dplyr)
library(tidyr)


get_filelist <- function(year = 2017) {
  
  # Note: You must log in to SharePoint and have this drive mapped
  sharePoint <- "//community.ices.dk/DavWWWRoot/"
  
  if(!dir.exists(sprintf("%sAdvice/Advice%s/", sharePoint, year))) {
    stop("Note: You must be on the ICES network and have sharepoint mapped to a local drive.")
  }
  
  ## ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ##
  ## Download and prepare the stock information data ##
  ## ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ##
  
  rawsd <- jsonlite::fromJSON("http://sd.ices.dk/services/odata3/StockListDWs3")$value %>% 
    filter(ActiveYear == year,
           YearOfNextAssessment == year + 1) %>% ## This should be fixed when ActiveYear is updated in SID 
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
    
    ### Hopefully, future folderNames are consistent, if not, map by hand, as above.
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
  ## Repeat for additional years
  ) %>% 
  mutate(URL = ifelse(is.na(filepath),
                      NA,
                      paste0(gsub("//community.ices.dk/DavWWWRoot/", 
                                  "https://community.ices.dk/", 
                                  filepath), 
                             "?Web=1")))

return(fileList)
} # Close get_filelist

fileList <- get_filelist(2017)
stock_name <- fileList$StockKeyLabel[grepl("HAWG", fileList$ExpertGroup)][11]

format_advice(stock_name, path_name = "~/") 

format_advice <- function(stock_name, path_name = "/") {

  ## ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ##
  ## Identify the tables to keep and clean ##
  ## ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ##
  # stock_name <- "cod.27.47d20"
  
  ## Start function here by grabbing info for a stock
  stock_sd <- fileList %>% 
    filter(StockKeyLabel == stock_name) 
  
  ## Identify the file 
  fileName <- stock_sd %>% 
    pull(filepath)
  
  ## Grab the last advice
  doc <- officer::read_docx(fileName)
  
  ## Pull out a data.frame of text (content[]) and tables (both tabs[] and content[])
  tabs <- docxtractr::read_docx(fileName)
  content <- officer::docx_summary(doc)
  
  ## ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ##
  ## Define tables according to to header and caption ##  
  ## ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ##
  
  ## Get all tables by searching for table cells and get the header row
  table_header_name <- content %>% 
    filter(content_type %in% "table cell",
           row_id == 1) %>%
    select(-dplyr::one_of("content_type", "style_name", "level", 
                          "num_id", "row_id", "is_header", "col_span", "row_span")) %>% 
    spread(cell_id, text) %>% 
    mutate(header_name = case_when(grepl("^[V-v]ariable", `1`) &
                                     grepl("^[V-v]alue", `2`) ~ "catchoptionsbasis",
                                   grepl("^[I-i]ndex", `1`) ~ "catchoptionsbasis",
                                   grepl("^[B-b]asis", `1`) ~ "catchoptions",
                                   grepl("^[A-a]dvice\\sbasis", `1`) ~ "advicebasis",
                                   grepl("^[D-d]escription", `1`) &
                                     grepl("^[V-v]alue", `2`) ~ "ranges",
                                   grepl("^[F-f]ramework", `1`) ~ "referencepoints",
                                   grepl("^ICES\\sstock\\sdata\\scategory", `1`) ~ "assessmentbasis",
                                   grepl("^[Y-y]ear", `1`) &
                                     grepl("^ICES\\sadvice", `2`) ~ "advice",
                                   grepl("^[C-c]atch", `1`) &
                                     grepl("^[W-w]anted\\s[C-c]atch", `2`) ~ "catchdistribution",
                                   grepl("^[C-c]atch", `1`) &
                                     grepl("^[L-l]andings", `2`) ~ "catchdistribution",
                                   grepl("^[Y-year]", `1`) ~ "assessmentsummary",
                                   TRUE ~ "other"),
           caption_index = doc_index - 1) %>% 
    select(doc_index,
           caption_index,
           header_name)
  
  ## Stock summary table is inserted as an image, so it won't be found searching for table cells,
  ## so it is added using the caption label
  
  table_1 <- content %>% 
    filter(content_type %in% "paragraph",
           grepl("state of the stock and fishery", tolower(text))) %>% 
    mutate(header_name = "stocksummary",
           caption_index = doc_index,
           doc_index = caption_index + 1) %>% 
    select(doc_index,
           caption_index,
           header_name)
  
  table_header_name <- 
    bind_rows(table_header_name,
              table_1) %>% 
    arrange(doc_index)
  
  ## Add caption text to the data.frame based on the doc_index
  table_header_name$text <- content$text[content$doc_index %in% table_header_name$caption_index]
  
  ## Grab tibble of table names - to link with the table header info, below
  tab_heads <- table_header_name %>% 
    select(doc_index,
           table_name = header_name, 
           text)
  
  tab_names <- tab_heads %>% 
    select(table_name) %>% 
    filter(!table_name %in% c("stocksummary")) %>% 
    distinct(.keep_all = TRUE)
  
  ## Holds all the table information (headers, values, ugly tables... everything)
  table_cells <- content %>%
    filter(content_type %in% "table cell") %>%
    group_by(doc_index) %>%
    mutate(is_header = case_when(row_id == 1 ~ TRUE,
                                 row_id != 1 ~ FALSE),
           table_name = case_when(is_header == TRUE & grepl("^variable$", tolower(text)) ~ "catchoptionsbasis",
                                  is_header == TRUE & grepl("^index\\sa", tolower(text)) ~ "catchoptionsbasis",
                                  is_header == TRUE & grepl("^basis$", tolower(text)) ~ "catchoptions",
                                  is_header == TRUE & grepl("^advice basis$", tolower(text)) ~ "advicebasis",
                                  is_header == TRUE & grepl("^description$", tolower(text)) ~ "ranges",
                                  is_header == TRUE & grepl("^framework$", tolower(text)) ~ "referencepoints",
                                  is_header == TRUE & grepl("^ices stock data category$", tolower(text)) ~ "assessmentbasis",
                                  is_header == TRUE & grepl("^ices advice$", tolower(text)) ~ "advice",
                                  is_header == TRUE & grepl("^catch \\(\\d{4}\\)$", tolower(text)) ~ "catchdistribution",
                                  ### Add additional for Nephrops cat 3+ and other special cases ###
                                  TRUE ~ NA_character_),
           table_name = case_when(is_header == TRUE &  ave(is.na(table_name), doc_index, FUN = all) ~ "REMOVE",
                                  TRUE ~ table_name),
           table_name = ave(table_name, doc_index, FUN = function(x) unique(x[!is.na(x)]))) %>%
    left_join(tab_names, by = c("table_name")) %>% 
    ungroup() %>%
    filter(table_name != "REMOVE") %>%
    select(doc_index, table_name, is_header, row_id, cell_id, text)
  
  ## Captions for figures that will need to be removed
  fig_heads <- content %>% 
    filter(content_type %in% "paragraph",
           grepl(paste0("Figure\\s\\d", stock_sd$SpeciesCommonName, collapse = "|"), text)) %>% 
    mutate(figure_name = case_when(grepl("Summary of the stock assessment", text) ~ "stocksummary",
                                   grepl("Historical assessment results", text) ~ "historicalassessment",
                                   TRUE ~ NA_character_)) %>% 
    select(doc_index, figure_name, text)
  
  
  # table <- "catchoptionsbasis"
  # table <- "catchoptions"
  # update_cols = NULL
  # update_header = FALSE
  # erase_cols = NULL
  # add_year = TRUE
  # add_column = FALSE
  
  ## Find table and perform necessary actions based on instructions for table_name ##
  table_fix <- function(x, table, 
                        update_header = FALSE, 
                        update_cols = NULL,
                        erase_cols = NULL, 
                        add_column = FALSE,
                        add_year = FALSE, 
                        messages = TRUE) {
    
    if(!table %in% unique(table_cells$table_name)) {
      cat(paste0("table_name = ", table, " is not a table cell in the table_cells tibble"))
    } # Close table name catch
    
    if(table %in% unique(table_cells$table_name)){
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
      
      names(table_body) <- sprintf("col_%s", names(table_body))
      
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
      
      #### Table editing steps ####
      
      ## Update columns logic
      if(!is.null(update_cols)){
        
        ## If a numeric column number
        if(is.numeric(update_cols)) {
          update_cols_id <- update_cols
        } # close the column number if
        
        ## If a character string column name
        if(is.character(update_cols)) {
          if("all" %in% update_cols) {
            update_cols_id <- unique(tc$cell_id)
          } else {
            update_cols_id <- tc$cell_id[tc$text %in% update_cols]
          }
        } # Close the update column if else
        
        for(i in update_cols_id) {
          # j = unlist(table_body[,i])
          
          for(j in 1:nrow(table_body[,i])) {
            
            # First, get all the plain values e.g., SSB (2019)
            new_vals <- sprintf("(%s)", as.numeric(gsub("\\(|\\)", 
                                                        "",
                                                        stringr::str_extract(table_body[j,i], "\\(\\d{4}\\)"))) + 1)
            
            table_body[j,i] <- stringr::str_replace_all(table_body[j,i], "\\(\\d{4}\\)", new_vals)
            
            # Next, get the hyphenated values e.g., Rage 3 (2017-2018)
            rangers <- data.frame(X = gsub("\\(|\\)", "", 
                                           stringr::str_extract(table_body[j,i], "\\(\\d{4}.*-?\\d{4}\\)")),
                                  stringsAsFactors = FALSE) %>% 
              tidyr::separate(X, c("A", "B")) %>% 
              mutate(C = sprintf("(%s-%s)", as.numeric(A) + 1, as.numeric(B) + 1))
            
            table_body[j,i] <- stringr::str_replace(table_body[j,i], "\\(\\d{4}.*-?\\d{4}\\)", 
                                                    rangers$C)
            
            ## If there are references w/ characters in the parentheses, just put ICES "(UPDATE REF)"
            if(grepl("ICES \\(\\d{4}.*\\)", table_body[j,i])) {
              table_body[j,i] <- "ICES (UPDATE REF)"
            } # close reference loop
            
            ## If there are any F years to update
            if(grepl("F\\d{4}", table_body[j,i])) {
              f_vals <- sprintf("F%s", as.numeric(gsub("F", 
                                                       "",
                                                       stringr::str_extract(table_body[j,i], "F\\d{4}"))) + 1)
              table_body[j, i] <- stringr::str_replace_all(table_body[j, i], "F\\d{4}", f_vals)
            } # Close update F years
            
            ## If there are tonne values, replace with "update value"
            if(table == "catchdistribution" &
               grepl("\\d.* [T-t]onnes|\\d+\\.*\\d*%", table_body[j,i])){
              
              table_body[j,i] <- stringr::str_replace_all(table_body[j, i], "\\d.* [T-t]onnes|\\d+\\.*\\d*%", "UPDATE VALUE")
            } # Close update tonne values
            
          } # close add a year loop
        } # Close update_cols_id loop
      } # Close update_cols logic
      
      ## Erase columns logic
      if(!is.null(erase_cols)){
        ## Erase columns by number
        if(is.numeric(erase_cols)) {
          erase_cols_id <- erase_cols
        } # Close erase by number
        
        ## Erase columns by name
        if(is.character(erase_cols)) {
          if("all" %in% erase_cols) {
            erase_cols_id <- unique(tc$cell_id)[-1]
          } else {
            erase_cols_id <- tc$cell_id[tc$text %in% erase_cols]
          }
        } # Close erase by name
        table_body[ , erase_cols_id] <- ""
      } # Close erase_cols logic
      
      ## Add extra rows for years
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
      } # Close add_year logic
      
      ## Add extra column for % advice change
      if(add_column){
        if(table %in% c("catchoptions")){
          empty_table <- table_body[,1]
          empty_table[] <- ""
          table_body <- bind_cols(table_body, empty_table)
          table_header <- c(table_header, "% Advice change")
        } else {
          stop("add_column only works on the 'catch options' table. Check your table")
        }
      } # close extra column logic
      
      ## Update the year values in headers
      if(update_header){
        # If all values are numeric, add a year
        table_header[grepl("\\(\\d{4}\\)", table_header)] <- sprintf("%s (%s)", gsub("\\(\\d{4}\\)", "",
                                                                                     table_header[grepl("\\(\\d{4}\\)", 
                                                                                                        table_header)]), 
                                                                     as.numeric(gsub(".*\\((.*)\\).*", "\\1", 
                                                                                     table_header[grepl("\\(\\d{4}\\)",
                                                                                                        table_header)])) + 1)
      } # Close update_header logic
      
      # Catch options merge rows
      if(table %in% "catchoptions"){
        body_merge_rows <- grep("ICES advice basis|Other options|Mixed fisheries", c(table_body[[1]]))
        for(i in body_merge_rows){
          table_body[i,] <- table_body[i,1]
        }
      } # close catch options merge rows 
      
      #### Formatting steps #### 
      
      table_info <- data.frame(col_keys = colnames(table_body),
                               table_header,
                               stringsAsFactors = FALSE)
      
      ## Default body style ##
      def_text_body <- fp_text(color = "black", font.size = 9, font.family = "Calibri", bold = FALSE, shading.color = "transparent")
      def_cell_body <- fp_cell(background.color = "transparent", border = fp_border(color = "black"))
      def_par_body <- fp_par(text.align = "right", padding.right = 3)
      
      ## Default header style ##
      def_text_head <- update(def_text_body) # NO CHANGE
      def_cell_head <- update(def_cell_body, background.color = "#E8EAEA")
      def_par_head <- update(def_par_body, text.align = "center", padding.right = 0)
      
      ## Column alignment rules
      # •	Column containing YEAR: align content center center
      # •	Column containing TEXT: align content center left
      # •	Column containing NUMBERS: align content center right 
      
      ## align columns
      align_cols <- function(alignment = c("left", "center"),
                             table_type = c("catchoptionsbasis", "catchoptions", "advicebasis",
                                            "referencepoints", "assessmentbasis", "advice",
                                            "catchdistribution")) {
        table_type <- match.arg(table_type)
        if(alignment == "left") {
          switch(table_type, 
                 catchoptionsbasis = 1,
                 catchoptions = 1,
                 advicebasis = c(1, 2), 
                 referencepoints = c(1,2,3,4),   
                 assessmentbasis  = c(1,2),
                 advice = 2,
                 catchdistribution = NULL)
        } else if(alignment == "center") {
          switch(table_type,
                 catchoptionsbasis = 2,
                 catchoptions = NULL,
                 advicebasis = NULL,
                 referencepoints = 5,
                 assessmentbasis  = NULL,
                 advice = 1,
                 catchdistribution = 1:nrow(table_info))
        }
      }
      
      # col_num <- max(as.numeric(table_info$col_keys))
      col_num <- nrow(table_info)
      col_width <- rep.int((18/2.54)/col_num, col_num)
      row_num <- nrow(table_body)
      row_height <- rep.int(0.1, row_num)
      
      ft_style <- flextable(table_body) %>%
        set_header_df(mapping = table_info, key = "col_keys") %>% 
        flextable::style(pr_t = def_text_body, # fp_text
                         pr_p = def_par_body,
                         pr_c = def_cell_body,
                         part = "body") %>% 
        flextable::style(pr_t = def_text_head , # fp_text
                         pr_p = def_par_head,
                         pr_c = def_cell_head,
                         part = "header") %>% 
        flextable::width(width = col_width) %>% 
        flextable::height(height = row_height, part = "body") %>% 
        flextable::height(height = 0.1, part = "header") 
      
      ## Align cells
      body_align_left <- align_cols(alignment = "left",
                                    table_type = table)
      body_align_center <- align_cols(alignment = "center",
                                      table_type = table)
      if(!is.null(body_align_left)) {
        ft_style <- flextable::align(ft_style, 
                                     i = NULL, 
                                     j = body_align_left, align = "left", part = "body")
        
        ft_style <- flextable::padding(ft_style, 
                                       i = NULL, 
                                       j = body_align_left, padding.left = 3, part = "body")
      }
      if(!is.null(body_align_center)){
        ft_style <- flextable::align(ft_style, 
                                     i = NULL, 
                                     j = body_align_center, align = "center", part = "body")
        
        ft_style <- flextable::padding(ft_style, 
                                       i = NULL, 
                                       j = body_align_center, padding.left = 0, part = "body")
      } 
      
      ## Merge catch options rows
      if(table %in% "catchoptions") {
        body_merge_rows <- grep("ICES advice basis|Other options|Mixed fisheries", c(table_body[[1]]))
        ft_style <- flextable::merge_h(ft_style, 
                                       i = body_merge_rows,
                                       part = "body")
        ft_style <- flextable::bg(ft_style,
                                  i = body_merge_rows, 
                                  bg = "#E8EAEA", part = "body")
      } # merge catch options rows
      
      
      ## Figure out where to put the table ##
      cc <- tab_heads %>% 
        filter(table_name == table)
      
      caption_index <- unique(cc$doc_index)[tab_index]
      
      caption_keyword <- tab_heads %>% 
        filter(table_name == table,
               doc_index == caption_index) %>% 
        mutate(text = gsub("\\(", "\\\\(", text),
               text = gsub("\\)", "\\\\)", text)) %>% 
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
          cursor_backward %>% 
          body_add_flextable(ft_style)
      }
    } # Close tab_index loop
    } # Close table name go
  } # Close table_fix function
  
  table_remove <- function(x, table){
    
    table_name <- switch(table,
                         # other = "other",
                         stocksummary = "stock summary table",
                         catchhistory = "catch history table",
                         assessmentsummary = "assessment summary table")
    
    ## Figure out where to put the table ##
    cc <- tab_heads %>% 
      filter(table_name == table)
    
    replacement_prop <- fp_text(color = "#00B0F0", 
                                font.family = "Calibri",
                                font.size = 11, 
                                bold = TRUE, 
                                italic = FALSE, 
                                underline = FALSE)
    
    replace_string <- switch(table,
                             # other = "Please look at the previous advice for this stock and insert the necessary %s here.",
                             stocksummary = "Please go to sag.ices.dk to find the %s and replace table here.",
                             catchhistory = "Please look at the previous advice for this stock and insert the %s here.",
                             assessmentsummary = "Please go to sag.ices.dk to find the %s and insert table here.")
    
    replace_text <- fpar(ftext(sprintf(replace_string, table_name),
                               prop = replacement_prop))
    
    # if(table == "stocksummary"){
    #   num_tab <- content$doc_index[grepl("^catch options$", tolower(content$text))] - cc$doc_index
    # } 
    # if(!table == "stocksummary") {
    #   num_tab <- 1
    # }
    # 
    
    for(tab_index in 1:length(unique(cc$doc_index))) {
      caption_index <- unique(cc$doc_index)[tab_index]
      
      caption_keyword <- tab_heads %>% 
        filter(table_name == table,
               doc_index == caption_index) %>% 
        mutate(text = gsub(paste0(stock_sd$SpeciesCommonName, ".*$"),"", text),
               text = paste0(text, stock_sd$SpeciesCommonName)) %>%
        pull(text)
      
      if(table == "other"){
        caption_keyword <- content %>%
          filter(doc_index == caption_index,
                 row_id == 1,
                 cell_id == 1) %>%
          pull(text)
        # filter(content_type %in% "table cell") %>%
        # group_by(doc_index)
      }
      
      if(table == "stocksummary"){
        cursor_reach(x, keyword = caption_keyword) %>%
          cursor_forward() %>%
          body_remove %>%
          # cursor_forward() %>%
          # body_remove %>%
          body_add_fpar(replace_text) %>%
          cursor_begin
      }
      if(table != "stocksummary"){
        cursor_reach(x, keyword = caption_keyword) %>% 
          cursor_forward %>% 
          body_remove %>%
          body_add_fpar(replace_text) %>% 
          cursor_begin
      }
      # cursor_backward %>% 
    }
  }
  
  figure_remove <- function(x, figure){
    
    figure_name <- switch(figure,
                          stocksummary = "stock summary figure",
                          historicalassessment = "historical assessment figure")
    if(!figure %in% fig_heads$figure_name){
      cat("Figure is either not present or the code is unable to identify it as 'stocksummary' or 'historicalassessment'")
    } else if(figure %in% fig_heads$figure_name){
      cc <- fig_heads %>% 
        filter(figure_name == figure)
      
      if(figure == "stocksummary"){
        num_fig <- cc$doc_index - content$doc_index[grepl("stock development over time", tolower(content$text))] - 3
      } 
      if(figure == "historicalassessment") {
        num_fig <- 3
      }
      
      replacement_prop <- fp_text(color = "#00B0F0", 
                                  font.family = "Calibri",
                                  font.size = 11, 
                                  bold = TRUE, 
                                  italic = FALSE, 
                                  underline = FALSE)
      
      replace_text <- fpar(ftext(sprintf("Please go to sag.ices.dk to find the %s and insert figure here.", figure_name),
                                 prop = replacement_prop))
      
      nuke_figures <- function(x, rep_num = 1) {
        for(i in 1:rep_num){
          cursor_backward(x) %>% 
            body_remove
        }
        return(x)
      }
      
      
      for(tab_index in 1:length(unique(cc$doc_index))) {
        caption_index <- unique(cc$doc_index)[tab_index]
        
        caption_keyword <- fig_heads %>% 
          filter(figure_name == figure,
                 doc_index == caption_index) %>% 
          mutate(text = gsub(paste0(stock_sd$SpeciesCommonName, ".*$"),"", text),
                 text = paste0(text, stock_sd$SpeciesCommonName)) %>% 
          pull(text)
        
        
        cursor_reach(x, keyword = caption_keyword) %>% 
          nuke_figures(rep_num = num_fig) %>%
          cursor_backward %>%
          body_add_fpar(replace_text) %>%
          cursor_begin
      }
    }
  }
  ## remove Source in catchoptions basis 
  

## ~~~~~~~~~~~~~ ##
## Format Advice ##
## ~~~~~~~~~~~~~ ##
  ## Grab the last advice
  
  # x <- officer::read_docx(fileName)
  # print(doc, target = "cursor.docx")
  
  
  # doc <- officer::read_docx(fileName)
  figure_remove(doc, figure = "stocksummary")
  figure_remove(doc, figure = "historicalassessment")
  table_remove(doc, table = "stocksummary") # %>% 
  table_fix(x = doc,
            table = "catchoptionsbasis", 
            update_header = FALSE, 
            update_cols = c("Variable", "Source"), 
            erase_cols = c("Value", "Notes"),
            add_column = FALSE,
            add_year = FALSE, 
            messages = FALSE)
  
  table_fix(x = doc,
            table = "catchoptions", 
            update_header = TRUE,
            update_cols = "Basis", 
            erase_cols = "all",
            add_column = TRUE,
            add_year = FALSE, 
            messages = FALSE) 

  table_fix(x = doc,
            table = "advice", 
            update_header = FALSE,
            update_cols = NULL, 
            erase_cols = NULL,
            add_column = FALSE,
            add_year = TRUE, 
            messages = FALSE) 
  
  table_fix(x = doc,
            table = "catchdistribution", 
            update_header = TRUE,
            update_cols = NULL, 
            erase_cols = "all",
            add_column = FALSE,
            add_year = FALSE, 
            messages = FALSE) 
  
  table_remove(x = doc, table = "assessmentsummary") 
  
  print(doc, target = sprintf("%s%s.docx", path_name, stock_name))
  return(tab_heads)
} # Close format_advice
