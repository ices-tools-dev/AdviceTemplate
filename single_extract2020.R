# dat <- List[[n]]

stock_name <- dat$StockKeyLabel

# Check the correct doi in the excel

# doi <- doi_list[n]

fulldoi<- paste0("https://doi.org/",dat$DOI)

#You can select from here to the end and just run it

## ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ##
## Identify the tables to keep and clean ##
## ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ##

stock_sd <- fileList %>% 
  filter(StockKeyLabel == stock_name) 

stock_sd <- unique(stock_sd)

fileName <- stock_sd %>% 
  pull(filepath)
## Grab the last advice
doc <- officer::read_docx(fileName)
# doc <- officer::read_docx("pra.27.3a4a.docx")

#little trick to change the first heading from "ICES stock advice" to 
#"ICES advice on fishing opportunities". To be removed next year.
# cursor_begin(doc)
# cursor_forward(doc)
# cursor_forward(doc)

#THIS TWO NEXT LINES ARE ONLY NEEDED WHEN THE SENTENCE IS NOT ICES advice on fishing opportunities

if(dat$YearOfLastAssessment < 2018){
body_replace_all_text(x= doc, "ICES stock advice", "ICES advice on fishing opportunities", only_at_cursor = TRUE, ignore.case=FALSE)
}

## Pull out a data.frame of text (content[]) and tables (both tabs[] and content[])
tabs <- docxtractr::read_docx(fileName)
content <- officer::docx_summary(doc)


def_text_ref <- fp_text(color = "black", font.size = 10, font.family = "Calibri", bold = FALSE, italic= FALSE)
text1<-paste(c("Recommended citation: ICES. 2020. ",stock_sd$StockKeyDescription, ". In Report of the ICES Advisory Committee, 2020. ICES Advice 2020, ", stock_sd$StockKeyLabel,", ", fulldoi ), collapse="")
fpar <- fpar(ftext(text1,
                   prop= def_text_ref))

textn <- "ICES. 2019. Advice basis. In Report of the ICES Advisory Committee, 2019. ICES Advice 2019, Book 1, Section 1.2. https://doi.org/10.17895/ices.pub.5757"
fpar2 <- fpar(ftext(textn,
                   prop= def_text_ref))

# This piece to add the reccomended citation:

if(dat$YearOfLastAssessment < 2019){
  cursor_end(doc)%>%
  body_add_fpar(fpar2)%>%
    body_add_fpar(fpar)
}

# For 2019 advice, I will try to fix it a bit, but maybe it has to be formatted

if(dat$YearOfLastAssessment == 2019){
  cursor_end(doc) %>% body_remove()%>%
    body_add_fpar(fpar2)%>%
    body_add_fpar(fpar)
  # text2<-paste0("Recommended citation: ICES. 2019. ",stock_sd$StockKeyDescription, ". In Report of the ICES Advisory Committee, 2019. ICES Advice 2019, ", stock_sd$StockKeyLabel,", ", "https://doi.org/",dat$oldDOI)
  # 
  # body_replace_all_text(x= doc, text2, text1, only_at_cursor = FALSE, ignore.case=FALSE)
  # text3 <- fulldoi
  # fpar <- fpar(ftext(text3, prop = def_text_ref))
  # cursor_end(doc)%>%
  #   body_add_fpar(fpar)
}



## ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ##
## Define tables according to header and caption ##  
## ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ##

table_header_name <- content %>% 
  filter(content_type %in% "table cell",
         row_id == 1) %>%
  select(-dplyr::one_of("content_type", "style_name", "level", 
                        "num_id", "row_id", "is_header", "col_span", "row_span")) %>% 
  spread(cell_id, text) %>% 
  mutate(header_name = case_when(grepl("^[F-f]ishing pressure", `2`) ~ "stocksummary",
                                 grepl("^[V-v]ariable", `1`) &
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
                                 grepl("^[T-t]otal catch", `1`) &
                                   grepl("^[L-l]andings", `2`) ~ "catchdistribution",
                                 grepl("^[C-c]atch", `1`) &
                                   grepl("^[C-c]ommercial landings", `2`) ~ "catchdistribution",
                                 grepl("^[Y-year]", `1`) &
                                   grepl("^[R-r]ecruitment", `2`) ~ "assessmentsummary",
                                 # grepl( `2`, (\\d{4})) &
                                 #    grepl(`3`,(\\d{4}) ) ~ "catchhistory",
                                 TRUE ~ "other"),
         caption_index = doc_index - 1) %>% 
  select(doc_index,
         caption_index,
         header_name)


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

table_header_name <- unique(table_header_name)


## Add caption text to the data.frame based on the doc_index

table_header_name$text <- content$text[content$doc_index %in% table_header_name$caption_index]

# table_header_name <- table_header_name%>% filter(header_name  == "catchdistribution")%>%
#     body_replace_at(table_header_name$text, "2016", "2017")

## Grab tibble of table names - to link with the table header info, below
tab_heads <- table_header_name %>%
  select(doc_index,
         table_name = header_name,
         text)

tab_names <- tab_heads %>%
  select(table_name) %>%
  filter(!table_name %in% c("stocksummary")) %>%
  distinct(.keep_all = TRUE)
#removing source column in table catchbasis, check next year, as wont be needed
#content <- content %>% filter(!(text=="Source" & cell_id==3)) %>%
#  mutate(cell_id = replace(cell_id, which(doc_index == 26 & cell_id == 4), 3))

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
  select(content_type,doc_index, table_name, is_header, row_id, cell_id, text)

# removing source column in this table, check next year, as wont be needed  



content<- left_join(content, table_cells)%>%
  filter(!(content_type %in% "table cell"& table_name=="catchoptionsbasis" & cell_id==3))%>%
  # mutate(cell_id = replace(cell_id, which(table_name == "catchoptionsbasis"& cell_id == 4), 3))%>%
  select(-table_name)

table_cells <- table_cells %>% #filter(!(table_name=="catchoptionsbasis" & cell_id==3))%>%
  #mutate(cell_id = replace(cell_id, which(table_name == "catchoptionsbasis"& cell_id == 4), 3))%>%
  select(-content_type)


# Change in Assessment basis the reference to ICES, 2019 

table_cells$text[which(table_cells$table_name == "assessmentbasis" & 
                         table_cells$row_id == 1 & table_cells$cell_id == 2)]  <- gsub("[0-9]{4}", "2019",table_cells$text[which(table_cells$table_name == "assessmentbasis" & 
                                                       table_cells$row_id == 1 & table_cells$cell_id == 2)] )


## Captions for figures that will need to be removed
fig_heads <- content %>% 
  filter(content_type %in% "paragraph",
         grepl(paste0( "Figure\\s\\d",stock_sd$SpeciesCommonName, collapse = "|"), text)) %>% 
  mutate(figure_name = case_when(grepl("Summary of the stock assessment", text) ~ "stocksummary",
                                 grepl("Historical assessment results", text) ~ "historicalassessment",
                                 TRUE ~ NA_character_)) %>% 
  select(doc_index, figure_name, text)


## ~~~~~~~~~~~~~ ##
## Format Advice ##
## ~~~~~~~~~~~~~ ##

# doc <- officer::read_docx(fileName)
try(figure_remove(doc, figure = "stocksummary"))
#here
try(figure_remove(doc, figure = "historicalassessment"))

# figure = "historicalassessment"  
# 
# figure_name <- switch(figure,
#                       stocksummary = "stock summary figure",
#                       historicalassessment = "historical assessment figure",
#                       stockstatus= "stock status table")
# # if(!figure %in% fig_heads$figure_name){
# #   cat("Figure is either not present or the code is unable to identify it as 'stocksummary' or 'historicalassessment'")
# # } else if(figure %in% fig_heads$figure_name){
# cc <- fig_heads %>%
#   filter(figure_name == figure)
# 
# num_fig <- 3
# 
# nuke_figures <- function(x, rep_num = 1) {
#   for(i in 1:rep_num){
#     cursor_backward(x) %>% 
#       body_remove
#   }
#   return(x)
# }
# 
# 
# caption_index <- unique(cc$doc_index)
# 
# caption_keyword <- fig_heads %>% 
#   filter(figure_name == figure,
#          doc_index == caption_index) %>% 
#   mutate(text = gsub(paste0(stock_sd$SpeciesCommonName, ".*$"),"", text),
#          text = paste0(text, stock_sd$SpeciesCommonName)) %>% 
#   pull(text)
# 
# 
# cursor_reach(doc, keyword = caption_keyword) %>% 
#   nuke_figures(rep_num = num_fig) %>%
#   cursor_backward %>%
#   body_add_fpar(replace_text) %>%
#   cursor_begin



table = "stocksummary"

table_name <- switch(table,
                     # other = "other",
                     stocksummary = "stock summary table",
                     catchhistory = "catch history table",
                     assessmentsummary = "assessment summary table")

## Figure out where to put the table ##
cc <- tab_heads %>% 
  filter(table_name == table)

replace_string <- switch(table,
                         # other = "Please look at the previous advice for this stock and insert the necessary %s here.",
                         stocksummary = "Please go to sag.ices.dk to find the %s and replace table here.",
                         catchhistory = "Please look at the previous advice for this stock and insert the %s here.",
                         assessmentsummary = "Please go to sag.ices.dk to find the %s and insert table here.")

replacement_prop <- fp_text(color = "#00B0F0", 
                            font.family = "Calibri",
                            font.size = 11, 
                            bold = TRUE, 
                            italic = FALSE, 
                            underline = FALSE)

replace_text <- fpar(ftext(sprintf(replace_string, table_name),
                           prop = replacement_prop))

  caption_index <- unique(cc$doc_index)
  
  caption_keyword <- tab_heads %>% 
    filter(table_name == table,
           doc_index == caption_index) %>% 
    mutate(text = gsub(paste0(stock_sd$SpeciesCommonName, ".*$"),"", text),
           text = paste0(text, stock_sd$SpeciesCommonName)) %>%
    pull(text)
  
  # if(table == "other"){
  #   caption_keyword <- content %>%
  #     filter(doc_index == caption_index,
  #            row_id == 1,
  #            cell_id == 1) %>%
  #     pull(text)
  #   # filter(content_type %in% "table cell") %>%
  #   # group_by(doc_index)
  # }
  
  
    #rep_num = content$doc_index[grepl("Catch options", (content$text))] - content$doc_index[grepl("Table 1", tolower(content$text))]%>%
    try(cursor_reach(doc, keyword = caption_keyword) %>%
      #for(i in 1:rep_num){
      # cursor_forward() %>%
      #   body_remove}
      
      
      # cursor_reach(x, keyword = caption_keyword) %>%
      #           cursor_forward() %>%
      #           body_remove %>%
      cursor_forward() %>%
      body_remove %>%
      body_add_fpar(replace_text) %>%
      # 
      # return(x)
      # body_add_fpar(replace_text) %>%
      cursor_begin)
  

# if(table != "stocksummary"){
#   cursor_reach(x, keyword = caption_keyword) %>% 
#     cursor_forward %>% 
#     body_remove %>%
#     body_add_fpar(replace_text) %>% 
#     cursor_begin
# }
# cursor_backward %>% 
# }#closes table remove
#}

# table_remove(doc, table = "stocksummary")
try(table_fix(x = doc,
          table = "catchoptionsbasis", 
          update_header = FALSE, 
          update_cols = "Variable", 
          erase_cols = c("Value", "Notes"),
          add_column = FALSE,
          add_year = FALSE, 
          messages = FALSE))

try(table_fix(x = doc,
          table = "catchoptions", 
          update_header = TRUE,
          update_cols = "Basis", 
          erase_cols = "all",
          add_column = TRUE,
          add_year = FALSE, 
          messages = FALSE))

try(table_fix(x = doc,
          table = "advice", 
          update_header = FALSE,
          update_cols = NULL, 
          erase_cols = NULL,
          add_column = FALSE,
          add_year = TRUE, 
          messages = FALSE))

# This is getting too difficult, not worth it

# table_fix(x = doc,
#           table = "assessmentbasis",
#           update_header = FALSE,
#           update_cols = NULL,
#           erase_cols = NULL,
#           add_column = FALSE,
#           add_year = FALSE,
#           messages = FALSE)


# table = "assessmentbasis"
# # I would like to add the formatting of this table here.
# table[, 1] = cellProperties(background.color = "#E8EAEA")
# assessmentBasisTable[, c(1:2)] = parProperties(text.align = "left")
# setFlexTableWidths(assessmentBasisTable, c((3.55/2.54), (14.45/2.54)))



warning_prop <- fp_text(color = "#00B0F0", 
                        font.family = "Calibri",
                        font.size = 11, 
                        bold = TRUE, 
                        italic = FALSE, 
                        underline = FALSE) 



fpar <- fpar(ftext(paste0("Please go to sag.ices.dk to find the assessment summary table and insert table here."),
                   prop= warning_prop)) 

try(table_remove(doc, table = "assessmentsummary"))

cursor_reach(doc, "Sources and references")%>%
  cursor_backward() %>%
  body_add_fpar(fpar)


#a few changes in the text for 2018, check next year  

body_replace_all_text(x= doc, "atch options", "atch scenarios", only_at_cursor = FALSE, ignore.case=TRUE)
body_replace_all_text(x= doc, "atch option", "atch scenario", only_at_cursor = FALSE, ignore.case=TRUE)
body_replace_all_text(x= doc, " Catch distribution by fleet in 2018", " Catch distribution by fleet in 2019", only_at_cursor = FALSE, ignore.case=TRUE)
body_replace_all_text(x= doc, "Please go to sag.ices.dk to find the stock summary figure and insert figure here.", "Please go to sag.ices.dk to find the stock summary figure and insert figure here. (MANAGE DATA AND GRAPHS (LOGIN))", only_at_cursor = FALSE, ignore.case=TRUE)
body_replace_all_text(x= doc, "Please go to sag.ices.dk to find the stock summary table and replace table here.", "Please go to sag.ices.dk to find the stock summary table and replace table here. (MANAGE DATA AND GRAPHS (LOGIN))", only_at_cursor = FALSE, ignore.case=TRUE)

## Changing reference to Advice basis 2019, several ways to cite...
# nothing of these seems to work!!
# text1 <- "ICES. 2018c. Advice basis. In Report of the ICES Advisory Committee, 2018. ICES Advice 2018, Book 1, Section 1.2. https://doi.org/10.17895/ices.pub.4503"
# text2 <- "ICES. 2018b. Advice basis. In Report of the ICES Advisory Committee, 2018. ICES Advice 2018, Book 1, Section 1.2. https://doi.org/10.17895/ices.pub.4503"
# text3 <- "ICES. 2018a. Advice basis. In Report of the ICES Advisory Committee, 2018. ICES Advice 2018, Book 1, Section 1.2. https://doi.org/10.17895/ices.pub.4503"
# text4 <- "ICES. 2018a. Advice basis. In Report of the ICES Advisory Committee, 2018. ICES Advice 2018, Book 1, Section 1.2. https://doi.org/10.17895/ices.pub.4503"
# 
# textn <- "ICES. 2019. Advice basis. In Report of the ICES Advisory Committee, 2019. ICES Advice 2019, Book 1, Section 1.2. https://doi.org/10.17895/ices.pub.XXXX"
# 
# body_replace_all_text(x= doc, "ICES. 2018c. Advice basis. In Report of the ICES Advisory Committee, 2018. ICES Advice 2018, Book 1, Section 1.2. https://doi.org/10.17895/ices.pub.4503",
#                       "ICES. 2019. Advice basis. In Report of the ICES Advisory Committee, 2019. ICES Advice 2019, Book 1, Section 1.2. https://doi.org/10.17895/ices.pub.XXXX",
#                       only_at_cursor = FALSE, ignore.case=TRUE)
# body_replace_all_text(x= doc, "ICES. 2018. Advice basis.", "ICES. 2019. Advice basis.", only_at_cursor = FALSE, ignore.case=TRUE)
# body_replace_all_text(x= doc, "ICES. 2018c. Advice basis.", "ICES. 2019. Advice basis.", only_at_cursor = FALSE, ignore.case=TRUE)
# body_replace_all_text(x= doc, "ICES. 2016. Advice basis.", "ICES. 2019. Advice basis. ", only_at_cursor = FALSE, ignore.case=TRUE)
# body_replace_all_text(x= doc, "ICES. 2016a. Advice basis.", "ICES. 2019. Advice basis. ", only_at_cursor = FALSE, ignore.case=TRUE)
# body_replace_all_text(x= doc, "ICES. 2016b. Advice basis.", "ICES. 2019. Advice basis. ", only_at_cursor = FALSE, ignore.case=TRUE)
# 
# body_replace_all_text(x= doc, "Report of the ICES Advisory Committee, 2018", "Report of the ICES Advisory Committee, 2019", only_at_cursor = FALSE, ignore.case=TRUE)
# body_replace_all_text(x= doc, "2018, Book 1, Section 1.2", "2019, Section 1.2", only_at_cursor = FALSE, ignore.case=TRUE)
# body_replace_all_text(x= doc, "https://doi.org/10.17895/ices.advice. 4503", "https://doi.org/10.17895/ices.advice. 5757", only_at_cursor = FALSE, ignore.case=TRUE)
# 
# body_replace_all_text(x= doc, "Report of the ICES Advisory Committee, 2016", "Report of the ICES Advisory Committee, 2019", only_at_cursor = FALSE, ignore.case=TRUE)
# body_replace_all_text(x= doc, "2016, Book 1, Section 1.2", "2019, Section 1.2. https://doi.org/10.17895/ices.advice. 5757", only_at_cursor = FALSE, ignore.case=TRUE)


#warning messages to be added in the draft:

warning_prop <- fp_text(color = "#00B0F0", 
                        font.family = "Calibri",
                        font.size = 11, 
                        bold = TRUE, 
                        italic = FALSE, 
                        underline = FALSE)

# Need to ask Jette about this one
fpar <- fpar(ftext(("Please add a sentence describing the status of the stock. (see Guidance for preparing single-stock advice)."),
                   prop= warning_prop))
try(cursor_bookmark(x=doc, "STOCK_STATUS_TABLE_CAPTION") %>% 
  cursor_backward()%>%  # move cursor forward once, we will be on the content to be remove
  body_add_fpar(fpar))
# set cursor on paragraph containing 'a_bkmk'
# cursor_backward()%>%  # move cursor forward once, we will be on the content to be removed
# body_add_fpar(fpar)           


if(dat$YearOfLastAssessment < 2018){
fpar <- fpar(ftext(("Include a paragraph explaining any change in advice, no matter if this is big or small."),
                   prop= warning_prop))
cursor_bookmark(x=doc, "ADVICE_BASIS_TABLE_CAPTION")%>%
  cursor_backward()%>%
  cursor_backward()%>%
  cursor_backward()%>%
  body_add_fpar(fpar)
}

# [sag.ices.dk](http://standardgraphs.ices.dk/stockList.aspx)


# body_replace_all_text(x= doc, "sag.ices.dk", hyperlink_text(x = "sag.ices.dk", url = "http://standardgraphs.ices.dk/stockList.aspx" ), only_at_cursor = FALSE, ignore.case=TRUE)

print(doc, target = sprintf("%s%s.docx", path_name = "./", stock_name))

