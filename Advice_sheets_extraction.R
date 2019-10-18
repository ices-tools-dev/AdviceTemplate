#This process has to be repeated for each sheet,
#whereas the Advice_preparation_setup has to be run only once every session.



fileList <- get_filelist(2018)

#Write the name of the WG

WG<- fileList %>% filter(ExpertGroup=="NWWG")

#Check that the number of elements in this list corresponds to those in the excel
# Change the number form 1 to n number of names in the WG list

stock_name <- WG$StockKeyLabel[2]

stock_name 

# Check the correct doi in the excel

doi <- "10.17895/ices.advice.4735"

fulldoi<- paste0("https://doi.org/",doi)

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
cursor_begin(doc)
cursor_forward(doc)
cursor_forward(doc)

#THIS TWO NEXT LINES ARE ONLY NEEDED WHEN THE SENTENCE IS NOT ICES advice on fishing opportunities

# body_replace_all_text(x= doc, "stock", "", only_at_cursor = TRUE, ignore.case=FALSE)
# body_replace_all_text(x= doc, "advice", "advice on fishing opportunities", only_at_cursor = TRUE, ignore.case=FALSE)


## Pull out a data.frame of text (content[]) and tables (both tabs[] and content[])
tabs <- docxtractr::read_docx(fileName)
content <- officer::docx_summary(doc)


text3<-paste(c("Recommended citation: ICES. 2019. ",stock_sd$StockKeyDescription, ". In Report of the ICES Advisory Committee, 2019, ", stock_sd$StockKeyLabel,", ", fulldoi ), collapse="")

def_text_ref <- fp_text(color = "black", font.size = 10, font.family = "Calibri", bold = FALSE, italic= FALSE)
fpar <- fpar(ftext(text3,
                   prop= def_text_ref))

cursor_end(doc)%>%
  body_add_fpar(fpar)


## ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ##
## Define tables according to to header and caption ##  
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
figure_remove(doc, figure = "stocksummary")
figure_remove(doc, figure = "historicalassessment")
table_remove(doc, table = "stocksummary")  
table_fix(x = doc,
          table = "catchoptionsbasis", 
          update_header = FALSE, 
          update_cols = "Variable", 
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

table_remove(doc, table = "assessmentsummary")
warning_prop <- fp_text(color = "#00B0F0", 
                        font.family = "Calibri",
                        font.size = 11, 
                        bold = TRUE, 
                        italic = FALSE, 
                        underline = FALSE)

fpar <- fpar(ftext(("Please go to sag.ices.dk to find the assessment summary table and insert table here."),
                   prop= warning_prop))

cursor_forward(x = doc)%>%  # move cursor forward once, we will be on the content to be removed
  body_add_fpar(fpar)  

#a few changes in the text for 2018, check next year  

body_replace_all_text(x= doc, "atch options", "atch scenarios", only_at_cursor = FALSE, ignore.case=TRUE)
body_replace_all_text(x= doc, "atch option", "atch scenario", only_at_cursor = FALSE, ignore.case=TRUE)
body_replace_all_text(x= doc, " Catch distribution by fleet in 2017", " Catch distribution by fleet in 2018", only_at_cursor = FALSE, ignore.case=TRUE)

#warning messages to be added in the draft:

warning_prop <- fp_text(color = "#00B0F0", 
                        font.family = "Calibri",
                        font.size = 11, 
                        bold = TRUE, 
                        italic = FALSE, 
                        underline = FALSE)

fpar <- fpar(ftext(("Please add a sentence describing the status of the stock. (see Guidance for preparing single-stock advice)."),
                   prop= warning_prop))
cursor_bookmark(x=doc, "STOCK_STATUS_TABLE_CAPTION") %>% # set cursor on paragraph containing 'a_bkmk'
  cursor_backward()%>%  # move cursor forward once, we will be on the content to be removed
  body_add_fpar(fpar)           

fpar <- fpar(ftext(("Include a paragraph explaining any change in advice, no matter if this is big or small."),
                   prop= warning_prop))
cursor_bookmark(x=doc, "ADVICE_BASIS_TABLE_CAPTION")%>%
  cursor_backward()%>%
  cursor_backward()%>%
  cursor_backward()%>%
  body_add_fpar(fpar)

print(doc, target = sprintf("%s%s.docx", path_name = "~/", stock_name))

# Close format_advice

