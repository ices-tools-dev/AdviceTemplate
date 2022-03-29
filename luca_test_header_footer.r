library(officer)
library(magrittr)
library(dplyr)

# docact <- read_docx("test.docx")
# doc_advice <- read_docx("her.27.20-24_copy.docx")
# doc_advice2 <- read_docx("pok.27.3a46.docx")
# docx_summary(docact)
# tibble(docx_summary(doc_advice2))

# new <- footers_replace_all_text(
#   docact,
#   "old footer",
#   "ICES2022 footer",
#   only_at_cursor = FALSE,
#   warn = TRUE
# )
# new <- headers_replace_all_text(
#   docact,
#   "ICES2022",
#   "ICES2022 header",
#   only_at_cursor = FALSE,
#   warn = TRUE
# )
# new  %>% print(target = "Final-Output.docx")


###################
doc_advice2 <- read_docx("pok.27.3a46.docx")

## Headers
newAdvice <- headers_replace_all_text(
  doc_advice2,
  "\\d{4}.*?\\b",      #\\ba.*?\\b", #\\d{4}\\", #"^20[1-9][0-9]$","\\bn.*?\\b"
  "2054",
  only_at_cursor = FALSE,
  warn = TRUE,
  fixed=FALSE,
  ignore.case = TRUE
)

newAdvice <- headers_replace_all_text(
  doc_advice2,
  "pok.27.3a46",
  "cod.27.20-24",
  only_at_cursor = FALSE,
  warn = TRUE,
  fixed=TRUE,
  ignore.case = TRUE
)

newAdvice <- headers_replace_all_text(
  doc_advice2,
  "30 June",
  "30 January",
  only_at_cursor = FALSE,
  warn = TRUE,
  fixed=TRUE,
  ignore.case = TRUE
)
## Footers
newAdvice <- footers_replace_all_text(
  doc_advice2,
  "ICES Advice [^0-9]*[0-9]{4}",#"\\d{4}.*?\\b",#"\\ICES Advice (?=[d{4}].*?\\b",   # "\\ICES Advice d{4}.*?\\b", #paste0("ICES Advice ", "\\d{4}.*?\\b")  ICES\bAdvice\b[^0-9]*[0-9]{4}
  "ICES Advice 2070",
  only_at_cursor = FALSE,
  warn = TRUE,
  fixed=FALSE,
  ignore.case = FALSE
)

# newAdvice <- footers_replace_all_text(
#   doc_advice2,
#   "2020",
#   "2025",
#   only_at_cursor = FALSE,
#   warn = TRUE,
#   fixed=TRUE,
#   ignore.case = TRUE
# )

newAdvice <- footers_replace_all_text(
  doc_advice2,
  "advice\\.[^0-9]*[0-9]{4}", #not working
  "advice.9999",
  only_at_cursor = FALSE,
  warn = TRUE,
  fixed= TRUE,
  ignore.case = FALSE
)


newAdvice <- footers_replace_all_text(
  doc_advice2,
  "pok.27.3a46",
  "cod.27.20-24",
  only_at_cursor = FALSE,
  warn = TRUE,
  fixed=TRUE,
  ignore.case = TRUE
)
newAdvice  %>% print(target = "Final-Output_Advice2.docx")

###########################function that puts togher the code above################################

change_footer_header <- function(advice_word_document, Year, Stockcode, Day, Month, doiEnd) {
  doc_advice <- read_docx(advice_word_document)
  ## Headers
  newAdvice <- headers_replace_all_text(
    doc_advice,
    "\\d{4}.*?\\b", # \\ba.*?\\b", #\\d{4}\\", #"^20[1-9][0-9]$","\\bn.*?\\b"
    Year,
    only_at_cursor = FALSE,
    warn = TRUE,
    fixed = FALSE,
    ignore.case = TRUE
  )

  newAdvice <- headers_replace_all_text(
    doc_advice,
    "pok.27.3a46",
    Stockcode,
    only_at_cursor = FALSE,
    warn = TRUE,
    fixed = TRUE,
    ignore.case = TRUE
  )

  newAdvice <- headers_replace_all_text(
    doc_advice,
    "30 June",
    paste0(Day, " ", Month),
    only_at_cursor = FALSE,
    warn = TRUE,
    fixed = TRUE,
    ignore.case = TRUE
  )

  ## Footers
  newAdvice <- footers_replace_all_text(
    doc_advice,
    "ICES Advice [^0-9]*[0-9]{4}", # "\\d{4}.*?\\b",#"\\ICES Advice (?=[d{4}].*?\\b",   # "\\ICES Advice d{4}.*?\\b", #paste0("ICES Advice ", "\\d{4}.*?\\b")  ICES\bAdvice\b[^0-9]*[0-9]{4}
    paste0("ICES Advice", " ", Year),
    only_at_cursor = FALSE,
    warn = TRUE,
    fixed = FALSE,
    ignore.case = FALSE
  )
  # DOI not working
  newAdvice <- footers_replace_all_text(
    doc_advice,
    "ices\\.advice\\.[^0-9]*[0-9]{4}$",#"b/^10.17895/[^\s]+$/i", # not working   #
    paste0("ices.advice.", doiEnd),
    only_at_cursor = FALSE,
    warn = TRUE,
    fixed = TRUE,
    ignore.case = FALSE
  )

  newAdvice <- footers_replace_all_text(
    doc_advice,
    "pok.27.3a46",
    Stockcode,
    only_at_cursor = FALSE,
    warn = TRUE,
    fixed = TRUE,
    ignore.case = TRUE
  )
  newAdvice %>% print(target = "Final-Output_Advice2.docx")
}
change_footer_header("pok.27.3a46.docx", "2070", "cod.27.20-24", "30", "March","9999")
