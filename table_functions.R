###########
# PURPOSE #
###########
# Wrappers to take tables from previous advice and to format them so they can be
# added to the draft documents

############
# FUNCTIONS #
############

# This function uses a stock code to find the previous advice .docx in sharepoint and uses some simple 
# grep searches to identify the target tables.
# StockCode <- "cod.27.3a47d"
# adviceTable <- "adviceBasisTable"
tableData <- function(stock.code,
                      adviceTable , # SAG
                      header = FALSE) {
  
  fileName <- fileList$file.path[fileList$StockCode == stock.code]
  doc <- read_docx(fileName)
  
  tableCount <- unlist(lapply(fileName, function(x) docx_tbl_count(doc)))
  columnNames <- lapply(seq(1: tableCount),
                        function(x) colnames(docx_extract_tbl(doc, tbl_number = x, header = TRUE)))
  
  if(adviceTable %in% c("stockExploitationTable", "assessmentSummaryTable")) {
    stop(paste0("The ", adviceTable, " is provided by SAG.\n"))
  }
  
  if(adviceTable %in% c("catchDistributionTable", "catchHistoryTable")) {
    stop(paste0("The ", adviceTable, " is to be drafted by the stock assessor.\n"))
  }
  
  #######################
  # Catch Options Basis #
  #######################
  if(adviceTable %in% c("catchOptionsBasisTable")) {
    if(fileList$DataCategory[fileList$StockCode == stock.code] >= 2) {
      stop(paste0(adviceTable, " is only applicable for Category 1 and 2 stocks."))
    }
    match_variable <- grep("variable",  tolower(columnNames))
    
    if(length(match_variable) >= 2) {
      print(match_variable)
      print(stock.code)
      stop(paste0("Found multiple matching column headers in ", adviceTable))
    }
    
    if(length(match_variable) == 0) {
      cat(paste0("Not able to find a matching column header in ", adviceTable,
                  " , will just insert a generic basis of the catch options table."))
      generic_table <- data.frame(Variable = c("F (UPDATE)", "SSB (UPDATE)", "R (UPDATE)", "Catch (UPDATE)"),
                 Value = rep(NA, 4),
                 Source = rep("ICES (2017)", 4),
                 Notes = rep(NA, 4))
      return(generic_table)
    }
    tbl_number <- match_variable
  }
  
  #######################
  # Basis of the advice #
  #######################
  if(adviceTable %in% c("adviceBasisTable")) {
    
    match_advice_basis <- grep("advice basis",  tolower(columnNames))
    
    if(length(match_advice_basis) >= 2) {
      print(match_advice_basis)
      print(stock.code)
      stop(paste0("Found multiple matching column headers in ", adviceTable))
    }
    
    if(length(match_advice_basis) == 0) {
      print(stock.code)
      stop(paste0("Not able to find a matching column header in ", adviceTable))
    }
    tbl_number <- match_advice_basis
  }
  
  ####################  
  # Reference points #
  ####################
  if(adviceTable %in% c("referencePointsTable")) {
    if(fileList$DataCategory[fileList$StockCode == stock.code] >= 2) {
      stop(paste0(adviceTable, " is only available for Category 1 and 2 stocks."))
    }
    match_framework <- grep("framework",  tolower(columnNames))
    
    if(length(match_framework) >= 2) {
      print(stock.code)
      print(match_framework)
      stop(paste0("Found multiple matching column headers in ", adviceTable))
    }
    
    if(length(match_framework) == 0) {
      print(stock.code)
      stop(paste0("Not able to find a matching column header in ", adviceTable))
    }
    tbl_number <- match_framework
  }
  
  ###########################
  # Basis of the assessment #
  ###########################
  if(adviceTable %in% c("assessmentBasisTable")) {
    
    match_ices_stock <- grep("ices stock data category",  tolower(columnNames))
    
    if(length(match_ices_stock) >= 2) {
      print(match_ices_stock)
      print(stock.code)
      stop(paste0("Found multiple matching column headers in ", adviceTable))
    }
    
    if(length(match_ices_stock) == 0) {
      print(stock.code)
      stop(paste0("Not able to find a matching column header in ", adviceTable))
    }
    tbl_number <- match_ices_stock
  }
  
  #######################
  # Basis of the advice #
  #######################
  if(adviceTable %in% c("adviceHistoryTable")) {
    
    match_ices_advice <- grep("ices advice",  tolower(columnNames))
    
    # if(length(match_ices_advice) >= 2) {
    #   print(match_ices_advice)
    #   print(stock.code)
    #   stop(paste0("Found multiple matching column headers in ", adviceTable))
    # }
    # 
    # if(length(match_ices_advice) == 0) {
    #   print(stock.code)
    #   stop(paste0("Not able to find a matching column header in ", adviceTable))
    # }
    tbl_number <- match_ices_advice
  }
  
  #####################
  # Extract the table #
  #####################
  tableDat <- lapply(tbl_number, function(x) docx_extract_tbl(doc, 
                                                         tbl_number = x,
                                                         header = header))
  names(tableDat) <- paste0("T", tbl_number)
  return(tableDat)
}


titlePot <- function(ecoregion.name, pub.date) {
  hPot <- pot(paste0(ecoregion.name, "\t\tPublished "),
              format = header_text_prop) +
    pot(pub.date,
        format = header_text_prop)
  return(hPot)
}

headingPot <- function(stock.code, stock.name) {
  hPot <- pot(section.number,
              format = heading_text_prop) + 
    "\t" +
    pot(stock.name,
        format = heading_text_prop)
  return(hPot)
}

footerPot <- function(stock.code, section.number) {
  fPot <- pot("ICES Advice 2017",
              format = header_text_prop) +
    pot(", Book ",
        format = header_text_prop) + 
    pot(gsub("(.*?)(\\..*)", "\\1",
             section.number),
             # stockList$SECTION.NUMBER[stockList$StockCode == stock.code]),
        format = header_text_prop)
  return(fPot)
}

headerPot <- function(stock.code, pub.date) {
  hPot <- pot("ICES Advice on fishing opportunities, catch, and effort\t\t Published ",
              format = header_text_prop) +
    pot(pub.date,
        format = header_text_prop)
  return(hPot)
}

captionPot <- function(stock.code,
                       type = c("Figure", "Table"), 
                       # section.number,
                       caption.number, 
                       caption.name,
                       caption.text) {
  
  cPot <- pot(paste0(type, " ", 
                     # section.number,
                     # ".",
                     caption.number, "\t"), 
              format = fig_bold_text_prop) + 
    pot(paste0(caption.name, ". ", caption.text), 
        format = fig_base_text_prop)
  return(cPot)
}

summaryPot <- function(common.name) {
  
  toMatch <- c("Greenland", "Norway", "Portuguese")
  if(length(grep(paste(toMatch,
                       collapse = "|"), 
                 common.name)) < 1) {
    common.name <- tolower(common.name)
  }  

  sPot <- pot(paste0("There is no assessment for ",
                     common.name,
                     " in this area."),
              format = base_text_prop)
  return(sPot)
}


# Advice Basis Table
advice_basis_table <- function(stock.code) {
  adviceBasisData <- tableData(stock.code, 
                               adviceTable = "adviceBasisTable",
                               header = FALSE)
  
  if(class(adviceBasisData) == "list") {
  adviceBasisData <- adviceBasisData[[1]]
  }

  colnames(adviceBasisData) <- c("DESCRIPTION", "VALUE")
  
  adviceBasisTable <-  FlexTable(adviceBasisData, header.columns = FALSE,
                                 body.cell.props = cellProperties(padding.left = 2,
                                                                  padding.right = 2,
                                                                  padding.bottom = 0),
                                 body.text.props = textProperties(font.family = "Calibri",
                                                                  font.weight = "normal",
                                                                  font.size = 9))
  
  adviceBasisTable[, 1] = cellProperties(background.color = "#E8EAEA")
  adviceBasisTable[, c(1:2)] = parProperties(text.align = "left")
  setFlexTableWidths(adviceBasisTable, c((3.89/2.54), (13.86/2.54)))
  return(adviceBasisTable)
}

# Assessment Basis Table
assessment_basis_table <- function(stock.code, data.category, expert.name, expert.url) {
  
  assessmentBasisData <- tableData(stock.code, 
                                   adviceTable = "assessmentBasisTable",
                                   header = FALSE)
  
  
  if(class(assessmentBasisData) == "list" &
     length(assessmentBasisData) == 1) {
    assessmentBasisData <- assessmentBasisData[[1]]
  }
  
  colnames(assessmentBasisData) <- c("DESCRIPTION", "VALUE")
  assessmentBasisData$VALUE[1] <- paste0(data.category,
                                         " (ICES, UPDATE REFERENCE).")
  assessmentBasisData$VALUE[1:2] <- gsub("ICES, 201[5-6].*?", "ICES, 2017", assessmentBasisData$VALUE[1:2])
  
  if(grepl(expert.name, assessmentBasisData$VALUE[7])) {
    pot_link <- pot(gsub(paste0("*\\([", expert.name, "\\)]+\\).*"), "", assessmentBasisData$VALUE[7]),
                    format = fig_base_text_prop) +
                pot(paste0("(", expert.name, ")"),
                    hyperlink = expert.url,
                    format = fig_base_text_prop)
    assessmentBasisData$VALUE[7] <- ""
  }
  
  assessmentBasisTable <-  FlexTable(assessmentBasisData, header.columns = FALSE,
                                     body.cell.props = cellProperties(padding.left = 2, padding.right = 2, padding.bottom = 0),
                                     body.text.props = textProperties(font.family = "Calibri", font.weight = "normal", font.size = 9)
  )
  
  assessmentBasisTable[ 7, 2, to = "body"] <- pot_link
  assessmentBasisTable[, 1] = cellProperties(background.color = "#E8EAEA")
  assessmentBasisTable[, c(1:2)] = parProperties(text.align = "left")
  setFlexTableWidths(assessmentBasisTable, c((3.55/2.54), (14.45/2.54)))
  return(assessmentBasisTable)
}

# assessment_basis_table("cod.27.3a47d")
# stock.code <- "dab.27.3a4"
catch_options_basis_table <- function(stock.code) {
  
  catchBasisData <- tableData(stock.code, 
                              adviceTable = "catchOptionsBasisTable",
                              header = TRUE)
  
  if(class(catchBasisData) == "list") {
    catchBasisData <- catchBasisData[[1]]
  }
  
  colnames(catchBasisData) <- gsub("[[:punct:]]", "", colnames(catchBasisData))
  
  catchBasisData$Variable <- gsub("\\s*\\([^\\)]+\\)", " (UPDATE)", as.character(catchBasisData$Variable))
  catchBasisData$Variable <- gsub("\\s*201[0-9]+", " (UPDATE)", as.character(catchBasisData$Variable))
  catchBasisData$Source <- gsub("201[5-6].*?", "2017", catchBasisData$Source)
  catchBasisData$Value <- ""
  catchBasisData$Notes <- ""
  
  catchBasisTable <-  FlexTable(catchBasisData,
                                header.columns = TRUE,
                                body.cell.props = cellProperties(padding.left = 2, padding.right = 2, padding.bottom = 0),
                                body.par.props =  parProperties(text.align = "right", padding = 1),
                                body.text.props = textProperties(font.family = "Calibri", font.weight = "normal", font.size = 9),
                                header.cell.props = cellProperties(background.color = "#E8EAEA"),
                                header.par.props =  parProperties(text.align = "center", padding = 1),
                                header.text.props = textProperties(font.family = "Calibri", font.weight = "normal", font.size = 9)
  )
  
  catchBasisTable[, c("Source"), to = "body"] = parProperties(text.align = "center", padding = 1)
  catchBasisTable[, c("Variable", "Notes"), to = "body"] = parProperties(text.align = "left", padding = 1)
  setFlexTableWidths(catchBasisTable, c(4.26/2.54, 2.25/2.54, 3/2.54, 8.45/2.54))
  
  return(catchBasisTable)
}


advice_history_table <- function(stock.code,
                                 data.category) {
  
  numYears <- ifelse(data.category %in% c(1,2),
                     1,
                     2)
  
  adviceHistoryData <- tableData(stock.code, 
                                 adviceTable = "adviceHistoryTable",
                                 header = TRUE)
  
  catchTables <- names(adviceHistoryData)
  
  adviceHistoryTable <- vector("list", length(catchTables))
  names(adviceHistoryTable) <- catchTables
  # i <- "T7"
  for(i in catchTables) {

    temprow <- data.frame(matrix(c(rep.int(NA, length(adviceHistoryData[[i]]))),
                                 nrow = numYears,
                                 ncol = length(adviceHistoryData[[i]])))
    colnames(temprow) <- colnames(adviceHistoryData[[i]])
    
    temprow[1] <- seq(from = 2018, length = numYears)
    adviceHistoryData[[i]] <- rbind(adviceHistoryData[[i]],
                               temprow)
    #
    adviceHistoryTable <-  FlexTable(adviceHistoryData[[i]] , header.columns = TRUE,
                                     body.cell.props = cellProperties(padding.left = 2, padding.right = 2, padding.bottom = 0),
                                     body.par.props =  parProperties(text.align = "right", padding = 1),
                                     body.text.props = textProperties(font.family = "Calibri", font.weight = "normal", font.size = 9),
                                     header.cell.props = cellProperties(background.color = "#E8EAEA"),
                                     header.par.props =  parProperties(text.align = "right", padding = 1),
                                     header.text.props = textProperties(font.family = "Calibri", font.weight = "normal", font.size = 9)
    )
    
    adviceHistoryTable[, 1, to = "body"] = parProperties(text.align = "center", padding = 1)
    adviceHistoryTable[, 2, to = "body"] = parProperties(text.align = "left", padding = 1)
    adviceHistoryTable[, 1, to = "header"] = parProperties(text.align = "center", padding = 1)
    adviceHistoryTable[, 2, to = "header"] = parProperties(text.align = "left", padding = 1)
    setFlexTableWidths(adviceHistoryTable, rep((17.91/2.54) / ncol(adviceHistoryData[[i]]),
                                               ncol(adviceHistoryData[[i]])))
    adviceHistoryTable[[i]] <- adviceHistoryTable
  }
  return(adviceHistoryTable)
}