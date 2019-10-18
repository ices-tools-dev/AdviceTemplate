fileList <- get_filelist(2018)

#Write the name of the WG

WG<- fileList %>% filter(ExpertGroup=="WGCSE")

#for nephrops of WGCSE

# WG <- WG %>% filter(AdviceDraftingGroup == "ADGNEPH")


#Check that the number of elements in this list corresponds to those in the excel
# Change the number form 1 to n number of names in the WG list

rm(fileName)
rm(stock_sd)


stock_name <- WG$StockKeyLabel[2]

# stock_name  <- "ane.27.8"

# Check the correct doi in the excel

doi <- "10.17895/ices.advice.4825"

fulldoi<- paste0("https://doi.org/",doi)



