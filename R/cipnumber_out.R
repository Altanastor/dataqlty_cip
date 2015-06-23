#install.packages("xlsx")
#install.packages("stringr")
#library(xlsx)
#library(stringr)

cip_number_check <- function(cipnumber){
  
  pattern <- "^(CIP)[0-9]{6}(\\.([0-9]{1,3}))?$"
  check <- stringr::str_detect(cipnumber,pattern)
  cipnumber_ok <- cipnumber[check]
  cipnumber_wrong <- cipnumber[!check]
  out <- list(cipnumber_ok=cipnumber_ok, cipnumber_wrong=cipnumber_wrong)
  
}

cipnumber_out<- function(file,cip_number,sheetname="mysheet"){
  
  file <- as.character(file)
  cip_number <- as.character(cip_number)
  sheetname <- as.character(sheetname)
  filename <- as.character(file)
  data <- xlsx::read.xlsx(file=file,sheetName = sheetname,stringsAsFactors=FALSE)
  d <- data  
  pos <- which(names(d)==cip_number)
  cols<-length(d[1,])
  
  cip_wrong <- cip_number_check(data[,pos])
  cip_wrong <- cip_wrong[[2]]
  
  wb <- xlsx::loadWorkbook(file)              # load workbook
  fo1 <- xlsx::Fill(foregroundColor="yellow")
  cs1 <- xlsx::CellStyle(wb, fill=fo1)        # create cell style # 1
  sheets <- xlsx::getSheets(wb)               # get all sheets
  sheet <- sheets[[sheetname]]          # get specific sheet
  rows <- xlsx::getRows(sheet, rowIndex=2:(nrow(d)+1))     # get rows
  cells <- xlsx::getCells(rows, colIndex = pos)         # get cells
  values <- lapply(cells, xlsx::getCellValue) # extract the cell values
  
  highlightblue <- NULL
  for (i in names(values)) {
    
    x <- as.character(values[i])  
    if(!is.na(x)){
      if(x %in% cip_wrong){
        highlightblue <- c(highlightblue, i)
      }
      
    }
  }
  
  # find cells meeting conditional criteria < 5
  #Finally, apply the formatting and save the workbook.
  
  lapply(names(cells[highlightblue]),
         function(ii)xlsx::setCellStyle(cells[[ii]],cs1))
  
  
  xlsx::saveWorkbook(wb, file)
  #shell.exec(file)
}

#   file <- "D:\\Users\\obenites\\Desktop\\PTDT201409_STRSIGUAS_VHT.xls"
#   cip_number <- "INSTN"
#   sheetname <- "Fieldbook"
#cipnumber_out(fp,cip_number=INSTN,sheetname="Fieldbook")