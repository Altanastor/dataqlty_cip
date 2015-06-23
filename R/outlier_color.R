outlier_color <- function(file=fp,trait,sheetname="Fieldbook"){
  
  sheetname <- as.character(sheetname)
  filename <- as.character(file)
  data <- xlsx::read.xlsx(file=file,sheetName = sheetname,stringsAsFactors=FALSE)
  trait <- as.character(trait)
  dict <- get.data.dict()
  trait <- trait[trait %in% names(data)]
  col <- data[,trait]
  
  if(has.data(col)){
    if(is.numeric(col)){
      lwr = as.numeric(dict[dict$ABBR==trait,"LOWER"])
      upr = as.numeric(dict[dict$ABBR==trait,"UPPER"])
      cds = get.codes(dict,trait)
      if(!(is.na(lwr) & is.na(upr) & is.na(cds))){
        
        d <- data  
        pos <- which(names(d)==trait)
        cols<-length(d[1,])
        
        wb <- xlsx::loadWorkbook(file)              # load workbook
        #fo1 <- xlsx::Fill(foregroundColor="blue")   # create fill object # 1
        #fo1 <- Fill(foregroundColor="lightblue", backgroundColor="lightblue",pattern="SOLID_FOREGROUND")
        #cs1 <- xlsx::CellStyle(wb, fill=fo1)        # create cell style # 1
        fo2 <- xlsx::Fill(foregroundColor="red")    # create fill object # 2
        #fo2 <- Fill(foregroundColor="tomato", backgroundColor="tomato",pattern="SOLID_FOREGROUND")
        cs2 <- xlsx::CellStyle(wb, fill=fo2)        # create cell style # 2 
        sheets <- xlsx::getSheets(wb)               # get all sheets
        sheet <- sheets[[sheetname]]          # get specific sheet
        rows <- xlsx::getRows(sheet, rowIndex=2:(nrow(d)+1))     # get rows
        # 1st row is headers
        #cells <- getCells(rows, colIndex = 4:cols)         # get cells
        cells <- xlsx::getCells(rows, colIndex = pos)         # get cells
        values <- lapply(cells, xlsx::getCellValue) # extract the cell values
        
        #   highlightblue <- NULL
        #   for (i in names(values)) {
        #     x <- as.numeric(values[i])
        #     
        #     if(!is.na(x)){
        #       if (x>=ll & x<=ul) {
        #         highlightblue <- c(highlightblue, i)
        #       }
        #     }
        #     
        #   }
        
        # find cells meeting conditional criteria < 5
        highlightred <- NULL
        for (i in names(values)) {
          x <- as.numeric(values[i])
          
          if(!is.na(x)){
            if (x<lwr || x>upr){
              highlightred <- c(highlightred, i)
            }
            
            if(!is.na(cds)){
              if (!(x %in% cds)){highlightred <- c(highlightred, i)}
            }
            
          }
        }
        #Finally, apply the formatting and save the workbook.
        
        #   lapply(names(cells[highlightblue]),
        #          function(ii)xlsx::setCellStyle(cells[[ii]],cs1))
        
        lapply(names(cells[highlightred]),
               function(ii)xlsx::setCellStyle(cells[[ii]],cs2))
        
        xlsx::saveWorkbook(wb, file)
      } 
    }
    #shell.exec(file)
  }
}

#fp<-file.choose()
#   varl = read.xlsx(fp,sheetName="Var List",colIndex = 1:5,h=TRUE)
#   #for(i in 1:5) varl[,i] <- as.character(varl[,i])
#     varl[,"Fieldbook"] <- as.character(varl[,"Fieldbook"])
#     trait = varl[varl[,"Fieldbook"]=="x"|varl[,"Fieldbook"]=="X",2]
#     trait = as.character(trait[!is.na(trait)])
#     for(i in 1:length(trait)){outlier_color(fp,trait[i],"Fieldbook")}


