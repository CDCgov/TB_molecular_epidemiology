##make heatmap of LITT rankings
library(xlsx)

##outPrefix is the prefix for the output Excel spreadsheet
##all = all potential sources table from LITT
littHeatmap <- function(outPrefix, all) {
  ##fix names (was developed for the all potential sources output Excel file)
  all = cleanHeaderForOutput(all)
  names(all) = make.names(names(all))

  cases = rev(as.character(sort(unique(all$Given.Case))))
  
  ##score matrix
  seq = matrix(NA, nrow=length(cases), ncol=length(cases))
  row.names(seq) = cases
  colnames(seq) = cases
  for(r in 1:nrow(seq)) {
    for(c in 1:ncol(seq)) {
      sc = round(as.numeric(as.character(all$Score[all$Given.Case==colnames(seq)[r] & all$Potential.Source==row.names(seq)[c]])),
                 digits=1)
      if(length(sc)) {
        seq[r,c] = sc
      }
    }
  }
  ##without SNP score matrix
  nonseq = matrix(NA, nrow=length(cases), ncol=length(cases))
  row.names(nonseq) = cases
  colnames(nonseq) = cases
  for(r in 1:nrow(nonseq)) {
    for(c in 1:ncol(nonseq)) {
      wo = round(as.numeric(as.character(
        all$Without.SNP.Score[all$Given.Case==colnames(nonseq)[r] & all$Potential.Source==row.names(nonseq)[c]])),
        digits = 1)
      if(length(wo)) {
        nonseq[r,c] = wo
      }
    }
  }
  
  ##rank matrix
  rank = matrix(NA, nrow=length(cases), ncol=length(cases))
  row.names(rank) = cases
  colnames(rank) = cases
  for(r in 1:nrow(rank)) {
    for(c in 1:ncol(rank)) {
      rk = as.character(all$Rank[all$Given.Case==colnames(rank)[r] & all$Potential.Source==row.names(rank)[c]])
      if(length(rk)) {
        rank[r,c] = rk
      }
    }
  }
  
  ###as an Excel spreadsheet
  fileName = paste(outPrefix, heatmapFileName, sep="")
  sheetName = "heatmap"
  if(file.exists(fileName)) {
    file.remove(fileName)
  }
  ###set up sheet
  workbook = createWorkbook(type="xlsx")
  sheet = createSheet(workbook, sheetName)
  num = length(cases)+15
  rows = createRow(sheet, rowIndex = 1:num)
  cells = createCell(rows, colIndex = 1:num)
  ###write potential source label
  setCellValue(cells[[1,3]], "Potential Source")
  labelheight=18
  cs = CellStyle(workbook) + #Alignment(horizontal = "ALIGN_CENTER") + 
    Font(workbook, isBold = T, heightInPoints = labelheight)
  setCellStyle(cells[[1,3]], cs)
  addMergedRegion(sheet, startRow = 1, endRow = 1, startColumn = 3, endColumn = length(cases)+2)
  rotate = CellStyle(workbook) + Alignment(rotation = 90)
  ###write given case label
  setCellValue(cells[[3,1]], "Given Case")
  addMergedRegion(sheet, startRow = 3, endRow = length(cases)+2, startColumn = 1, endColumn = 1)
  setCellStyle(cells[[3,1]], CellStyle(workbook, alignment = Alignment(rotation = 90, vertical = "VERTICAL_TOP"), 
                                       font = Font(workbook, isBold = T, heightInPoints = labelheight)))
  
  ##write potential source and rotate
  for(c in 1:ncol(rank)) {
    setCellValue(cells[[2, 2+c]], colnames(rank)[c])
    setCellStyle(cells[[2, 2+c]], CellStyle(workbook, alignment = Alignment(rotation = 90), border=Border(position="BOTTOM")))
  }
  setRowHeight(rows[2], 98)
  ##write each potential source with its row (leave blank if NA)
  bord = CellStyle(workbook) + Border(position="BOTTOM")
  for(r in 1:nrow(rank)) {
    ##given case
    setCellValue(cells[[2+r, 2]], row.names(rank)[r])
    setCellStyle(cells[[2+r, 2]], bord)
    ##scores
    for(c in 1:ncol(rank)) {
      if(!is.na(seq[r,c])) {
        setCellValue(cells[[2+r, 2+c]], as.numeric(seq[r,c]))
      } else if(!is.na(nonseq[r,c])) {
        setCellValue(cells[[2+r, 2+c]], paste(nonseq[r,c], "*", sep=""))
      }
    }
  }
  
  ##set up cell styles
  col = CellStyle(workbook, alignment = Alignment(horizontal = "ALIGN_CENTER"))
  red = col + Fill(foregroundColor = rgb(red=204, green=0, blue=0, maxColorValue = 255))#CellStyle(workbook, fill = Fill(foregroundColor = rgb(red=204, green=0, blue=0, maxColorValue = 255)))
  orange = col + Fill(foregroundColor = rgb(red=255, green=119, blue=78, maxColorValue = 255))# CellStyle(workbook, fill = Fill(foregroundColor = rgb(red=255, green=119, blue=78, maxColorValue = 255))) 
  yellow = col + Fill(foregroundColor = rgb(red=255, green=239, blue=156, maxColorValue = 255))#CellStyle(workbook, fill = Fill(foregroundColor = rgb(red=255, green=239, blue=156, maxColorValue = 255)))
  white = col + Fill(foregroundColor = "white") + Border(position="BOTTOM")
  
  ##get color for sequenced pair
  getSeqColor<-function(score) {
    if(score <= 1) {
      return(red)
    } else if(score > 1 & score <= 3) {
      return(orange)
    } else if(score > 3 & score <= 8) { #with rounding could get a score of 8
      return(yellow)
    } else {
      warning(paste("Bad score:", score))
      return(white)
    }
  }
  ##get color for non-sequenced pairs
  getWoSNPColor<-function(score) {
    if(score <= 1) {
      return(red)
    } else if(score > 1 & score <= 2) {
      return(orange)
    } else if(score > 2 & score <= 5) { #with rounding could get a score of 5
      return(yellow)
    } else {
      warning(paste("Bad score:", score))
      return(white)
    }
  }
  
  ###fill in colors
  for(r in 1:nrow(seq)) {
    for(c in 1:ncol(seq)) {
      if(!is.na(seq[r,c])) {
        cs = getSeqColor(seq[r,c])
        if(rank[r,c] == "1") {
          cs = cs + Border(position=c("BOTTOM", "LEFT", "TOP", "RIGHT"), pen="BORDER_MEDIUM") +
            Font(workbook, isBold = T)
        } else {
          cs = cs + Border(position="BOTTOM")
        }
        setCellStyle(cells[[r+2,c+2]], cellStyle = cs)
      } else if(!is.na(nonseq[r,c])) {
        cs = getWoSNPColor(nonseq[r,c])
        if(rank[r,c] == "1*") {
          cs = cs + Border(position=c("BOTTOM", "LEFT", "TOP", "RIGHT"), pen="BORDER_MEDIUM") +
            Font(workbook, isBold = T)
        } else {
          cs = cs + Border(position="BOTTOM")
        }
        setCellStyle(cells[[r+2,c+2]], cellStyle = cs)
      } else {
        setCellStyle(cells[[r+2,c+2]], cellStyle = white)
      }
    }
  }
  
  
  ###legend at bottom
  setCellValue(cells[[2, 2]], "Legend below")
  setCellStyle(cells[[2,2]], cellStyle = CellStyle(workbook, alignment = Alignment(horizontal = "ALIGN_CENTER", vertical = "VERTICAL_CENTER")))
  setCellValue(cells[[length(cases)+4, 1]], "Legend:")
  setCellStyle(cells[[length(cases)+4, 1]], cellStyle = CellStyle(wb=workbook, font=Font(workbook, isBold = T)))
  setCellValue(cells[[length(cases)+4, 2]], "Numbers indicate score")
  setCellValue(cells[[length(cases)+5, 2]], "* indicates no SNP")
  setCellValue(cells[[length(cases)+7, 2]], "Cell color:")
  setCellValue(cells[[length(cases)+8, 2]], "Top ranked source")
  setCellValue(cells[[length(cases)+9, 2]], "Highest likelihood")
  setCellValue(cells[[length(cases)+10, 2]], "Lower likelihood")
  setCellValue(cells[[length(cases)+11, 2]], "Lowest likelihood")
  setCellValue(cells[[length(cases)+12, 2]], "Filtered (not a potential source)")
  setCellStyle(cells[[length(cases)+8, 2]], cellStyle = CellStyle(workbook, 
                                                                  border = Border(position=c("BOTTOM", "LEFT", "TOP", "RIGHT"), 
                                                                                  pen="BORDER_MEDIUM"),
                                                                  font = Font(workbook, isBold = T),
                                                                  alignment = Alignment(horizontal = "ALIGN_CENTER")))
  setCellStyle(cells[[length(cases)+9, 2]], cellStyle = getSeqColor(0))
  setCellStyle(cells[[length(cases)+10, 2]], cellStyle = getSeqColor(2))
  setCellStyle(cells[[length(cases)+11, 2]], cellStyle = getSeqColor(4))
  setCellStyle(cells[[length(cases)+12, 2]], cellStyle = CellStyle(workbook, alignment = Alignment(horizontal = "ALIGN_CENTER")))
  ###fix column widths and print
  autoSizeColumn(sheet, 1:num)
  createFreezePane(sheet, rowSplit = 3, colSplit = 3)
  ###add additional score information
  setCellValue(cells[[length(cases)+9, 3]], "Score or without SNP score of 0-1")
  setCellValue(cells[[length(cases)+10, 3]], "Score of 2-3 or without SNP score of 2")
  setCellValue(cells[[length(cases)+11, 3]], "Score of 4-7 or without SNP score of 3-4")
  saveWorkbook(workbook, fileName)
}
