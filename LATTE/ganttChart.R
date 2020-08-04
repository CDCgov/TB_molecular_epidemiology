##Generate Gantt chart
##Author: Kathryn Winglee

library(xlsx)
library(lubridate)

##set colors
certCol = "black"
uncertCol = "gray"
ipCol = "red"

##set up the part of the header containing the dates
##workbook, sheet, and cells are the workbook, sheet, and cells to write the header to
##dates = dates to include in header
##centStyle = style to center the header
##colStart = column to start writing the header
##colEnd = column to stop writing the header
##numCols = number of columns total in the sheet/cells
##rowStart = row to start writing the data
##numRows = number of rows total in the sheet/cells
setUpDateSection <- function(workbook, sheet, cells, dates, centStyle,
                             colStart, colEnd, numCols, rowStart, numRows) {
  ##set up dates
  # years = year(dates)
  months = format(dates, "%B, %Y")#month(dates)
  days = day(dates)
  monthstart = unique(c(1,which(days==1))) #list of columns that start a month (need a separator) (include the first in the set of dates)
  
  ##add borders to date header
  for(c in (colStart+1):(colEnd)) { 
    if(c %in% (monthstart + colStart)) {
      if(c == colEnd) { #first of month is last column, so need borders on both sides
        setCellStyle(cells[[rowStart-1,c]],
                     cellStyle = CellStyle(workbook) + 
                       Border(position = c("TOP", "BOTTOM", "LEFT", "RIGHT"), 
                              pen = c("BORDER_THIN", "BORDER_THIN", "BORDER_THICK", "BORDER_THICK")))
        setCellStyle(cells[[rowStart,c]],
                     cellStyle = CellStyle(workbook) + 
                       Border(position = c("LEFT", "RIGHT", "TOP", "BOTTOM"),
                              pen = c("BORDER_THICK", "BORDER_THICK", rep("BORDER_THIN", 2))))
      } else {
        setCellStyle(cells[[rowStart-1,c]],
                     cellStyle = CellStyle(workbook) + 
                       Border(position = c("TOP", "BOTTOM", "LEFT"), 
                              pen = c("BORDER_THIN", "BORDER_THIN", "BORDER_THICK")))
        setCellStyle(cells[[rowStart,c]],
                     cellStyle = CellStyle(workbook) + 
                       Border(position = c("LEFT", "RIGHT", "TOP", "BOTTOM"),
                              pen = c("BORDER_THICK", rep("BORDER_THIN", 3))))
      }
    } else {
      setCellStyle(cells[[rowStart-1,c]],
                   cellStyle = CellStyle(workbook) + 
                     Border(position = c("TOP", "BOTTOM"), pen = "BORDER_THIN"))
      setCellStyle(cells[[rowStart,c]],
                   cellStyle = CellStyle(workbook) + 
                     Border(position = c("TOP", "BOTTOM", "LEFT", "RIGHT"), pen = "BORDER_THIN"))
    }
  }
  
  ##add borders between months
  for(r in (rowStart+1):numRows) {
    for(c in monthstart + colStart) {
      setCellStyle(cells[[r,c]],
                   cellStyle = CellStyle(workbook) +
                     Border(position = "LEFT", pen = "BORDER_THICK"))
    }
    if(colEnd %in% (monthstart + colStart)) { #first of month is last column, so need borders on both sides
      setCellStyle(cells[[r,colEnd]],
                   cellStyle = CellStyle(workbook) +
                     Border(position = c("RIGHT", "LEFT"), pen = "BORDER_THICK"))
    } else {
      setCellStyle(cells[[r,colEnd]],
                   cellStyle = CellStyle(workbook) +
                     Border(position = "RIGHT", pen = "BORDER_THICK"))
    }
  }
  
  ##border on bottom
  for(c in (colStart+1):(colEnd)) {
    if(c %in% (monthstart + colStart)) {
      setCellStyle(cells[[numRows,c]],
                   cellStyle = CellStyle(workbook) +
                     Border(position = c("BOTTOM", "LEFT"), pen = "BORDER_THICK"))
      
    } else {
      setCellStyle(cells[[numRows,c]],
                   cellStyle = CellStyle(workbook) +
                     Border(position = "BOTTOM", pen = "BORDER_THICK"))
    }
  }
  setCellStyle(cells[[numRows,colStart+1]],
               cellStyle = CellStyle(workbook) +
                 Border(position = c("BOTTOM", "LEFT"), pen = "BORDER_THICK"))
  if(colEnd %in% (monthstart + colStart)) { #first of month is last column, so need borders on both sides
    setCellStyle(cells[[numRows,colEnd]],
                 cellStyle = CellStyle(workbook) +
                   Border(position = c("BOTTOM", "RIGHT", "LEFT"), pen = "BORDER_THICK"))
  } else {
    setCellStyle(cells[[numRows,colEnd]],
                 cellStyle = CellStyle(workbook) +
                   Border(position = c("BOTTOM", "RIGHT"), pen = "BORDER_THICK"))
  }
  
  ##month date header
  mrow = rowStart-1 #row with month data
  for(d in unique(months)) {
    cols = which(months==d) + colStart
    setCellValue(cells[[mrow,cols[1]]], d)
    addMergedRegion(sheet, startRow = mrow, startColumn = cols[1], endRow = mrow, endColumn = max(cols))
    setCellStyle(cells[[mrow,cols[1]]], 
                 cellStyle = centStyle + 
                   Border(position = c("BOTTOM", "TOP", "LEFT", "RIGHT"), 
                          pen = c("BORDER_THIN", "BORDER_THIN", "BORDER_THICK", "BORDER_THICK")))
  }
  
  ##day date header
  drow = rowStart #row with day data
  for(i in 1:length(days)) {
    setCellValue(cells[[drow, i+colStart]], days[i])
    if(i %in% monthstart) {
      setCellStyle(cells[[drow,i+colStart]], 
                   cellStyle = centStyle + 
                     Border(position = c("BOTTOM", "TOP", "RIGHT", "LEFT"), 
                            pen = c(rep("BORDER_THIN", 3), "BORDER_THICK")))
    } else {
      setCellStyle(cells[[drow,i+colStart]], 
                   cellStyle = centStyle + 
                     Border(position = c("BOTTOM", "TOP", "RIGHT", "LEFT"), pen = "BORDER_THIN"))
    }
  }
  
  ##right hand borders
  if(colEnd %in% (monthstart + colStart)) {
    setCellStyle(cells[[rowStart-1,colEnd]],
                 cellStyle = CellStyle(workbook) + 
                   Border(position = c("TOP", "BOTTOM", "RIGHT", "LEFT"), 
                          pen = c("BORDER_THIN", "BORDER_THIN", "BORDER_THICK", "BORDER_THICK")))
    setCellStyle(cells[[rowStart,colEnd]],
                 cellStyle = CellStyle(workbook) + 
                   Border(position = c("RIGHT", "LEFT", "TOP", "BOTTOM"),
                          pen = c("BORDER_THICK", "BORDER_THICK", rep("BORDER_THIN", 2))))
  } else {
    setCellStyle(cells[[rowStart-1,colEnd]],
                 cellStyle = CellStyle(workbook) + 
                   Border(position = c("TOP", "BOTTOM", "RIGHT"), 
                          pen = c("BORDER_THIN", "BORDER_THIN", "BORDER_THICK")))
    setCellStyle(cells[[rowStart,colEnd]],
                 cellStyle = CellStyle(workbook) + 
                   Border(position = c("RIGHT", "LEFT", "TOP", "BOTTOM"),
                          pen = c("BORDER_THICK", rep("BORDER_THIN", 3))))
    
  }
  
  return(monthstart)
}

##generates a Gantt chart for the location data
##fileName = name of output file
##loc = location data (formatted from LATTE)
##ip = infectious period data (formatted from LATTE)
locationGanttChart <- function(fileName, loc, ip) {
  workbook = createWorkbook(type = "xlsx")
  locations = sort(unique(as.character(loc$Location)))
  
  ##create Gantt charts as one sheet per location
  for(l in locations) {
    ##set up sheet, dates, and ids
    sub = loc[loc$Location==l,]
    sheet = createSheet(workbook, l)
    dates = seq(from = min(sub$Start), to = max(sub$End), by = 1)
    ids = sort(unique(sub$ID))
    if(length(dates) > 16384-5) { #maximum number of column in Excel
      warning("More dates than possible Excel columns; Gantt chart will be truncated")
      dates = dates[1:(16384-5)]
    }
    
    ##set up cells
    rowStart = 5 #number of rows in the headers before first row of data
    numRows = length(ids)*3 + rowStart + 1
    colStart = 3 #number of columns with labels before first column of data
    numCols = length(dates) + colStart + 2 
    colEnd = numCols - 2
    
    rows = createRow(sheet, rowIndex = 1:numRows)
    cells = createCell(rows, colIndex = 1:max(c(numCols,(colStart + 24)))) #24 for the legend
    
    ##set up styles
    certStyle = CellStyle(workbook) + Fill(foregroundColor = certCol)
    uncertStyle = CellStyle(workbook) + Fill(foregroundColor = uncertCol)
    ipStyle = CellStyle(workbook) + Fill(foregroundColor = ipCol)
    boldStyle = CellStyle(workbook) + Font(workbook, isBold = T)
    centStyle = CellStyle(workbook) + Alignment(wrapText=TRUE, horizontal="ALIGN_CENTER")
    
    ##make legend
    setCellStyle(cells[[2,colStart+1]], 
                 cellStyle = certStyle + Border(position = c("BOTTOM", "LEFT", "TOP")))
    setCellValue(cells[[2, colStart+2]], "Date in location (certain)")
    setCellStyle(cells[[2,colStart+10]], 
                 cellStyle = uncertStyle + Border(position = c("BOTTOM", "LEFT", "TOP")))
    setCellValue(cells[[2, colStart+11]], "Date in location (uncertain)")
    setCellStyle(cells[[2,colStart+18]], 
                 cellStyle = ipStyle + Border(position = c("BOTTOM", "LEFT", "TOP")))
    setCellValue(cells[[2, colStart+19]], "Date during infectious period")
    
    for(c in c((colStart+2):(colStart+7), (colStart+11):(colStart+16), (colStart+19):(colStart+24))) {
      if(c %in% c(colStart+7, colStart+16, colStart+24)) {
        setCellStyle(cells[[2,c]], 
                     cellStyle = CellStyle(workbook) + Fill(foregroundColor = "white") +
                       Border(position = c("BOTTOM", "TOP", "RIGHT")))
      } else {
        setCellStyle(cells[[2,c]], 
                     cellStyle = CellStyle(workbook) + Fill(foregroundColor = "white") +
                       Border(position = c("BOTTOM", "TOP")))
      }
    }
    
    ##non-date header
    setCellValue(cells[[4,2]], "Extended dates")
    addMergedRegion(sheet, startRow = 4, endRow = 4, startColumn = 2, endColumn = 3)
    setCellStyle(cells[[4,2]], centStyle)
    setCellValue(cells[[4,colEnd+1]], "Extended dates")
    addMergedRegion(sheet, startRow = 4, endRow = 4, startColumn = colEnd+1, endColumn = numCols)
    setCellStyle(cells[[4,colEnd+1]], centStyle)
    setCellValue(cells[[5,1]], "ID")
    setCellStyle(cells[[5,1]], boldStyle)
    setCellValue(cells[[5,2]], "IP start date")
    setCellStyle(cells[[5,2]], centStyle)
    setCellValue(cells[[5,3]], "IP end date")
    setCellStyle(cells[[5,3]], centStyle)
    setCellValue(cells[[5,colEnd+1]], "IP start date")
    setCellStyle(cells[[5,colEnd+1]], centStyle)
    setCellValue(cells[[5,numCols]], "IP end date")
    setCellStyle(cells[[5,numCols]], centStyle)
    
    ##set up dates
    monthstart = setUpDateSection(workbook, sheet, cells, dates, centStyle,
                                  colStart, colEnd, numCols, rowStart, numRows)
    
    # ##set up dates
    # # years = year(dates)
    # months = format(dates, "%B, %Y")#month(dates)
    # days = day(dates)
    # monthstart = unique(c(1,which(days==1))) #list of columns that start a month (need a separator) (include the first in the set of dates)
    # 
    # ##add borders to date header
    # for(c in (colStart+1):(numCols-3)) {
    #   if(c %in% monthstart) {
    #     setCellStyle(cells[[4,c]],
    #                  cellStyle = CellStyle(workbook) + 
    #                    Border(position = c("TOP", "BOTTOM", "LEFT"), 
    #                           pen = c("BORDER_THIN", "BORDER_THIN", "BORDER_THICK")))
    #     setCellStyle(cells[[5,c]],
    #                  cellStyle = CellStyle(workbook) + 
    #                    Border(position = c("LEFT", "RIGHT", "TOP", "BOTTOM"),
    #                           pen = c("BORDER_THICK", rep("BORDER_THIN", 3))))
    #   } else {
    #     setCellStyle(cells[[4,c]],
    #                  cellStyle = CellStyle(workbook) + 
    #                    Border(position = c("TOP", "BOTTOM"), pen = "BORDER_THIN"))
    #     setCellStyle(cells[[5,c]],
    #                  cellStyle = CellStyle(workbook) + 
    #                    Border(position = c("TOP", "BOTTOM", "LEFT", "RIGHT"), pen = "BORDER_THIN"))
    #   }
    # }
    # 
    # ##add borders between months
    # for(r in rowStart:numRows) {
    #   for(c in monthstart + colStart) {
    #     setCellStyle(cells[[r,c]],
    #                  cellStyle = CellStyle(workbook) +
    #                    Border(position = "LEFT", pen = "BORDER_THICK"))
    #   }
    #   setCellStyle(cells[[r,colEnd]],
    #                cellStyle = CellStyle(workbook) +
    #                  Border(position = "RIGHT", pen = "BORDER_THICK"))
    # }
    # 
    # ##border on bottom
    # for(c in (colStart+1):(colEnd)) {
    #   if(c %in% (monthstart + colStart)) {
    #     setCellStyle(cells[[numRows,c]],
    #                  cellStyle = CellStyle(workbook) +
    #                    Border(position = c("BOTTOM", "LEFT"), pen = "BORDER_THICK"))
    #     
    #   } else {
    #     setCellStyle(cells[[numRows,c]],
    #                  cellStyle = CellStyle(workbook) +
    #                    Border(position = "BOTTOM", pen = "BORDER_THICK"))
    #   }
    # }
    # setCellStyle(cells[[numRows,colStart+1]],
    #              cellStyle = CellStyle(workbook) +
    #                Border(position = c("BOTTOM", "LEFT"), pen = "BORDER_THICK"))
    # setCellStyle(cells[[numRows,colEnd]],
    #              cellStyle = CellStyle(workbook) +
    #                Border(position = c("BOTTOM", "RIGHT"), pen = "BORDER_THICK"))
    # 
    # ##month date header
    # mrow = 4 #row with month data
    # for(d in unique(months)) {
    #   cols = which(months==d) + colStart
    #   setCellValue(cells[[mrow,cols[1]]], d)
    #   addMergedRegion(sheet, startRow = mrow, startColumn = cols[1], endRow = mrow, endColumn = max(cols))
    #   setCellStyle(cells[[mrow,cols[1]]], 
    #                cellStyle = centStyle + 
    #                  Border(position = c("BOTTOM", "TOP", "LEFT", "RIGHT"), 
    #                         pen = c("BORDER_THIN", "BORDER_THIN", "BORDER_THICK", "BORDER_THICK")))
    # }
    # 
    # ##day date header
    # drow = 5 #row with day data
    # for(i in 1:length(days)) {
    #   setCellValue(cells[[drow, i+colStart]], days[i])
    #   if(days[i]==1) {
    #     setCellStyle(cells[[drow,i+colStart]], 
    #                  cellStyle = centStyle + 
    #                    Border(position = c("BOTTOM", "TOP", "RIGHT", "LEFT"), 
    #                           pen = c(rep("BORDER_THIN", 3), "BORDER_THICK")))
    #   } else {
    #     setCellStyle(cells[[drow,i+colStart]], 
    #                  cellStyle = centStyle + 
    #                    Border(position = c("BOTTOM", "TOP", "RIGHT", "LEFT"), pen = "BORDER_THIN"))
    #   }
    # }
    # 
    # ##right hand borders
    # setCellStyle(cells[[4,colEnd]],
    #              cellStyle = CellStyle(workbook) + 
    #                Border(position = c("TOP", "BOTTOM", "RIGHT"), 
    #                       pen = c("BORDER_THIN", "BORDER_THIN", "BORDER_THICK")))
    # setCellStyle(cells[[5,colEnd]],
    #              cellStyle = CellStyle(workbook) + 
    #                Border(position = c("RIGHT", "LEFT", "TOP", "BOTTOM"),
    #                       pen = c("BORDER_THICK", rep("BORDER_THIN", 3))))
    
    ##fill in dates in location
    for(i in 1:length(ids)) {
      row = rowStart + 2 + (i-1)*3
      setCellValue(cells[[row, 1]], ids[i])
      setCellStyle(cells[[row, 1]], boldStyle)
      csub = sub[sub$ID == ids[i],]
      for(r in 1:nrow(csub)) {
        d = seq(csub$Start[r], csub$End[r], by=1)
        cols = which(dates %in% d)
        if(tolower(csub$Confidence[r])=="certain") {
          sty = certStyle
        } else {
          sty = uncertStyle
        }
        for(c in cols) {
          if(c %in% monthstart) {
            if((c + colStart) == colEnd) { #need border on both sides if at the end
              tmp = sty + Border(position = c("RIGHT", "LEFT"), pen = "BORDER_THICK")
            } else {
              tmp = sty + Border(position = "LEFT", pen = "BORDER_THICK")
            }
          } else if((c+colStart) == colEnd) {
            tmp = sty + Border(position = "RIGHT", pen = "BORDER_THICK")
          } else {
            tmp = sty
          }
          setCellStyle(cells[[row,c+colStart]], tmp)
        }
      }
    }
    
    ##fill in IP
    if(!all(is.na(ip))) {
      for(i in 1:length(ids)) {
        if(ids[i] %in% ip$ID) {
          row = rowStart + 3 + (i-1)*3
          ipstart = ip$IPStart[ip$ID==ids[i]]
          ipend = ip$IPEnd[ip$ID==ids[i]]
          
          if(!is.na(ipstart) & !is.na(ipend)) {
            ##extended dates
            if(ipstart < min(dates)) {
              setCellValue(cells[[row,2]], format(ipstart, "%m/%d/%Y"))
              setCellStyle(cells[[row,2]], ipStyle)
            }
            if(ipend < min(dates)) {
              setCellValue(cells[[row,3]], format(ipend, "%m/%d/%Y"))
              setCellStyle(cells[[row,3]], ipStyle)
            }
            if(ipstart > max(dates)) {
              setCellValue(cells[[row,colEnd+1]], format(ipstart, "%m/%d/%Y"))
              setCellStyle(cells[[row,colEnd+1]], ipStyle)
            }
            if(ipend > max(dates)) {
              setCellValue(cells[[row,numCols]], format(ipend, "%m/%d/%Y"))
              setCellStyle(cells[[row,numCols]], ipStyle)
            }
            
            ##dates in range
            ipd = seq(ipstart, ipend, by=1)
            ipd = ipd[ipd %in% dates]
            if(length(ipd) > 0) {
              for(d in ipd) {
                c = which(dates == d)
                if(c %in% monthstart) {
                  if((c + colStart) == colEnd) { #need border on both sides if at the end
                    sty = ipStyle + Border(position = c("RIGHT", "LEFT"), pen = "BORDER_THICK")
                  } else {
                    sty = ipStyle + Border(position = "LEFT", pen = "BORDER_THICK")
                  }
                } else if((c+colStart)==colEnd) {
                  sty = ipStyle + Border(position = "RIGHT", pen = "BORDER_THICK")
                } else {
                  sty = ipStyle
                }
                setCellStyle(cells[[row,c+colStart]], sty)
              }
            }
          }
        }
      }
    }
    
    ##fix column widths
    autoSizeColumn(sheet, 1) #fit ID lengths for first column
    setColumnWidth(sheet, colIndex=c(2:colStart, (colEnd+1):numCols), colWidth=14.3)
    setColumnWidth(sheet, colIndex=(colStart+1):colEnd, colWidth=4.5)
    
    ##freeze first column
    createFreezePane(sheet, rowSplit = rowStart+1, colSplit = 2)
  }
  
  saveWorkbook(workbook, fileName)
}

##generates a Gantt chart for the infectious period (IP) data
##fileName = name of output file
##ip = infectious period data (formatted from LATTE)
ipGanttChart <- function(fileName, ip) {
  workbook = createWorkbook(type = "xlsx")
  
  ##set up sheet, dates, and ids
  sheet = createSheet(workbook, "IP")
  dates = seq(from = min(ip$IPStart), to = max(ip$IPEnd), by = 1)
  ids = sort(unique(ip$ID))
  if(length(dates) > 16384-5) { #maximum number of column in Excel
    warning("More dates than possible Excel columns; Gantt chart will be truncated")
    dates = dates[1:(16384-5)]
  }
  
  ##set up cells
  rowStart = 5 #number of rows in the headers before first row of data
  numRows = length(ids) + rowStart + 2
  colStart = 1 #number of columns with labels before first column of data
  numCols = length(dates) + colStart + 2 
  colEnd = numCols - 2
  
  rows = createRow(sheet, rowIndex = 1:numRows)
  cells = createCell(rows, colIndex = 1:max(c(numCols,(colStart + 7)))) #7 for the legend
  
  ##set up styles
  ipStyle = CellStyle(workbook) + Fill(foregroundColor = ipCol)
  boldStyle = CellStyle(workbook) + Font(workbook, isBold = T)
  centStyle = CellStyle(workbook) + Alignment(wrapText=TRUE, horizontal="ALIGN_CENTER")
  
  ##make legend
  setCellStyle(cells[[2,colStart+1]], 
               cellStyle = ipStyle + Border(position = c("BOTTOM", "LEFT", "TOP")))
  setCellValue(cells[[2, colStart+2]], "Date during infectious period")
  
  for(c in (colStart+2):(colStart+7)) {
    if(c == colStart+7) {
      setCellStyle(cells[[2,c]], 
                   cellStyle = CellStyle(workbook) + Fill(foregroundColor = "white") +
                     Border(position = c("BOTTOM", "TOP", "RIGHT")))
    } else {
      setCellStyle(cells[[2,c]], 
                   cellStyle = CellStyle(workbook) + Fill(foregroundColor = "white") +
                     Border(position = c("BOTTOM", "TOP")))
    }
  }
  
  ##column headers
  setCellValue(cells[[5,1]], "ID")
  setCellStyle(cells[[5,1]], boldStyle)
  
  ##set up dates
  monthstart = setUpDateSection(workbook, sheet, cells, dates, centStyle,
                                colStart, colEnd, numCols, rowStart, numRows)
  
  ##fill in IP
  for(i in 1:length(ids)) {
    row = rowStart + 1 + i
    setCellValue(cells[[row, 1]], ids[i])
    setCellStyle(cells[[row, 1]], boldStyle)
    ipstart = ip$IPStart[ip$ID==ids[i]]
    ipend = ip$IPEnd[ip$ID==ids[i]]
    
    if(!is.na(ipstart) & !is.na(ipend)) 
      ##dates in range
      ipd = seq(ipstart, ipend, by=1)
    for(d in ipd) {
      c = which(dates == d)
      if(c %in% monthstart) {
        sty = ipStyle + Border(position = "LEFT", pen = "BORDER_THICK")
      } else if(c+colStart==colEnd) {
        sty = ipStyle + Border(position = "RIGHT", pen = "BORDER_THICK")
      } else {
        sty = ipStyle
      }
      setCellStyle(cells[[row,c+colStart]], sty)
    }
  }
  
  ##fix column widths
  autoSizeColumn(sheet, 1) #fit ID lengths for first column
  setColumnWidth(sheet, colIndex=(colStart+1):colEnd, colWidth=4.5)
  
  ##freeze first column
  createFreezePane(sheet, rowSplit = rowStart+1, colSplit = 2)
  
  saveWorkbook(workbook, fileName)
}
