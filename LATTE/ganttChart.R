##Generate Gantt chart
##Author: Kathryn Winglee

library(xlsx)
library(lubridate)
library(MMWRweek)

##set colors
certCol = "black"
uncertCol = "gray"
ipCol = "red"

##adds the right hand borders to the chart
##workbook, sheet, and cells are the workbook, sheet, and cells to write the header to
##colStart = column to start writing the header
##colEnd = column to stop writing the header
##rowStart = row to start writing the data
##datebreaks = where the breaks between date categories goes
addRightHandHeaderBorders <- function(workbook, cells, colStart, colEnd, rowStart, datebreaks) {
  ##right hand borders
  if(colEnd %in% (datebreaks + colStart)) {
    setCellStyle(cells[[rowStart-1,colEnd]],
                 cellStyle = CellStyle(workbook) + 
                   Border(position = c("TOP", "BOTTOM", "RIGHT", "LEFT"), 
                          pen = c("BORDER_THIN", "BORDER_THIN", "BORDER_THICK", "BORDER_THICK")) +
                   Alignment(wrapText=TRUE, horizontal="ALIGN_CENTER"))
    setCellStyle(cells[[rowStart,colEnd]],
                 cellStyle = CellStyle(workbook) + 
                   Border(position = c("RIGHT", "LEFT", "TOP", "BOTTOM"),
                          pen = c("BORDER_THICK", "BORDER_THICK", rep("BORDER_THIN", 2))) +
                   Alignment(wrapText=TRUE, horizontal="ALIGN_CENTER"))
  } else {
    setCellStyle(cells[[rowStart-1,colEnd]],
                 cellStyle = CellStyle(workbook) + 
                   Border(position = c("TOP", "BOTTOM", "RIGHT"), 
                          pen = c("BORDER_THIN", "BORDER_THIN", "BORDER_THICK")) +
                   Alignment(wrapText=TRUE, horizontal="ALIGN_CENTER"))
    setCellStyle(cells[[rowStart,colEnd]],
                 cellStyle = CellStyle(workbook) + 
                   Border(position = c("RIGHT", "LEFT", "TOP", "BOTTOM"),
                          pen = c("BORDER_THICK", rep("BORDER_THIN", 3))) +
                   Alignment(wrapText=TRUE, horizontal="ALIGN_CENTER"))
    
  }
}

##add borders to header
##workbook, sheet, and cells are the workbook, sheet, and cells to write the header to
##cells = cells where writing results
##colStart = column to start writing the header
##colEnd = column to stop writing the header
##numCols = number of columns total in the sheet/cells
##rowStart = row to start writing the data
##numRows = number of rows total in the sheet/cells
##datebreaks = where the breaks between date categories goes
addBorders <- function(workbook, cells, colStart, colEnd, numCols, rowStart, numRows, datebreaks) {
  ##add borders to date header
  for(c in (colStart+1):(colEnd)) { 
    if(c %in% (datebreaks + colStart)) {
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
    for(c in datebreaks + colStart) {
      setCellStyle(cells[[r,c]],
                   cellStyle = CellStyle(workbook) +
                     Border(position = "LEFT", pen = "BORDER_THICK"))
    }
    if(colEnd %in% (datebreaks + colStart)) { #first of month is last column, so need borders on both sides
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
  for(c in 1:numCols) {
    if(c %in% (datebreaks + colStart)) {
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
  if(colEnd %in% (datebreaks + colStart)) { #first of month is last column, so need borders on both sides
    setCellStyle(cells[[numRows,colEnd]],
                 cellStyle = CellStyle(workbook) +
                   Border(position = c("BOTTOM", "RIGHT", "LEFT"), pen = "BORDER_THICK"))
  } else {
    setCellStyle(cells[[numRows,colEnd]],
                 cellStyle = CellStyle(workbook) +
                   Border(position = c("BOTTOM", "RIGHT"), pen = "BORDER_THICK"))
  }
}

##set up the part of the header containing the dates
##workbook, sheet, and cells are the workbook, sheet, and cells to write the header to
##sheet = sheet where writing results
##cells = cells where writing results
##dates = dates to include in header
##centStyle = style to center the header
##colStart = column to start writing the header
##colEnd = column to stop writing the header
##numCols = number of columns total in the sheet/cells
##rowStart = row to start writing the data
##numRows = number of rows total in the sheet/cells
##progress = optional progress bar (for Rshiny interface)
##prog.loc.incr = how much of the progressbar to increment
setUpDateSectionByDay <- function(workbook, sheet, cells, dates, centStyle,
                             colStart, colEnd, numCols, rowStart, numRows, 
                             progress=NA, prog.loc.incr = 1) {
  ##set up dates
  # years = year(dates)
  months = format(dates, "%B, %Y")#month(dates)
  days = day(dates)
  datebreaks = unique(c(1,which(days==1))) #list of columns that start a month (need a separator) (include the first in the set of dates)
  
  addBorders(workbook, cells, colStart, colEnd, numCols, rowStart, numRows, datebreaks)
  
  if(all(class(progress)!="logical")) {
    progress.value = progress$getValue() # due to rounding errors, may not end at exactly +1
    prog.incr = prog.loc.incr/(length(unique(months)) + length(days)) #increment for progress
    # print(paste("Day increment is", prog.incr, "starting from", progress.value))
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
    if(all(class(progress)!="logical")) {
      progress$set(value = progress$getValue()+prog.incr)
    }
  }
  
  ##day date header
  drow = rowStart #row with day data
  for(i in 1:length(days)) {
    setCellValue(cells[[drow, i+colStart]], days[i])
    if(i %in% datebreaks) {
      setCellStyle(cells[[drow,i+colStart]], 
                   cellStyle = centStyle + 
                     Border(position = c("BOTTOM", "TOP", "RIGHT", "LEFT"), 
                            pen = c(rep("BORDER_THIN", 3), "BORDER_THICK")))
    } else {
      setCellStyle(cells[[drow,i+colStart]], 
                   cellStyle = centStyle + 
                     Border(position = c("BOTTOM", "TOP", "RIGHT", "LEFT"), pen = "BORDER_THIN"))
    }
    if(all(class(progress)!="logical")) {
      progress$set(value = progress$getValue()+prog.incr)
    }
  }
  
  # if(all(class(progress)!="logical")) {
  #   print(paste("After day header before correction: ", progress$getValue()))
  #   # progress$set(value = progress.value+1)
  # }
  
  addRightHandHeaderBorders(workbook, cells, colStart, colEnd, rowStart, datebreaks)
  
  return(datebreaks)
}

##set up the part of the header containing the dates
##workbook, sheet, and cells are the workbook, sheet, and cells to write the header to
##sheet = sheet where writing results
##cells = cells where writing results
##dates = dates to include in header
##centStyle = style to center the header
##colStart = column to start writing the header
##colEnd = column to stop writing the header
##numCols = number of columns total in the sheet/cells
##rowStart = row to start writing the data
##numRows = number of rows total in the sheet/cells
##progress = optional progress bar (for Rshiny interface)
##prog.loc.incr = how much of the progressbar to increment
setUpDateSectionByWeek <- function(workbook, sheet, cells, dates, centStyle,
                                  colStart, colEnd, numCols, rowStart, numRows, 
                                  progress=NA, prog.loc.incr = 1) {
  ##set up dates
  sp = strsplit(dates, "_")
  years = as.numeric(sapply(sp, "[[", 1))
  weeks = as.numeric(sapply(sp, "[[", 2))
  week.start = MMWRweek2Date(years, weeks)
  # months = format(week.start, "%B, %Y")#month, year
  months = as.numeric(format(week.start, "%Y")) #year only; make numeric to get rid of Excel warning
  datebreaks = unique(c(1,which(!duplicated(months)))) #list of columns that start a month (need a separator) (include the first in the set of dates)
  
  addBorders(workbook, cells, colStart, colEnd, numCols, rowStart, numRows, datebreaks)
  
  if(all(class(progress)!="logical")) {
    progress.value = progress$getValue() # due to rounding errors, may not end at exactly +1
    prog.incr = prog.loc.incr/(length(unique(months)) + length(weeks)) #increment for progress
    # print(paste("Week increment is", prog.incr, "starting from", progress.value))
  }
  ##month date header (note that the month is based on the first day; an MMWR week will be grouped in that month, regardless of how many days in that week are actually in that month)
  mrow = rowStart-1 #row with month data
  for(d in unique(months)) {
    cols = which(months==d) + colStart
    setCellValue(cells[[mrow,cols[1]]], d)
    addMergedRegion(sheet, startRow = mrow, startColumn = cols[1], endRow = mrow, endColumn = max(cols))
    setCellStyle(cells[[mrow,cols[1]]], 
                 cellStyle = centStyle + 
                   Border(position = c("BOTTOM", "TOP", "LEFT", "RIGHT"), 
                          pen = c("BORDER_THIN", "BORDER_THIN", "BORDER_THICK", "BORDER_THICK")))
    if(all(class(progress)!="logical")) {
      progress$set(value = progress$getValue()+prog.incr)
    }
  }
  
  ##week date header
  drow = rowStart #row with week data
  for(i in 1:length(weeks)) {
    setCellValue(cells[[drow, i+colStart]], weeks[i])
    if(i %in% datebreaks) {
      setCellStyle(cells[[drow,i+colStart]], 
                   cellStyle = centStyle + 
                     Border(position = c("BOTTOM", "TOP", "RIGHT", "LEFT"), 
                            pen = c(rep("BORDER_THIN", 3), "BORDER_THICK")))
    } else {
      setCellStyle(cells[[drow,i+colStart]], 
                   cellStyle = centStyle + 
                     Border(position = c("BOTTOM", "TOP", "RIGHT", "LEFT"), pen = "BORDER_THIN"))
    }
    if(all(class(progress)!="logical")) {
      progress$set(value = progress$getValue()+prog.incr)
    }
  }
  
  addRightHandHeaderBorders(workbook, cells, colStart, colEnd, rowStart, datebreaks)
  if(all(class(progress)!="logical")) {
    # print(paste("After day header before correction: ", progress$getValue()))
    # progress$set(value = progress.value+1)
  }
  
  return(datebreaks)
}

##set up the part of the header containing the dates
##workbook, sheet, and cells are the workbook, sheet, and cells to write the header to
##sheet = sheet where writing results
##cells = cells where writing results
##dates = dates to include in header
##centStyle = style to center the header
##colStart = column to start writing the header
##colEnd = column to stop writing the header
##numCols = number of columns total in the sheet/cells
##rowStart = row to start writing the data
##numRows = number of rows total in the sheet/cells
##progress = optional progress bar (for Rshiny interface)
##prog.loc.incr = how much of the progressbar to increment
setUpDateSectionByMonth <- function(workbook, sheet, cells, dates, centStyle,
                                   colStart, colEnd, numCols, rowStart, numRows, 
                                   progress=NA, prog.loc.incr = 1) {
  ##set up dates
  sp = strsplit(dates, "-")
  years = as.numeric(sapply(sp, "[[", 2))
  months = as.numeric(sapply(sp, "[[", 1))
  datebreaks = unique(c(1,which(!duplicated(years)))) #list of columns that start a month (need a separator) (include the first in the set of dates)
  
  addBorders(workbook, cells, colStart, colEnd, numCols, rowStart, numRows, datebreaks)
  
  if(all(class(progress)!="logical")) {
    progress.value = progress$getValue() # due to rounding errors, may not end at exactly +1
    prog.incr = prog.loc.incr/(length(unique(years)) + length(months)) #increment for progress
    # print(paste("Month increment is", prog.incr, "starting from", progress.value, "initial increment: ", prog.loc.incr,
    #             "years:", length(unique(years)), "months:", length(months)))
  }
  
  ##year date header
  mrow = rowStart-1 #row with year data
  for(d in unique(years)) {
    cols = which(years==d) + colStart
    setCellValue(cells[[mrow,cols[1]]], d)
    addMergedRegion(sheet, startRow = mrow, startColumn = cols[1], endRow = mrow, endColumn = max(cols))
    setCellStyle(cells[[mrow,cols[1]]], 
                 cellStyle = centStyle + 
                   Border(position = c("BOTTOM", "TOP", "LEFT", "RIGHT"), 
                          pen = c("BORDER_THIN", "BORDER_THIN", "BORDER_THICK", "BORDER_THICK")))
    if(all(class(progress)!="logical")) {
      progress$set(value = progress$getValue()+prog.incr)
    }
  }
  
  ##month date header
  drow = rowStart #row with month data
  for(i in 1:length(months)) {
    setCellValue(cells[[drow, i+colStart]], months[i])
    if(i %in% datebreaks) {
      setCellStyle(cells[[drow,i+colStart]], 
                   cellStyle = centStyle + 
                     Border(position = c("BOTTOM", "TOP", "RIGHT", "LEFT"), 
                            pen = c(rep("BORDER_THIN", 3), "BORDER_THICK")))
    } else {
      setCellStyle(cells[[drow,i+colStart]], 
                   cellStyle = centStyle + 
                     Border(position = c("BOTTOM", "TOP", "RIGHT", "LEFT"), pen = "BORDER_THIN"))
    }
    if(all(class(progress)!="logical")) {
      progress$set(value = progress$getValue()+prog.incr)
    }
  }
  
  addRightHandHeaderBorders(workbook, cells, colStart, colEnd, rowStart, datebreaks)
  # if(all(class(progress)!="logical")) {
  #   print(paste("After month header before correction: ", progress$getValue()))
  #   # progress$set(value = progress.value+1)
  # }
  
  return(datebreaks)
}

##converts the given set of dates to the given time.interval (does nothing if the interval is day)
##time.interval = time interval to use for columns (can be day, week, or month)
##dates = list of dates to convert
convertDates<-function(time.interval, dates) {
  if(time.interval=="month") {
    dates = unique(format(dates, "%m-%Y"))
  } else if(time.interval=="week") { #use MMWR week
    mmwr = MMWRweek(dates)
    dates = unique(paste(mmwr$MMWRyear, mmwr$MMWRweek, sep="_"))
  }
  return(dates)
}

##generates a Gantt chart for the location data
##fileName = name of output file
##loc = location data (formatted from LATTE)
##ip = infectious period data (formatted from LATTE)
##time.interval = time interval to use for columns (can be day, week, or month)
##workbook = if NA, create new workbook, otherwise write results as new sheet in workbook
##save = if true, save the workbook at the end
##progress = optional progress bar (for Rshiny interface)
##returns the workbook used to write the results to
locationGanttChart <- function(fileName, loc, ip, time.interval = "day",
                               workbook=NA, save=T, progress = NA) {
  ##check time.interval
  time.interval = tolower(time.interval)
  if(!time.interval %in% c("day", "week", "month")) {
    return("Invalid time interval: ", time.interval, 
           ". Valid time intervals are: day, week, or month.\r\n")
  }
  
  ##set up
  if(class(workbook)!="jobjRef") {
    workbook = createWorkbook(type = "xlsx")
  }
  
  locations = sort(unique(as.character(loc$Location)))
  
  if(all(class(progress)!="logical")) {
    # print(paste("before updating", time.interval, ":", progress$getValue()))
    progress$set(detail = paste0("writing location Gantt chart: ", time.interval))
    prog.loc = ifelse(time.interval=="day", 3, 2) / length(locations) #how much to increment for each location
  }
  
  ##create Gantt charts as one sheet per location
  for(l in locations) {
    ##set up sheet, dates, and ids
    sub = loc[loc$Location==l,]
    sheet = createSheet(workbook, paste(l, "by", time.interval))
    fulldates = seq(from = min(sub$Start), to = max(sub$End), by = 1) #needed for IP comparison
    dates = convertDates(time.interval, fulldates)
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
    numCells = max(c(numCols,(colStart + 24))) #24 for the legend; number of columns of cells created
    
    rows = createRow(sheet, rowIndex = 1:numRows)
    cells = createCell(rows, colIndex = 1:numCells) 
    
    ##set up styles
    certStyle = CellStyle(workbook) + Fill(foregroundColor = certCol)
    uncertStyle = CellStyle(workbook) + Fill(foregroundColor = uncertCol)
    ipStyle = CellStyle(workbook) + Fill(foregroundColor = ipCol)
    boldStyle = CellStyle(workbook) + Font(workbook, isBold = T)
    centStyle = CellStyle(workbook) + Alignment(wrapText=TRUE, horizontal="ALIGN_CENTER")
    topStyle = CellStyle(workbook) + Border(position = "TOP", pen = "BORDER_THIN")
    
    ##make legend
    setCellValue(cells[[1,2]], 
                 paste("Dates are grouped by", 
                       ifelse(time.interval=="week", 
                              "MMWR week", # (listed in row 4; the month in row 4 is the month the MMWR week starts in)", 
                              time.interval)))
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
    
    ##add border to location styles
    certStyle = certStyle + Border(position = "TOP", pen = "BORDER_THIN")
    uncertStyle = uncertStyle + Border(position = "TOP", pen = "BORDER_THIN")
    
    ##set up header
    if(time.interval=="month") {
      datebreaks = setUpDateSectionByMonth(workbook, sheet, cells, dates, centStyle,
                                           colStart, colEnd, numCols, rowStart, numRows, 
                                           progress, prog.loc/2)
    } else if(time.interval=="week") {
      datebreaks = setUpDateSectionByWeek(workbook, sheet, cells, dates, centStyle,
                                          colStart, colEnd, numCols, rowStart, numRows, 
                                          progress, prog.loc/2)
    } else { #day
      datebreaks = setUpDateSectionByDay(workbook, sheet, cells, dates, centStyle,
                                         colStart, colEnd, numCols, rowStart, numRows, 
                                         progress, prog.loc/2)
    }
    
    # if(all(class(progress)!="logical")) {
    #   prog.id = prog.loc * 1/(length(ids)*2)#how much to increment for each ID; *2 for IP
    #   print(paste("after writing location header for", time.interval, ":", progress$getValue()))
    #   # progress$set(value = progress$getValue()+prog.id, 
    #   #              detail = paste0("writing location Gantt chart: ", time.interval))
    # }
    
    ##fill in dates in location
    for(i in 1:length(ids)) {
      row = rowStart + 2 + (i-1)*3
      setCellValue(cells[[row, 1]], ids[i])
      setCellStyle(cells[[row, 1]], boldStyle + Border(position = "TOP", pen = "BORDER_THIN"))
      
      ##add top separating border
      for(c in 2:numCols) {
        if(c < colStart | !((c-colStart) %in% c(datebreaks, (colEnd-colStart)))) {
          setCellStyle(cells[[row, c]], topStyle) 
        } else if(c==colEnd) {
          if((colEnd-colStart) %in% datebreaks) {
            setCellStyle(cells[[row, c]], 
                         topStyle + Border(position = c("RIGHT", "LEFT", "TOP"), 
                                           pen = c("BORDER_THICK", "BORDER_THICK", "BORDER_THIN"))) 
          } else {
            setCellStyle(cells[[row, c]], 
                         topStyle + Border(position = c("RIGHT", "TOP"), pen = c("BORDER_THICK", "BORDER_THIN"))) 
          }
        } else {
          setCellStyle(cells[[row, c]], 
                       topStyle + Border(position = c("LEFT", "TOP"), pen = c("BORDER_THICK", "BORDER_THIN"))) 
        }
      }
      
      ##for weeks and months, there may be overlaps (e.g. same month in two different intervals)
      ##have certain be overriding (if both certain and uncertain present, mark as certain)
      csub = sub[sub$ID == ids[i],]
      uncert = NA
      cert = NA
      for(r in 1:nrow(csub)) {
        d = seq(csub$Start[r], csub$End[r], by=1)
        d = convertDates(time.interval, d) 
        if(tolower(csub$Confidence[r])=="certain") {
          if(all(is.na(cert))) {
            cert = d
          } else {
            cert = c(cert, d)
          }
        } else {
          if(all(is.na(uncert))) {
            uncert = d
          } else {
            uncert = c(uncert, d)
          }
        }
      }
      list = NA #because know at least one row, know that this will not be NA at the end of this chunk
      strengths = NA
      if(!all(is.na(uncert))) {
        uncert = unique(uncert[!uncert %in% cert])
        list = uncert
        strengths = rep("uncertain", length(uncert))
      }
      if(!all(is.na(cert))) {
        cert = unique(cert)
        if(!all(is.na(list))) {
          list = c(list, cert)
          strengths = c(strengths, rep("certain", length(cert)))
        } else {
          list = cert
          strengths = rep("certain", length(cert))
        }
      }
      
      ##fill in dates in location
      # for(r in 1:nrow(csub)) {
      #   d = seq(csub$Start[r], csub$End[r], by=1)
      #   d = convertDates(time.interval, d) 
      for(stren in unique(strengths)) {
        d = list[strengths==stren]
        cols = which(dates %in% d)
        # if(tolower(csub$Confidence[r])=="certain") {
        if(stren == "certain") {
          sty = certStyle
        } else {
          sty = uncertStyle
        }
        for(c in cols) {
          if(c %in% datebreaks) {
            if((c + colStart) == colEnd) { #need border on both sides if at the end
              tmp = sty + 
                Border(position = c("RIGHT", "LEFT", "TOP"), pen = c("BORDER_THICK", "BORDER_THICK", "BORDER_THIN"))
            } else {
              tmp = sty + Border(position = c("LEFT", "TOP"), pen = c("BORDER_THICK", "BORDER_THIN"))
            }
          } else if((c+colStart) == colEnd) {
            tmp = sty + Border(position = c("RIGHT", "TOP"), pen = c("BORDER_THICK", "BORDER_THIN"))
          } else {
            tmp = sty
          }
          setCellStyle(cells[[row,c+colStart]], tmp)
        }
      }
      # if(all(class(progress)!="logical")) {
      #   prog.id = prog.loc * 1/length(ids)#how much to increment for each ID
      #   progress$set(value = progress$getValue()+prog.id, 
      #                detail = paste0("writing location Gantt chart: ", time.interval))
      # }
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
            if(ipstart < min(fulldates)) {
              setCellValue(cells[[row,2]], format(ipstart, "%m/%d/%Y"))
              setCellStyle(cells[[row,2]], ipStyle)
            }
            if(ipend < min(fulldates)) {
              setCellValue(cells[[row,3]], format(ipend, "%m/%d/%Y"))
              setCellStyle(cells[[row,3]], ipStyle)
            }
            if(ipstart > max(fulldates)) {
              setCellValue(cells[[row,colEnd+1]], format(ipstart, "%m/%d/%Y"))
              setCellStyle(cells[[row,colEnd+1]], ipStyle)
            }
            if(ipend > max(fulldates)) {
              setCellValue(cells[[row,numCols]], format(ipend, "%m/%d/%Y"))
              setCellStyle(cells[[row,numCols]], ipStyle)
            }
            
            ##dates in range
            ipd = seq(ipstart, ipend, by=1)
            ipd = ipd[ipd %in% fulldates]
            if(length(ipd) > 0) {
              ipd = convertDates(time.interval, ipd)
              for(d in ipd) {
                c = which(dates == d)
                if(c %in% datebreaks) {
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
      # if(all(class(progress)!="logical")) {
      #   prog.id = prog.loc * 1/length(ids)#how much to increment for each ID
      #   progress$set(value = progress$getValue()+prog.id, 
      #                detail = paste0("writing location Gantt chart: ", time.interval))
      # }
    }
    
    ##fix column widths
    autoSizeColumn(sheet, 1) #fit ID lengths for first column
    setColumnWidth(sheet, colIndex=c(2:colStart, (colEnd+1):numCols), colWidth=14.3)
    setColumnWidth(sheet, colIndex=(colStart+1):colEnd, colWidth=4.5)
    if(numCols < numCells) { #extra cells created for legend; also need to be made smaller
      setColumnWidth(sheet, colIndex=(numCols+1):numCells, colWidth=4.5)
    }
    
    ##freeze first column
    createFreezePane(sheet, rowSplit = rowStart+1, colSplit = 2)
    
    if(all(class(progress)!="logical")) {
      progress$set(value = progress$getValue()+prog.loc/2, 
                   detail = paste0("writing location Gantt chart: ", time.interval))
    }
  }
  
  # print(paste("after writing location for", time.interval, ":", progress$getValue()))
  
  if(save) {
    saveWorkbook(workbook, fileName)
  }
  return(workbook)
}

##generates a Gantt chart for the infectious period (IP) data
##fileName = name of output file
##ip = infectious period data (formatted from LATTE, assumes no missing IP end or start)
##time.interval = time interval to use for columns (can be day, week, or month)
##workbook = if NA, create new workbook, otherwise write results as new sheet in workbook
##save = if true, save the workbook at the end
##progress = optional progress bar (for Rshiny interface)
##returns the workbook used to write the results to 
ipGanttChart <- function(fileName, ip, time.interval = "week",
                         workbook=NA, save=T, progress = NA) {
  ##check time.interval
  time.interval = tolower(time.interval)
  if(!time.interval %in% c("day", "week", "month")) {
    return("Invalid time interval: ", time.interval, 
           ". Valid time intervals are: day, week, or month.\r\n")
  }

  ##set up sheet, dates, and ids
  if(class(workbook)!="jobjRef") {
    workbook = createWorkbook(type = "xlsx")
  }
  sheet = createSheet(workbook, paste("IP by", time.interval))
  dates = seq(from = min(ip$IPStart), to = max(ip$IPEnd), by = 1)
  dates = convertDates(time.interval, dates)
  ids = sort(unique(ip$ID))
  if(length(dates) > 16384-5) { #maximum number of column in Excel
    warning("More dates than possible Excel columns; Gantt chart will be truncated")
    dates = dates[1:(16384-5)]
  }
  
  if(all(class(progress)!="logical")) {
    progress$set(detail = paste0("writing IP Gantt chart: ", time.interval))
    prog.incr = ifelse(time.interval=="day", 3, 2) #how much to increment for total Gantt chart
    prog.id = prog.incr / length(ids) #how much to increment for each person
    # print(paste("start of IP Gantt:", progress$getValue(), "length:", length(ids), "increment:", prog.id))
  }
  
  ##set up cells
  rowStart = 5 #number of rows in the headers before first row of data
  numRows = length(ids) + rowStart + 2
  colStart = 1 #number of columns with labels before first column of data
  numCols = length(dates) + colStart 
  colEnd = numCols
  numCells = max(c(numCols,(colStart + 7))) #7 for the legend; number of columns of cells created
  
  rows = createRow(sheet, rowIndex = 1:numRows)
  cells = createCell(rows, colIndex = 1:numCells) 
  
  ##set up styles
  ipStyle = CellStyle(workbook) + Fill(foregroundColor = ipCol)
  boldStyle = CellStyle(workbook) + Font(workbook, isBold = T)
  centStyle = CellStyle(workbook) + Alignment(wrapText=TRUE, horizontal="ALIGN_CENTER")
  topStyle = CellStyle(workbook) + Border(position = "TOP", pen = "BORDER_THIN")
  
  ##make legend
  setCellValue(cells[[1,2]], 
               paste("Dates are grouped by", 
                      ifelse(time.interval=="week", 
                             "MMWR week", # (listed in row 4; the month in row 4 is the month the MMWR week starts in)", 
                             time.interval)))
  setCellStyle(cells[[2,colStart+1]], 
               cellStyle = ipStyle + Border(position = c("BOTTOM", "LEFT", "TOP")))
  setCellValue(cells[[2, colStart+2]], "Date during infectious period")
  
  ##add top border
  # ipStyle = ipStyle + Border(position = "TOP", pen = "BORDER_THIN")
  
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
  
  ##set up header
  if(time.interval=="month") {
    datebreaks = setUpDateSectionByMonth(workbook, sheet, cells, dates, centStyle,
                                 colStart, colEnd, numCols, rowStart, numRows, 
                                 progress, prog.incr/2)
  } else if(time.interval=="week") {
    datebreaks = setUpDateSectionByWeek(workbook, sheet, cells, dates, centStyle,
                                 colStart, colEnd, numCols, rowStart, numRows, 
                                 progress, prog.incr/2)
  } else { #day
    datebreaks = setUpDateSectionByDay(workbook, sheet, cells, dates, centStyle,
                                 colStart, colEnd, numCols, rowStart, numRows, 
                                 progress, prog.incr/2)
  }
  
  # if(all(class(progress)!="logical")) {
  #   progress$set(value = progress$getValue()+prog.id, 
  #                detail = paste0("writing IP Gantt chart: ", time.interval))
  # }
  # if(all(class(progress)!="logical")) {
  #   print(paste("after IP Gantt header:", progress$getValue()))
  # }

  ##fill in IP
  for(i in 1:length(ids)) {
    row = rowStart + 1 + i
    setCellValue(cells[[row, 1]], ids[i])
    setCellStyle(cells[[row, 1]], boldStyle)
    ipstart = ip$IPStart[ip$ID==ids[i]]
    ipend = ip$IPEnd[ip$ID==ids[i]]
    
    ##add top separating border
    # for(c in 2:numCols) {
    #   if(c < colStart | !((c-colStart) %in% c(datebreaks, (colEnd-colStart)))) {
    #     setCellStyle(cells[[row, c]], topStyle) 
    #   } else if(c==colEnd) {
    #     if((colEnd-colStart) %in% datebreaks) {
    #       setCellStyle(cells[[row, c]], 
    #                    topStyle + Border(position = c("RIGHT", "LEFT", "TOP"), 
    #                                      pen = c("BORDER_THICK", "BORDER_THICK", "BORDER_THIN"))) 
    #     } else {
    #       setCellStyle(cells[[row, c]], 
    #                    topStyle + Border(position = c("RIGHT", "TOP"), pen = c("BORDER_THICK", "BORDER_THIN"))) 
    #     }
    #   } else {
    #     setCellStyle(cells[[row, c]], 
    #                  topStyle + Border(position = c("LEFT", "TOP"), pen = c("BORDER_THICK", "BORDER_THIN"))) 
    #   }
    # }
    
    
    ##add dates
    ipd = seq(ipstart, ipend, by=1)
    ipd = convertDates(time.interval, ipd)
    for(d in ipd) {
      c = which(dates == d)
      # if(c %in% datebreaks) {
      #   sty = ipStyle + Border(position = c("LEFT", "TOP"), pen = c("BORDER_THICK", "BORDER_THIN"))
      # } else if(c+colStart==colEnd) {
      #   sty = ipStyle + Border(position = c("RIGHT", "TOP"), pen = c("BORDER_THICK", "BORDER_THIN"))
      # } else {
      #   sty = ipStyle
      # }
     
      if(c %in% datebreaks) {
        if((c+colStart)==colEnd) {
          sty = ipStyle + Border(position = c("LEFT", "RIGHT"), pen = "BORDER_THICK")
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
    
    if(all(class(progress)!="logical")) {
      progress$set(value = progress$getValue()+prog.id/2, 
                   detail = paste0("writing IP Gantt chart: ", time.interval))
    }
  }
  
  # if(all(class(progress)!="logical")) {
  #   print(paste("after IP Gantt:", progress$getValue()))
  # }

  ##fix column widths
  autoSizeColumn(sheet, 1) #fit ID lengths for first column
  setColumnWidth(sheet, colIndex=(colStart+1):colEnd, colWidth=4.5)
  if(numCols < numCells) { #extra cells created for legend; also need to be made smaller
    setColumnWidth(sheet, colIndex=(numCols+1):numCells, colWidth=4.5)
  }
  
  ##freeze first column
  createFreezePane(sheet, rowSplit = rowStart+1, colSplit = 2)
  
  if(save) {
    saveWorkbook(workbook, fileName)
  }
  return(workbook)
}
