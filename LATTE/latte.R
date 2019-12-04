##Identify overlap in time and location between people

source("../sharedFunctions.R") ##need convertToDate
library(xlsx)

afterIPstring = "after IP end" #string used to indicate that an overlap occurs after an IP end
noIPstring = "no IP available" #string used to indicate when no IP is available for a person
defaultCut = 2 #default time cutoff
defaultLogName = "LATTE_log.txt" #default name of log file

##function that takes the given data frame and corrects the ID column capitalization to be uniform across the program
fixIDName <- function(df) {
  if(any(!is.na(df))) {
    names(df)[grep("ID", names(df), ignore.case = F)] = "ID"
  }
  return(df)
}

##function that takes the given data frame and corrects the column location name capitalization
fixLocNames <- function(df, log) {
  names(df) = removeWhitespacePeriods(names(df))
  df = df[!apply(df, 1, function(x) all(is.na(x))),
          !apply(df, 2, function(x) all(is.na(x)))] #remove rows and columns that are all NA
  names(df)[grep("^location$", names(df), ignore.case = T)] = "Location"
  names(df)[grepl("start", names(df), ignore.case = T) & 
              !grepl("ip", names(df), ignore.case = T)] = "Start"
  names(df)[(grepl("end", names(df), ignore.case = T) | grepl("stop", names(df), ignore.case = T)) &
              !grepl("ip", names(df), ignore.case = T)] = "End"
  names(df)[grepl("confidence", names(df), ignore.case = T) | 
              grepl("probability", names(df), ignore.case = T) | 
              grepl("strength", names(df), ignore.case = T)] = "Confidence"
  reqCols = c("ID", "Location", "Start", "End", "Confidence")
  if(!all(reqCols %in% names(df))) {
    cat("Location table is missing these required columns: ", file = log, append = T)
    cat(paste(reqCols[!reqCols %in% names(df)], collapse=", "), file = log, append = T)
    cat(". Analysis has been stopped.\r\n", file = log, append = T)
    stop("Location table is missing required columns. Columns must include: ID, Location, Start, End, Confidence.")
  }
  return(df)
}

##check infectious period column names
##check for duplicates; if same isolate has two different IP dates, give warning and take the earlier
fixIPnames <- function(df, log) {
  if(any(!is.na(df))) {
    names(df) = removeWhitespacePeriods(names(df))
    df = df[!apply(df, 1, function(x) all(is.na(x))),
            !apply(df, 2, function(x) all(is.na(x)))] #remove rows and columns that are all NA
    ##start
    col = grepl("ip[. ]*start", names(df), ignore.case = T) | grepl("infectious[. ]*period[. ]*start", names(df), ignore.case = T)
    names(df)[col] = "IPStart"
    ##end
    col = grepl("ip[. ]*end", names(df), ignore.case = T) | grepl("infectious[. ]*period[. ]*end", names(df), ignore.case = T) |
      grepl("ip[. ]*stop", names(df), ignore.case = T) | grepl("infectious[. ]*period[. ]*stop", names(df), ignore.case = T)
    names(df)[col] = "IPEnd"
    if(sum(names(df) == "IPStart") != 1 | sum(names(df) == "IPEnd") != 1) {
      cat("Invalid number of IP start or end columns, so IP will not be included\r\n", file = log, append = T)
      return(NA)
    }
    
    ##check for duplicates and clean up if present
    df$ID = as.character(df$ID)
    dup = duplicated(df$ID)
    if(any(dup)) {
      ids = unique(df$ID[dup])
      cat(paste("The following IDs have multiple provided IPs. The first IP in the table will be used: ", 
                paste(ids, collapse = ", ") ,"\r\n", file = log, append = T))
      for(i in ids) {
        rows = which(df$ID==i)
        df = df[-rows[-1],]
      }
    }
  }
  reqcols = c("ID", "IPStart", "IPEnd")
  if(!all(reqcols %in% names(df))) {
    cat("Infectious period inputs need column for ID, IP start and IP end, but some or all of this data is missing so IP will not be included\r\n", file = log, append = T)
    df=NA
  }
  df = df[,names(df) %in% reqcols]
  return(df)
}

##look for overlap between the date ranges start1-end, start2-end
##if no overlap, return NA, else return c(overlap start, overlap end)
getOverlap <- function(start1, end1, start2, end2) {
  if((start1 >= start2 & start1 <= end2) |
     (start2 >= start1 & start2 <= end1)) { #overlap
    ##calculate actual overlap
    if(start1 >= start2) {
      olstart = start1
    } else {
      olstart = start2
    }
    if(end1 <= end2) {
      olend = end1
    } else {
      olend = end2
    }
    return(c(olstart, olend))
  } else {
    return(NA)
  }
}

##function that returns the number of days the infectious period from ip for case overlaps the time range represented by start and end
overlapIP <- function(start, end, ip, case) {
  val = noIPstring
  if(!all(is.na(ip))) {
    if(case %in% ip$ID) {
      if(!is.na(ip$IPStart[ip$ID==case]) & !is.na(ip$IPEnd[ip$ID==case])) {
        ol = getOverlap(start, end, ip$IPStart[ip$ID==case], ip$IPEnd[ip$ID==case])
        if(all(is.na(ol))) {
          if(ip$IPEnd[ip$ID==case] < start) {
            val = afterIPstring
          } else {
            val = 0
          }
        } else {
          val = as.numeric(ol[2]-ol[1])+1
        }
      }
    }
  }
  return(val)
}

##cleans up the column names of the given data frame df for output
namesForOutput <- function(df) {
  #ID1, ID2, location, strength, location fine
  names(df)[names(df)=="NumDaysOverlap"] = "Number of overlapping days"
  names(df)[names(df)=="OverlapStart"] = "Overlap start date"
  names(df)[names(df)=="OverlapEnd"] = "Overlap end date"
  names(df)[names(df)=="OverlapID1IP"] = "Number of days of overlap in ID1 IP" #"Number of days ID2 overlaps in ID1's IP"
  names(df)[names(df)=="OverlapID2IP"] = "Number of days of overlap in ID2 IP"
  names(df)[names(df)=="IPStart"] = "Infectious period start"
  names(df)[names(df)=="IPEnd"] = "Infectious period end"
  names(df)[names(df)=="NumCert"] = "Total number of days of certain overlap"
  names(df)[names(df)=="NumTot"] = "Total number of overlapped days"
  names(df)[names(df)=="Start"] = "Location start"
  names(df)[names(df)=="End"] = "Location end"
  names(df)[names(df)=="numCertOverlap"] = "Total number of days of certain overlap with another person"
  names(df)[names(df)=="numTotOverlap"] = "Total number of overlapped days with another person"
  names(df)[names(df)=="numCertIPOverap"] = "Total number of days of certain overlap with another person during their IP"
  names(df)[names(df)=="numTotIPOverlap"] = "Total number of overlapped days with another person during their IP"
  
  ##format dates
  for(c in 1:ncol(df)) {
    if(class(df[,c])=="Date") {
      df[,c] = format(df[,c], format="%m/%d/%Y")
    }
  }
  return(df)
}

##writes the data frame df to the Excel file or workbook and sheet
##df = data to write
##if fileName is not NA, generate the workbook first
##workbook should be provided if no fileName is given, otherwise the value will be ignored
##returns the workbook in case additional sheets need to be written
##if wrapHeader is true, set column widths to the longest length and wrap the header (rather than using auto, which will not wrap the header)
##save = if true, save the workbook
writeExcelTable<-function(fileName, workbook=NA, sheetName="Sheet1", df, wrapHeader=F, save = T) {
  df = namesForOutput(df)
  if(class(workbook)!="jobjRef") {
    workbook = createWorkbook()
  }
  sheet = createSheet(workbook, sheetName = sheetName)
  rows = createRow(sheet, rowIndex = 1:(nrow(df)+1))
  cells = createCell(rows, colIndex = 1:ncol(df))
  
  ##header
  headstyle = CellStyle(workbook) + Fill(foregroundColor = rgb(red = 192, green = 192, blue = 192, maxColorValue = 255))
  if(wrapHeader) {
    headstyle = headstyle + Alignment(wrapText = T)
  }
  for(c in 1:ncol(df)) {
    setCellValue(cells[[1,c]], names(df)[c])
    setCellStyle(cells[[1,c]], headstyle)
  }
  ##write data
  for(r in 1:nrow(df)) {
    for(c in 1:ncol(df)) {
      if(!is.na(df[r,c])) {
        if(as.character(df[r,c])!="") {
          if(all(grepl("[0-9.]", strsplit(as.character(df[r,c]), "")[[1]]))) {
            setCellValue(cells[[r+1,c]], as.numeric(as.character(df[r,c])))
          } else {
            setCellValue(cells[[r+1,c]], df[r,c])
          }
        }
      }
    }
  }
  ##set column widths
  if(wrapHeader) {
    cwidth = (sapply(1:ncol(df), function(c){max(nchar(as.character(df[,c]), keepNA=F))})+5)#*256 #width is 1/256 of char
    minWidth = 11# 10*256 #minimum width for column
    cwidth[cwidth < minWidth] = minWidth
    for(c in 1:length(cwidth)) {
      setColumnWidth(sheet = sheet, colIndex = c, colWidth = cwidth[c])
    }
  } else {
    autoSizeColumn(sheet, 1:ncol(df))
  }
  ##add filter
  r1 = sheet$getRow(1L)
  lastcol = r1$getCell(as.integer(ncol(df)-1))
  addAutoFilter(sheet = sheet, cellRange =paste("A1:", lastcol$getReference(), sep=""))
  if(save) {
    saveWorkbook(workbook, fileName)
  }
  return(workbook)
}

###returns a table of epi links, using:
##res = results table containing all overlaps
##cutoff = time cutoff; cases must overlap for at least this many days to be a definite or probable epi link
getEpiLinks <-function(res, cutoff) {
  epi = NA
  allCombo = strsplit(unique(paste(res$ID1, res$ID2, sep=",.")), ",.")
  for(c in allCombo) {
    sub = res[res$ID1==c[1] & res$ID2==c[2],]
    if(nrow(sub) > 0) {
      temp = data.frame(ID1 = c[1],
                        ID2 = c[2],
                        Strength = NA,
                        Location = paste(sort(unique(sub$Location)), collapse = ","),
                        NumCert = sum(sub$NumDaysOverlap[sub$Strength=="certain"]),
                        NumTot = sum(sub$NumDaysOverlap))
      if(temp$NumCert >= cutoff) {
        temp$Strength = "definite"
      } else if(temp$NumTot >= cutoff) {
        temp$Strength = "probable"
      } else {
        temp$Strength = "possible"
      } 
      
      if(all(is.na(epi))) {
        epi = temp
      } else {
        epi = rbind(epi, temp, stringsAsFactors=F)
      }
    } 
  }
  return(epi)
}

###returns a table of IP epi links, using:
##res = results table containing all overlaps
##cutoff = time cutoff; cases must overlap for at least this many days to be a definite or probable epi link
##removeAfter = if true, remove overlaps that occur after the IP of either case
getIPEpiLinks <- function(res, cutoff, removeAfter) {
  epi = NA
  allCombo = strsplit(unique(paste(res$ID1, res$ID2, sep=",.")), ",.")
  for(c in allCombo) {
    sub = res[res$ID1==c[1] & res$ID2==c[2],]
    if(nrow(sub) > 0 & removeAfter) { #remove overlaps that occur after the IP of either ID
      sub = sub[sub$OverlapID1IP != afterIPstring & sub$OverlapID2IP != afterIPstring,]
    } else {
      sub$OverlapID1IP[sub$OverlapID1IP == afterIPstring] = NA
      sub$OverlapID2IP[sub$OverlapID2IP == afterIPstring] = NA
    }
    if(nrow(sub) > 0) {
      sub$OverlapID1IP[sub$OverlapID1IP==noIPstring] = NA
      sub$OverlapID2IP[sub$OverlapID2IP==noIPstring] = NA
      sub$OverlapID1IP = as.numeric(as.character(sub$OverlapID1IP))
      sub$OverlapID2IP = as.numeric(as.character(sub$OverlapID2IP))
      temp = data.frame(ID1 = c[1],
                        ID2 = c[2],
                        Strength = NA,
                        Location = paste(sort(unique(sub$Location)), collapse = ","))
      numCert1 = sum(sub$OverlapID1IP[sub$Strength=="certain"], na.rm = T)
      numCert2 = sum(sub$OverlapID2IP[sub$Strength=="certain"], na.rm = T)
      numTot1 = sum(sub$OverlapID1IP, na.rm = T)
      numTot2 = sum(sub$OverlapID2IP, na.rm = T)
      temp$NumCert = numCert1 + numCert2
      temp$NumTot = numTot1 + numTot2
      if(numCert1 >= cutoff | numCert2 >= cutoff) {
        temp$Strength = "definite"
      } else if(numTot1 >= cutoff | numTot2 >= cutoff) {
        temp$Strength = "probable"
      } else {
        temp = NA
      }
      
      if(!all(is.na(temp))) {
        if(all(is.na(epi))) {
          epi = temp
        } else {
          epi = rbind(epi, temp, stringsAsFactors=F)
        }
      }
    } 
  }
  return(epi)
}

##run the LATTE algorithm
##loc = table of locations, expected columns = ID, location, start, end, confidence
##ip = optional table of infectious periods, expected columns = ID, IPStart, IPEnd
##cutoff = overlap must be more than cutoff number of days to be considered for an epi or IP epi link
##ipEpiLink = if true, calculate an IP epi link, otherwise calculate an epi link
##removeAfter = if true, remove overlaps that occur after the IP of either case (for IP epi links)
##log = log file name (where messages will be written)
##progress = progress bar for R Shiny interface (NA if not running through interface)
latte <- function(loc, ip = NA, cutoff = defaultCut, ipEpiLink = F, removeAfter = F, progress = NA, log = defaultLogName) {
  ##set up log
  cat("LATTE analysis\r\n", file = log)
  
  ###clean up headers
  loc = fixIDName(loc)
  loc = fixLocNames(loc, log)
  
  ###if location end is empty, assume it is one day and assign location start
  loc[] = lapply(loc, as.character)
  loc$End[is.na(loc$End) | loc$End==""] = loc$Start[is.na(loc$End) | loc$End==""] 
  
  ###convert dates
  loc$Start = convertToDate(loc$Start)
  loc$End = convertToDate(loc$End)
  
  ##fix strengths
  loc$Confidence = tolower(loc$Confidence)
  if(any(loc$Confidence!="certain" & loc$Confidence!="uncertain")) {
    cat("The following rows in the location table have invalid confidence values and will be set to uncertain: ", file = log, append = T)
    # cat(loc[loc$Confidence!="certain" & loc$Confidence!="uncertain",], file = log, append = T)
    cat(row.names(loc)[loc$Confidence!="certain" & loc$Confidence!="uncertain"], file = log, append = T)
    cat("\r\n", file = log, append = T)
    cat("\tValid confidence values are: certain, uncertain\r\n", file = log, append = T)
    loc$Confidence[loc$Confidence!="certain" & loc$Confidence!="uncertain"] = "uncertain"
  }
  
  ###test for bad dates
  if(any(is.na(loc$Start) | is.na(loc$End))) {
    cat("The following rows in the location table have invalid dates and will be removed from analysis: ", file = log, append = T)
    # cat(loc[is.na(loc$Start) | is.na(loc$End),], file = log, append = T)
    cat(row.names(loc)[is.na(loc$Start) | is.na(loc$End)], file = log, append = T)
    cat("\r\n", file = log, append = T)
    loc = loc[!is.na(loc$Start) & !is.na(loc$End),]
  }
  
  ###test for bad date ranges
  if(any(loc$Start > loc$End)) {
    cat("The following rows in the location table have an end date before the start date and will be removed from analysis: ", file = log, append = T)
    # cat(loc[loc$Start > loc$End,], file = log, append = T)
    cat(row.names(loc)[loc$Start > loc$End], file = log, append = T)
    cat("\r\n", file = log, append = T)
    loc = loc[loc$Start <= loc$End,]
  }
  
  cases = sort(unique(as.character(loc$ID)))
  
  ##clean IP
  if(!all(is.na(ip))) {
    ip = fixIDName(ip)
    ip = ip[ip$ID %in% cases,]
    ip = fixIPnames(ip, log)
  }
  if(!all(is.na(ip))) {
    ip$IPStart = convertToDate(ip$IPStart)
    ip$IPEnd = convertToDate(ip$IPEnd)
    ##message if missing an IP
    if(any(!cases %in% ip$ID)) {
      cat("The following people are in the location table but not in the IP table so will have no IP in analysis: ", file = log, append = T)
      cat(paste(as.character(cases[!cases %in% ip$ID]), collapse=", "), file = log, append = T)
      cat("\r\n", file = log, append = T)
    }
    ##test for bad IPs:
    if(any(is.na(ip$IPStart) | is.na(ip$IPEnd))) {
      cat("The following people in the IP table have a missing IP start or end so will have no IP in analysis: ", file = log, append = T)
      cat(paste(as.character(ip$ID[is.na(ip$IPStart) | is.na(ip$IPEnd)]), collapse=", "), file = log, append = T)
      cat("\r\n", file = log, append = T)
    }
    ip = ip[!is.na(ip$IPStart),]
    ip = ip[!is.na(ip$IPEnd),]
    if(any(ip$IPStart >= ip$IPEnd)) {
      cat("The following people in the IP table have an IP end date before the start date so will have no IP in analysis: ", file = log, append = T)
      cat(as.character(ip$ID[ip$IPStart > ip$IPEnd]), file = log, append = T)
      cat("\r\n", file = log, append = T)
      ip = ip[ip$IPStart < ip$IPEnd,]
    }
    if(nrow(ip)==0) {
      ip = NA
    }
  }
  
  if(all(class(progress)!="logical")) {
    progress$set(value = 1)
  }
  
  ###check still have data after removing bad rows
  if(length(cases) < 2) {
    stop("There are ", length(cases), " people to analyze. At least two people are needed.")
  }
  
  ###check that no date ranges for a particular case are overlapping (so don't double count days)
  dedup = NA
  for(c in cases) {
    cloc = sort(unique(loc$Location[loc$ID==c])) #location case was in
    if(length(cloc)) {
      for(l in cloc) {
        sub = loc[loc$ID==c & loc$Location==l,]
        rep = T #repeat the loop if overlaps found
        while(rep) {
          rep = F
          if(nrow(sub) >= 2) {
            i=1
            while(i < nrow(sub)) {
              j = i+1
              if(rep) {
                break;
              }
              while(j <= nrow(sub)) {
                ol = getOverlap(sub$Start[i], sub$End[i], sub$Start[j], sub$End[j])
                if(!all(is.na(ol))) {
                  cat(paste(c, " has rows with overlapping dates in ", l, ", which have been merged\r\n", sep=""), file = log, append = T)
                  rep = T
                  if(sub$Confidence[i]==sub$Confidence[j]) { #same confidence so can merge
                    sub$Start[i] = min(sub$Start[i], sub$Start[j])
                    sub$End[i] = max(sub$End[i], sub$End[j])
                    sub = sub[-j,]
                  } else { #remove overlap from uncertain
                    cert = ifelse(sub$Confidence[i] == "certain", i, j) #certain row
                    un = ifelse(sub$Confidence[i] == "uncertain", i, j) #uncertain row
                    if(ol[2] >= sub$Start[un] & ol[1] <= sub$Start[un]) { #uncertain starts in middle of certain
                      if(ol[2] < sub$End[un]) { #uncertain ends after certain
                        sub$Start[un] = ol[2] + 1 #change uncertain to start after overlap
                      } else { #uncertain contained within certain
                        sub = sub[-un,]
                      }
                    } else if(ol[1] <= sub$End[un] & ol[2] >= sub$End[un]) { #certain starts in middle of uncertain and ends after
                      sub$End[un] = ol[1] - 1
                    } else { #certain contained within uncertain
                      sub = rbind(sub, 
                                  data.frame(ID = c,
                                             Location = l,
                                             Start = ol[2]+1,
                                             End = sub$End[un],
                                             Confidence = "uncertain"))
                      sub$End[un] = ol[1] - 1
                    }
                    i = 1
                    break;
                  } 
                } else {
                  rep=F
                  j = j + 1
                }
              }
              i = i+1
            }
          }
        }
        if(all(is.na(dedup))) {
          dedup = sub
        } else {
          dedup = rbind(dedup, sub)
        }
      }
    }
  }
  loc = dedup
  
  if(all(class(progress)!="logical")) {
    progress$set(value = 2)
  }
  
  ###get list of overlaps
  res = NA
  for(i in 1:(length(cases)-1)) {
    cloc = sort(unique(loc$Location[loc$ID==cases[i]])) #location case was in
    if(length(cloc)) {
      for(l in cloc) {
        dates1 = loc[loc$ID==cases[i] & loc$Location==l,]
        for(j in (i+1):length(cases)) {
          dates2 = loc[loc$ID==cases[j] & loc$Location==l,]
          if(nrow(dates2)) {
            for(r1 in 1:nrow(dates1)) {
              for(r2 in 1:nrow(dates2)) {
                ol = getOverlap(dates1$Start[r1], dates1$End[r1], dates2$Start[r2], dates2$End[r2])
                if(!all(is.na(ol))) {
                  temp = data.frame(ID1=cases[i],
                                    ID2=cases[j],
                                    Location=l,
                                    Strength=ifelse(dates1$Confidence[r1]=="uncertain" | dates2$Confidence[r2]=="uncertain",
                                                    "uncertain", "certain"),
                                    NumDaysOverlap=as.numeric(ol[2]-ol[1])+1,
                                    OverlapStart=ol[1],
                                    OverlapEnd=ol[2],
                                    OverlapID1IP=as.character(overlapIP(ol[1], ol[2], ip, cases[i])),
                                    OverlapID2IP=as.character(overlapIP(ol[1], ol[2], ip, cases[j])))
                  if(all(is.na(res))) {
                    res = temp
                  } else {
                    res = rbind(res, temp, stringsAsFactors=F)
                  }
                }
              }
            }
          }
        }
      }
    }
  }
  
  if(all(class(progress)!="logical")) {
    progress$set(value = 3)
  }
  
  ###convert to epi links (one link per pair; if pair has multiple overlaps in time and space, strength = strongest, location=all locations (even ones with weaker epi))
  epi = NA
  if(!all(is.na((res)))) {
    if(ipEpiLink) {
      if(all(is.na(ip))) {
        cat("There are no IP epi links because there are no valid infectious periods\r\n", file = log, append = T)
        # cat("Generating IP epi links, but there are no valid infectious periods, so there are no IP epi links\r\n", file = log, append = T)
      } else {
        cat(paste("Generating IP epi links with cutoff of ", cutoff, " days overlapping each other and an IP", 
                  ifelse(removeAfter, ". Overlaps after either IP has ended were removed from link analysis.", 
                         ". Overlaps after either IP has ended were kept in link analysis."), "\r\n", sep=""), file = log, append = T)
        epi = getIPEpiLinks(res, cutoff, removeAfter)
      }
    } else {
      cat("Generating epi links with cutoff of", cutoff, "days for a definite or probable epi link\r\n", file = log, append = T)
      epi = getEpiLinks(res, cutoff)
    }
  } else {
    # cat("There are no links because no overlaps were found\r\n", file = log, append = T)
    cat("No overlaps were found. As a result, there are no links.\r\n", file = log, append = T)
  }
  
  if(all(class(progress)!="logical")) {
    progress$set(value = 4)
  }
  
  ###summarize the total overlap for each case
  if(all(is.na(res))) {
    tot = NA
  } else {
    ids = sort(unique(c(as.character(res$ID1), as.character(res$ID2))))
    tot = data.frame(ID = ids,
                     Locations = NA,
                     numCertOverlap = NA,
                     numTotOverlap = NA,
                     numCertIPOverap = NA,
                     numTotIPOverlap = NA,
                     stringsAsFactors = F)
    for(i in 1:nrow(tot)) {
      id = tot$ID[i]
      sub = res[res$ID1==id | res$ID2==id,]
      tot$Locations[i] = paste(sort(unique(as.character(sub$Location))), collapse = ", ")
      ##total overlap
      tot$numCertOverlap[i] = sum(sub$NumDaysOverlap[sub$Strength=="certain"], na.rm = T)
      tot$numTotOverlap[i] = sum(sub$NumDaysOverlap, na.rm = T)
      ##overlap with IP of another person
      sub$OverlapID1IP = as.character(sub$OverlapID1IP)
      sub$OverlapID2IP = as.character(sub$OverlapID2IP)
      sub$OverlapID1IP[sub$OverlapID1IP==afterIPstring | sub$OverlapID1IP==noIPstring] = 0
      sub$OverlapID2IP[sub$OverlapID2IP==afterIPstring | sub$OverlapID2IP==noIPstring] = 0
      sub$OverlapID1IP = as.numeric(sub$OverlapID1IP)
      sub$OverlapID2IP = as.numeric(sub$OverlapID2IP)
      tot$numCertIPOverap[i] = sum(c(sub$OverlapID1IP[sub$ID1!=id & sub$Strength=="certain"],
                                     sub$OverlapID2IP[sub$ID2!=id & sub$Strength=="certain"]), na.rm = T)
      tot$numTotIPOverlap[i] = sum(c(sub$OverlapID1IP[sub$ID1!=id],
                                     sub$OverlapID2IP[sub$ID2!=id]), na.rm = T)
    }
  }
  
  if(all(class(progress)!="logical")) {
    progress$set(value = 5)
  }
  return(list(allOverlaps = res, 
              epiLinks = epi,
              location = loc,
              ip = ip,
              summary = tot))
}

##for the given loc and ip tables, draw a timeline of dates of stay and IPs, one figure per location
###only include cases at the location
###only include times at the location (don't show IP if not on plot)
###if certLegOnly is true, only show the dark blue in the legend as "time" without certainty (for simplified presentations)
timelineFig <- function(outPrefix, loc, ip, certLegOnly=F) {
  for(l in unique(loc$Location)) {
    sub = loc[loc$Location==l,]
    lcases = sort(unique(sub$ID))
    y = length(lcases):1
    xrange = range(c(sub$Start, sub$End))
    jpeg(paste(outPrefix, "LATTE_Timeline_", l, ".jpg", sep=""), width=700, height=length(lcases)*15+500)
    layout(matrix(1:2, ncol=1, nrow=2), heights = c(1,.09))
    par(mar=c(4,8,4,.3)) 
    plot(NA, xaxt="n", yaxt="n", xlim=xrange, ylim=c(0, length(lcases)+.5), xlab = "", ylab = "", main=l)
    axis(side=2, at=y, labels=lcases, las=1)
    xlabs=seq(xrange[1], xrange[2], length.out = 10)
    axis(side=1, at=xlabs, labels=format(xlabs, "%m/%d/%Y"))
    ##separate cases
    abline(h=y-0.5, col="gray", lty=3)
    ##plot case data
    for(i in y) {
      c = lcases[length(lcases)-i+1] 
      ##plot times in location
      sub2 = sub[sub$ID==c,]
      for(r in 1:nrow(sub2)) {
        if(sub2$Confidence[r] == "certain") {
          lines(x=c(sub2$Start[r], sub2$End[r]), y=c(i, i), col="blue", lwd=4)
        } else if(sub2$Confidence[r] == "uncertain") {
          lines(x=c(sub2$Start[r], sub2$End[r]), y=c(i, i), col="lightskyblue", lwd=4)
        } else {
          cat(paste("Bad strength for plotting in row ", r, ": ", sub2$Confidence[r], "\r\n", sep=""), file = log, append = T)
          lines(x=c(sub2$Start[r], sub2$End[r]), y=c(i, i), col="gray", lwd=4)
        }
      }
      ##plot IP
      if(!all(is.na(ip))) {
        if(c %in% ip$ID) {
          lines(x=c(ip$IPStart[ip$ID==c], ip$IPEnd[ip$ID==c]), y=c(i+.25, i+.25), col="red", lwd=4)
        }
      }
    }
    ##add legend (only include the strengths in the table)
    par(mar=c(.1,4,.1,.3))
    plot(NA, xaxt="n", yaxt="n", bty="n", xlim=c(0,1), ylim=c(0,1), xlab="", ylab="")
    if(certLegOnly) {
      legend("center", 
             horiz=T,
             legend=c("Time in location", "Infectious period"),
             col=c("blue", "red"),
             lwd=2,
             bty="n",
             cex=1.3)
    } else {
      legend("center", 
             horiz=T,
             legend=c("Certain time in location", "Uncertain time in location", "Infectious period"),
             col=c("blue", "lightskyblue", "red"),
             lwd=2,
             bty="n")
    }
    dev.off()
  }
}

##run the LATTE algorithm and then produce Excel files of results and timeline figures
##outPrefix = prefix for output files
##loc = table of locations, expected columns = ID, location, start, end, confidence
##ip = optional table of infectious periods, expected columns = ID, IPStart, IPEnd
##cutoff = overlap must be more than cutoff number of days to be considered for an epi or IP epi link
##ipEpiLink = if true, calculate an IP epi link, otherwise calculate an epi link
##removeAfter = if true, remove overlaps that occur after the IP of either case (for IP epi links)
##progress = optional progress bar (for Rshiny interface)
##drawTimeline = if true, generate timeline jpeg
latteWithOutputs <- function(outPrefix, loc, ip = NA, cutoff = defaultCut, ipEpiLink = F, removeAfter = F, progress = NA,
                             drawTimeline = F) {
  log = paste(outPrefix, defaultLogName, sep="")
  results = latte(loc = loc, ip = ip, cutoff = cutoff, ipEpiLink = ipEpiLink, removeAfter = removeAfter, log = log)
  res = results$allOverlaps
  epi = results$epiLinks
  loc = results$location
  ip = results$ip
  tot = results$summary
  
  ##if Excel spreadsheet already exists, delete file; otherwise will not write the new results
  overlapName = paste(outPrefix, "LATTE_All_Overlaps.xlsx", sep="")
  epiName = paste(outPrefix, ifelse(ipEpiLink, "LATTE_IPEpi_Links_", "LATTE_Epi_Links_"), cutoff, "DCutoff",
                  ifelse(ipEpiLink, ifelse(removeAfter, "", "_KeepOLAfterIPEnd"), ""), ".xlsx", sep="")
  locName = paste0(outPrefix, "LATTE_Location.xlsx")
  ipName = paste0(outPrefix, "LATTE_IP.xlsx")
  summaryName = paste0(outPrefix, "LATTE_Summary_By_Person.xlsx")
  outputExcelFiles = c(overlapName, epiName, locName, ipName)
  if(any(file.exists(outputExcelFiles))) {
    del = outputExcelFiles[file.exists(outputExcelFiles)]
    file.remove(del)
  }
  if(all(class(progress)!="logical")) {
    progress$set(value = 6)
  }
  
  ###write out overlaps
  if(!all(is.na(res))) {
    # write.table(res, sub(".xlsx", ".txt", overlapName), row.names = F, col.names = T, quote = F, sep = "\t")
    res$OverlapStart = format(res$OverlapStart, format="%m/%d/%Y")
    res$OverlapEnd = format(res$OverlapEnd, format="%m/%d/%Y")
    names(res)[names(res)=="Strength"] = "Confidence"
    writeExcelTable(df=res, fileName=overlapName, wrapHeader = T, sheetName = "All Overlaps")
  } else {
    outputExcelFiles = outputExcelFiles[outputExcelFiles != overlapName]
  }
  if(all(class(progress)!="logical")) {
    progress$set(value = 7)
  }
  
  ###write out epi links
  if(!all(is.na(epi))) {
    # write.table(epi, sub(".xlsx", ".txt", overlapName), row.names = F, col.names = T, quote = F, sep = "\t")
    writeExcelTable(df=epi, fileName = epiName, 
                    sheetName = paste(ifelse(ipEpiLink, "IPEpi ", "Epi "), cutoff, "D", 
                                      ifelse(ipEpiLink, ifelse(removeAfter, "", " KeepOLAfterIPend"), ""), sep=""),
                    wrapHeader = T) 
  } else {
    outputExcelFiles = outputExcelFiles[outputExcelFiles != epiName]
    if(ipEpiLink & !all(is.na(ip)) & !all(is.na(res))) {
      cat(paste("No IP epi links detected with ", cutoff, " day cutoff", 
                ifelse(removeAfter, " and removing overlaps after IP end", ""), "\r\n", sep=""), file = log, append = T)
    } else if(!ipEpiLink) {
      if(!all(is.na(res))) {
        cat(paste("No epi links detected with", cutoff, "day cutoff\r\n"), file = log, append = T)
      }
    }
  }
  if(all(class(progress)!="logical")) {
    progress$set(value = 9)
  }
  
  ###write out location table
  if(!all(is.na(loc))) {
    writeExcelTable(df=loc, fileName=locName, wrapHeader = T, sheetName = "Location Data")
  } else {
    outputExcelFiles = outputExcelFiles[outputExcelFiles != locName]
  }
  
  if(all(class(progress)!="logical")) {
    progress$set(value = 10)
  }
  
  ###write out IP table
  if(!all(is.na(ip))) {
    writeExcelTable(df=ip, fileName=ipName, wrapHeader = T, sheetName = "IP Data")
  } else {
    outputExcelFiles = outputExcelFiles[outputExcelFiles != ipName]
  }
  if(all(class(progress)!="logical")) {
    progress$set(value = 11)
  }
  
  ###write out person summary table
  if(!all(is.na(tot))) {
    writeExcelTable(df=tot, fileName=summaryName, wrapHeader = T, sheetName = "Summary By Person")
  } else {
    outputExcelFiles = outputExcelFiles[outputExcelFiles != summaryName]
  }
  if(all(class(progress)!="logical")) {
    progress$set(value = 12)
  }
  
  ###generate figure
  if(drawTimeline) {
    figs = timelineFig(outPrefix, loc, ip)
    outputExcelFiles = c(outputExcelFiles, figs) #add figs to output list
  }
  
  results$outputFiles = c(log, outputExcelFiles)
  return(results)
}
