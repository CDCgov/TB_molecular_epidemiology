##logically inferred tuberculosis transmission (LITT)
##for each case, produces a ranked list of potential sources

source("../sharedFunctions.R")
library(lubridate) #for adding months
library(xlsx) #for Excel writing functions (not used directly by LITT)

###function that returns the time score: 2 if IP start after target date, 1 if IP end 2 years earlier than target date, 0 otherwise
## sourceID = source state case number
## dates = table of dates
## targetDate = target date used as reference (target IP start)
timeScore <- function(sourceID, dates, targetDate) {
  if(dates$IPStart[dates$ID==sourceID] >= targetDate) { #IP start after target date
    return(2)
  } else if(is.na(dates$IPEnd[dates$ID==sourceID])) {
    return(ifelse(time_length(difftime(targetDate, dates$IPStart[dates$ID==sourceID]), "years") < 2,
                  0, 1))
  } else { #use IP start if no IP end available
    return(ifelse(time_length(difftime(targetDate, dates$IPEnd[dates$ID==sourceID]), "years") < 2,
                  0, 1))
  }
}

###function that returns the epi link score for a pair (case1, case2) in a dataframe of epi links (epi)
noEpiScore = 3
epiLinkScore <- function(case1, case2, epi, log) {
  stren = tolower(epi$strength[(epi$case1==case1 & epi$case2==case2) |
                                 (epi$case1==case2 & epi$case2==case1)])
  if(length(stren) == 0) { #no link
    return(noEpiScore)
  } else if(stren == "definite") {
    return(0)
  } else if(stren == "probable") {
    return(1)
  } else if(stren == "possible") {
    return(2)
  } else {
    cat(paste("Bad epi link strength:", case1, case2, stren, "\nThis will be treated as no epi link.\n"), file = log, append = T)
    return(noEpiScore)
  }
}

##function that returns the epi link label for a pair (case1, case2) in a dataframe of epi links (epi)
epiLinkLabel <- function(case1, case2, epi) {
  lab = epi$label[(epi$case1==case1 & epi$case2==case2) |
                    (epi$case1==case2 & epi$case2==case1)]
  if(length(lab) == 0) { 
    return(" ")
  } else if(is.na(lab)) {
    return(" ")
  } else if(lab == "NA") {
    return(" ")
  } else {
    return(lab)
  }
}

##function that returns the label with the shared additional risk factor rfname appended to the given label
addlRFLabel <- function(label, rfname) {
  rfname = paste("RF:", rfname)
  if(is.na(label) | label==" " | label == "") {
    return(rfname)
  } else {
    return(paste(label, rfname, sep=" | "))
  }
}

##function that returns the user weights
getAddlRFUserUserWeights <- function(addlRiskFactor) {
  userWeight = addlRiskFactor[grepl("weight", addlRiskFactor$ID, ignore.case = T),-1]
  userWeight[] = lapply(userWeight, as.character)
  # if(class(userWeight)=="factor") {
  if(class(userWeight)!="data.frame") {
    userWeight = as.numeric(as.character(userWeight))
  } else {
    userWeight = as.numeric(userWeight[1,])
  }
  return(userWeight)
}

##for the given epi data, make sure the columns are correctly named
fixEpiNames <- function(epi, log) {
  epicolnames = c("case1", "case2", "strength", "label")
  if(any(!is.na(epi))) {
    ##case1
    c1 = grepl("case1", names(epi), ignore.case = T) | grepl("source", names(epi), ignore.case = T) | 
      grepl("stcaseno1", names(epi), ignore.case = T) | grepl("id1", names(epi), ignore.case = T)
    if(sum(c1) == 1) {
      names(epi)[c1] = "case1"
    } else if(sum(c1) > 1) {
      cat("Epi table should have the following columns: ", paste(epicolnames, collapse=", "), file = log, append = T)
      cat(paste("\nHowever, more than one column for first case was found in epi link table:", paste(names(epi)[c1], collapse = ", ")),
           "\nPlease label one column case1.\n", file = log, append = T)
      stop(paste("More than one column for first case in epi link table:", paste(names(epi)[c1], collapse = ", ")),
           "\nPlease label one column case1.")
    } else {
      cat("Epi table should have the following columns: ", paste(epicolnames, collapse=", "), file = log, append = T)
      cat("\nColumn for first case in epi link table not found. Please label one column case1.\n", file = log, append = T)
      stop("Column for first case in epi link table not found. Please label one column case1.")
    }
    ##case2
    c1 = grepl("case2", names(epi), ignore.case = T) | grepl("target", names(epi), ignore.case = T)| 
      grepl("stcaseno2", names(epi), ignore.case = T) | grepl("id2", names(epi), ignore.case = T)
    if(sum(c1) == 1) {
      names(epi)[c1] = "case2"
    } else if(sum(c1) > 1) {
      cat("Epi table should have the following columns: ", paste(epicolnames, collapse=", "), file = log, append = T)
      cat(paste("\nHowever, more than one column for second case in was found in epi link table:", paste(names(epi)[c1], collapse = ", ")),
          "\nPlease label one column case2.\n", file = log, append = T)
      stop(paste("More than one column for second case in epi link table:", paste(names(epi)[c1], collapse = ", ")),
           "\nPlease label one column case2")
    } else {
      cat("Epi table should have the following columns: ", paste(epicolnames, collapse=", "), file = log, append = T)
      cat("\nColumn for second case in epi link table not found. Please label one column case2.\n", file = log, append = T)
      stop("Column for second case in epi link table not found\n Please label one column case2.")
    }
    ##strength
    c1 = grepl("strength", names(epi), ignore.case = T)
    if(sum(c1) == 1) {
      names(epi)[c1] = "strength"
    } else if(sum(c1) > 1) {
      cat("Epi table should have the following columns: ", paste(epicolnames, collapse=", "), file = log, append = T)
      cat(paste("\nHowever, more than one column forepi link strength was found in epi link table:", 
                paste(names(epi)[c1], collapse = ", ")),
          "\nPlease label one column strength\n", file = log, append = T)
      stop(paste("More than one column for epi link strength in epi link table:", paste(names(epi)[c1], collapse = ", ")),
           "\nPlease label one column strength")
    } else {
      cat("Epi table should have the following columns: ", paste(epicolnames, collapse=", "), file = log, append = T)
      cat("\nColumn for epi link strength in epi link table not found. Please label one column strength.\n", file = log, append = T)
      stop("Column for epi link strength in epi link table not found\n Please label one column strength.")
    }
    epi$strength = tolower(epi$strength)
    epi$strength = trimws(epi$strength) #remove extra spaces, which can mess up comparisons
    ##label
    c1 = grepl("label", names(epi), ignore.case = T) 
    if(sum(c1) == 1) {
      names(epi)[c1] = "label"
    } else {
      c1 = grepl("location", names(epi), ignore.case = T)
      if(sum(c1) == 1) {
        names(epi)[c1] = "label"
      } else if(sum(c1) > 1) {
        epi$label=""
        cat(paste("More than one column for epi link label in epi link table:", paste(names(epi)[c1], collapse = ", "), 
                  "\n"), file = log, append = T)
        cat("To have labels, please name one column \"label\".\n", file = log, append = T)
      } else {
        epi$label = ""
        cat("Column for epi link label in epi link table not found. To have labels, please name one column \"label\".\n", 
            file = log, append = T)
        cat("No labels will be used in this run.\n", file = log, append = T)
      }
    }
    ##if SNP distance already present, remove with warning
    if(any(!names(epi) %in% c("case1", "case2", "strength", "label"))) {
      cat(paste("There are extra columns in the epi table, which will be removed:",
                paste(names(epi)[!names(epi) %in% c("case1", "case2", "strength", "label")], collapse = ", "), 
                "\n"), file = log, append = T)
    }
    epi = epi[,names(epi) %in% epicolnames] #remove extra columns (needed for rbind with TB GIMS data)
  }
  epi$case1 = as.character(epi$case1)
  epi$case2 = as.character(epi$case2)
  epi$strength = as.character(epi$strength)
  if("label" %in% names(epi)) {
    epi$label = as.character(epi$label)
  }
  return(epi)
}

##check infectious period column names
##check for duplicates; if same isolate has two different IP dates, give warning and take the earlier
fixIPnames <- function(df, log) {
  if(any(!is.na(df))) {
    ##start
    col = grepl("ip[ .]*start", names(df), ignore.case = T) | grepl("infectious[ .]*period[ .]*start", names(df), ignore.case = T)
    if(sum(col) == 1) {
      names(df)[col] = "IPStart"
    } else if(sum(col) > 1) {
      warning(paste("More than one column for infectious period (IP) start:", paste(names(df)[col], collapse = ", ")),
              "\nPlease label one column IPStart.\n User-input IP start was not used in this analysis.")
      cat(paste("More than one column for infectious period (IP) start:", paste(names(df)[col], collapse = ", ")),
          "\nPlease label one column IPStart.\n User-input IP start was not used in this analysis.\n", 
          file = log, append = T)
      return(NA)
    } else if(sum(col) < 1) {
      cat("No infectious period (IP) start column, so user-input IP was not used in the analysis.\nPlease add a column named IPStart to have user IP in analysis.\n", 
          file = log, append = T)
      return(NA)
    }
    
    ##end is optional
    col = grepl("ip[ .]*end", names(df), ignore.case = T) | grepl("infectious[ .]*period[ .]*end", names(df), ignore.case = T) |
      grepl("ip[ .]*stop", names(df), ignore.case = T) | grepl("infectious[ .]*period[ .]*stop", names(df), ignore.case = T)
    if(sum(col) == 1) {
      names(df)[col] = "IPEnd"
    }else if(sum(col) > 1) {
      warning(paste("More than one column for infectious period (IP) end:", paste(names(df)[col], collapse = ", ")),
              "\nPlease label one column IPEnd.\n User-input IP end was not used in this analysis.")
      cat(paste("More than one column for infectious period (IP) end:", paste(names(df)[col], collapse = ", ")),
          "\nPlease label one column IPEnd.\n User-input IP end was not used in this analysis.\n", 
          file = log, append = T)
      df$IPEnd = NA
    } else if(sum(col) < 1) {
      cat("No infectious period (IP) end column, so user-input IP end was not used in the analysis.\nPlease add a column named IPEnd to use IP end in analysis.\n", 
          file = log, append = T)
      df$IPEnd = NA
    }
    
    ##check for duplicates and clean up if present
    df$ID = as.character(df$ID)
    dup = duplicated(df$ID)
    if(any(dup)) {
      ids = df$ID[dup]
      for(i in ids) {
        rows = which(df$ID==i)
        df$IPStart[rows[1]] = getMin(df$IPStart[df$ID==i], i, "IP start")
        df$IPEnd[rows[1]] = getMin(df$IPEnd[df$ID==i], i, "IP end", min=F, message=F)
        df = df[-rows[-1],]
      }
    }
  }
  return(df)
}

##remove first X from names if they are supposed to start with a number (e.g. column names of distance matrix)
removeXFromNames <- function(names) {
  fix = grepl("^X[0-9]", names)
  names[fix] = sub("^X", "", names[fix])
  return(names)
}

##function that returns the score calculated from the given rating matrix score
getScore <- function(score) {
  return(score$snpDistance + score$infRate + score$epiRate + score$timeRate)
}

##function that returns the score calculated from the given rating matrix score, ignoring WGS
getScoreWithoutWGS <- function(score) {
  return(score$infRate + score$epiRate + score$timeRate)
}

##returns the data frame converted from numeric score to score category
##cat = data.frame to convert to category
getScoreCategories <- function(cat) {
  cat$snpDistance[is.na(cat$snpDistance) & cat$source!="no potential sources"] = "no SNP data"
  cat$timeRate = as.character(cat$timeRate)
  cat$timeRate = ifelse(is.na(cat$time) | cat$time =="", "",
                        ifelse(cat$time=="2", 
                               "source IP start later than given case IP start",
                               ifelse(cat$time=="1",
                                      "source IP end 2+ years earlier than given case IP start",
                                      "source earlier than given case and within two years")))
  cat$infRate = as.character(cat$infRate)
  cat$infRate = ifelse(cat$infRate=="", "",
                       ifelse(cat$infRate=="0", "cavitary",
                              ifelse(cat$infRate=="1", "smear positive non-cavitary", "smear negative non-cavitary")))
  cat$epiRate = as.character(cat$epiRate)
  cat$epiRate = ifelse(cat$epiRate=="", "",
                       ifelse(cat$epiRate=="0", "definite epi link",
                              ifelse(cat$epiRate=="1", "probable epi link",
                                     ifelse(cat$epiRate=="3", "no epi link or shared risk factors",
                                            ifelse(cat$epiRate=="2", "possible epi link or all risk factors shared",
                                                   "shared risk factors")))))
  ##make columns numeric for Excel
  cat[,8:9] = apply(cat[,8:9], 2, as.numeric) #don't make rank numeric because sometimes will have asterisk
  return(cat)
}

##function that makes the header of the given data frame df nice for output
##use this function to be consistent between tables and avoid issues of changing variable order
##df = data frame to clean names of 
##snpRate = if true, snpDistance is a rating, otherwise it is just SNP distance
##stcasenolab = if true, ID column is named state case number, otherwise is ID
##returns data frame with cleaned names
cleanHeaderForOutput <- function(df, snpRate = F, stcasenolab = F) {
  ##variables in case line list
  names(df)[names(df)=="STCASENO"] = "State Case Number"
  names(df)[names(df)=="ID"] = ifelse(stcasenolab, "State Case Number", "Case ID")
  names(df)[names(df)=="accessionNumber"] = "Accession Number"
  names(df)[names(df)=="zipcode"] = "ZIP Code"
  names(df)[names(df)=="county"] = "County"
  names(df)[names(df)=="state"] = "State"
  names(df)[names(df)=="gender"] = "Gender"
  names(df)[names(df)=="RACEHISP"] = "Race/Ethnicity of Patient" #attr(gimsAll$RACEHISP, "label")
  names(df)[names(df)=="age"] = "Age"
  names(df)[names(df)=="earliestDate"] = "Calculated Earliest Date"
  names(df)[names(df)=="IPStart"] = "Calculated Infectious Period Start"
  names(df)[names(df)=="IPEnd"] = "Calculated Infectious Period End"
  # names(df)[names(df)=="infRate"] = "Infectious Category"
  names(df)[names(df)=="SPSMEAR"] = "Sputum Smear"
  names(df)[names(df)=="XRAYCAV"] = "Evidence of Cavity by X-Ray"
  names(df)[names(df)=="UserDateData"] = "User Input Date Data Available"
  names(df)[names(df)=="sequenceAvailable"] = "Sequence Available In Analysis"
  names(df)[names(df)=="pediatric"] = "Pediatric"
  names(df)[names(df)=="ExtrapulmOnly" | names(df)=="ExtrapulmonaryOnly"] = "Extrapulmonary Only"
  names(df)[names(df)=="numEpiLinks"] = "Number of Epi Links"
  names(df)[names(df)=="numTimesIsRank1"] = "Number of Times is Ranked 1st in Potential Source List"
  names(df)[names(df)=="numPotSourcesRanked1"] = "Number of Potential Sources Tied for Rank 1"
  names(df)[names(df)=="totNumPotSources"] = "Total Number of Potential Sources"
  
  ##variables in possible GIMS risk factors
  names(df)[names(df)=="HOMELESS"] = "GIMS Homeless"# Within Past Year"
  names(df)[names(df)=="HIVSTAT"] = "GIMS HIV Status"
  names(df)[names(df)=="CORRINST"] = "GIMS Correctional Facility Resident"
  names(df)[names(df)=="LONGTERM"] = "GIMS Long-Term Care Facility Resident"
  names(df)[names(df)=="IDU"] = "GIMS Injecting Drug Use"
  names(df)[names(df)=="NONIDU"] = "GIMS Non-Injecting Drug Use"
  names(df)[names(df)=="ALCOHOL"] = "GIMS Alcohol"
  names(df)[names(df)=="OCCUHCW"] = "GIMS Health Care Worker"
  names(df)[names(df)=="OCCUCORR"] = "GIMS Correctional Facility Employee"
  names(df)[names(df)=="RISKTNF"] = "GIMS TNF alpha Antagonist Therapy"
  names(df)[names(df)=="RISKORGAN"] = "GIMS Post-Organ Transplant"
  names(df)[names(df)=="RISKDIAB"] = "GIMS Diabetes Mellitus"
  names(df)[names(df)=="RISKRENAL"] = "GIMS End-Stage Renal Disease"
  names(df)[names(df)=="RISKIMMUNO"] = "GIMS Immunosuppression Not HIV/AIDS"
  
  ##additional variables in date file
  names(df)[names(df)=="inputSxOnset"] = "Input Symptom Onset Date"
  names(df)[names(df)=="inputIPStart"] = "Input Infectious Period Start"
  names(df)[names(df)=="inputIPEnd"] = "Input Infectious Period End"
  names(df)[names(df)=="ISUSDATE"] = "Date of Initial Susceptibility Testing"
  names(df)[names(df)=="RXDATE"] = "Start Therapy Date"
  names(df)[names(df)=="CNTDATE"] = "Count Date"
  names(df)[names(df)=="RPTDATE"] = "Report Date"
  names(df)[names(df)=="sp_coll_date"] = "Specimen Collection Date"
  names(df)[names(df)=="sxOnset"] = "Calculated Symptom Onset Date"
  
  ##additional variables in epi link file
  names(df)[names(df)=="case1"] = "Case1"
  names(df)[names(df)=="case2"] = "Case2"
  names(df)[names(df)=="strength"] = "Epi Link Strength"
  names(df)[names(df)=="label"] = "Label"
  names(df)[names(df)=="SNPdistance" | names(df)=="snpDistance"] = ifelse(snpRate, "SNP Rating", "SNP Distance")
  
  ##additional variables in transmission file
  names(df)[names(df)=="target"] = "Given Case"
  names(df)[names(df)=="source"] = "Potential Source"
  names(df)[names(df)=="score"] = "Score"
  names(df)[names(df)=="scoreWithoutWGS" | names(df)=="scoreWoWGS"] = "Without SNP Score"
  
  ##additional variables in all potential sources
  names(df)[names(df)=="infRate"] = "Infectious Rating" 
  names(df)[names(df)=="timeRate"] = "Time Rating"
  names(df)[names(df)=="epiRate"] = "Epi and Risk Factor Rating"
  names(df)[names(df)=="rank"] = "Rank"
  
  ##additional variables in filtered list
  names(df)[names(df)=="filteredCase"] = "Filtered Case"
  names(df)[names(df)=="reasonFiltered"] = "Reason Filtered"
  
  return(df)
}

# ##writes the data frame df to the Excel file or workbook and sheet (using XLConnect)
# ##df = data to write
# ##if fileName is not NA, generate the workbook first
# ##workbook should be provided if no fileName is given, otherwise the value will be ignored
# ##returns the workbook in case additional sheets need to be written
# ##if wrapHeader is true, set column widths to the longest length and wrap the header (rather than using auto, which will not wrap the header)
# ##stcasenolab = if true, ID column is named state case number, otherwise is ID
# ##snpRate = if true, SNP column is SNP rating, otherwise is SNP distance
# writeExcelTable<-function(fileName=NA, workbook, sheetName="Sheet1", df, wrapHeader=F, stcasenolab = F, snpRate = F) {
#   df = cleanHeaderForOutput(df, stcasenolab = stcasenolab, snpRate = snpRate)
#   if(!is.na(fileName)) {
#     workbook = loadWorkbook(fileName, create=T)
#   }
#   createSheet(workbook, name=sheetName)
#   # # writeWorksheet(workbook, df, sheet = sheetName) ##this results in occasional duplicated rows, and flags that numbers were treated as strings
#   # if(wrapHeader) {
#   #   ##set column widths
#   #   cwidth = (sapply(1:ncol(df), function(c){max(nchar(as.character(df[,c]), keepNA=F))})+5)*256 #width is 1/256 of char
#   #   minWidth = 10*256 #minimum width for column
#   #   cwidth[cwidth < minWidth] = minWidth
#   #   setColumnWidth(workbook, sheetName, column=1:ncol(df), width=cwidth)
#   #   
#   #   # wrap header and not other columns
#   #   wrap = createCellStyle(workbook)
#   #   setWrapText(wrap, wrap=T)
#   #   setFillPattern(wrap, fill = XLC$"FILL.SOLID_FOREGROUND")
#   #   setFillForegroundColor(wrap, color=XLC$"COLOR.GREY_25_PERCENT")
#   #   setCellStyle(workbook, sheet=sheetName, row=1, col=1:ncol(df), cellstyle=wrap)
#   #   
#   # } else {
#   #   ##autofit column width
#   #   setColumnWidth(workbook, sheetName, column=1:ncol(df), width=-1)
#   # }
#   
#   ##set up header
#   writeWorksheet(workbook, t(names(df)), sheet = sheetName, header = F, rownames = F)
#   ##write data
#   for(r in 1:nrow(df)) {
#     for(c in 1:ncol(df)) {
#       if(!is.na(df[r,c])) {
#         if(as.character(df[r,c])!="") {
#           # if(grepl("[A-Z]|[a-z]", df[r,c]) | grepl("*", df[r,c], fixed=T) | grepl("-", df[r,c]) | grepl("/", df[r,c]) | 
#           #    grepl("+", df[r,c], fixed = T) | is.Date(df[r,c])) {
#           #   writeWorksheet(workbook, df[r,c], sheet = sheetName, startRow = r+1, startCol = c, header = F, rownames = F)
#           # } else {
#           #   writeWorksheet(workbook, as.numeric(df[r,c]), sheet = sheetName, startRow = r+1, startCol = c, header = F, rownames = F)
#           # }
#           if(all(grepl("[0-9.]", strsplit(as.character(df[r,c]), "")[[1]]))) {
#             writeWorksheet(workbook, as.numeric(df[r,c]), sheet = sheetName, startRow = r+1, startCol = c, header = F, rownames = F)
#           } else {
#             writeWorksheet(workbook, df[r,c], sheet = sheetName, startRow = r+1, startCol = c, header = F, rownames = F)
#           }
#         }
#       }
#     }
#   }
#   ##format worksheet
#   if(wrapHeader) {
#     ##set column widths
#     cwidth = (sapply(1:ncol(df), function(c){max(nchar(as.character(df[,c]), keepNA=F))})+5)*256 #width is 1/256 of char
#     minWidth = 10*256 #minimum width for column
#     cwidth[cwidth < minWidth] = minWidth
#     setColumnWidth(workbook, sheetName, column=1:ncol(df), width=cwidth)
#     
#     # wrap header and not other columns
#     wrap = createCellStyle(workbook)
#     setWrapText(wrap, wrap=T)
#     setFillPattern(wrap, fill = XLC$"FILL.SOLID_FOREGROUND")
#     setFillForegroundColor(wrap, color=XLC$"COLOR.GREY_25_PERCENT")
#     setCellStyle(workbook, sheet=sheetName, row=1, col=1:ncol(df), cellstyle=wrap)
#     
#   } else {
#     ##autofit column width
#     setColumnWidth(workbook, sheetName, column=1:ncol(df), width=-1)
#     cs  = createCellStyle(workbook)
#     setFillPattern(cs, fill = XLC$"FILL.SOLID_FOREGROUND")
#     setFillForegroundColor(cs, color=XLC$"COLOR.GREY_25_PERCENT")
#     setCellStyle(workbook, sheet=sheetName, row=1, col=1:ncol(df), cellstyle=cs)
#   }
#   ##add filter
#   setAutoFilter(workbook, sheet = sheetName, reference = aref("A1", dim(df)))
#   saveWorkbook(workbook)
#   return(workbook)
# }

##writes the data frame df to the Excel file or workbook and sheet
##df = data to write
##if fileName is not NA, generate the workbook first
##workbook should be provided if no fileName is given, otherwise the value will be ignored
##returns the workbook in case additional sheets need to be written
##if wrapHeader is true, set column widths to the longest length and wrap the header (rather than using auto, which will not wrap the header)
##stcasenolab = if true, ID column is named state case number, otherwise is ID
##snpRate = if true, SNP column is SNP rating, otherwise is SNP distance
##save = if true, save the workbook
writeExcelTable<-function(fileName, workbook=NA, sheetName="Sheet1", df, wrapHeader=F, stcasenolab = F, snpRate = F, save = T) {
  # detach(package:XLConnect, unload=T)
  # detach(package:XLConnectJars, unload=T)
  # library(xlsx)
  df = cleanHeaderForOutput(df, stcasenolab = stcasenolab, snpRate = snpRate)
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
            setCellValue(cells[[r+1,c]], as.numeric(df[r,c]))
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
    minWidth = 10# 10*256 #minimum width for column
    cwidth[cwidth < minWidth] = minWidth
    # setColumnWidth(sheet = sheet, colIndex = 1:ncol(df), colWidth = cwidth)
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
  # addAutoFilter(sheet = sheet, cellRange = "1")
  if(save) {
    saveWorkbook(workbook, fileName)
  }
  # detach(package:xlsx, unload=T)
  # library(XLConnect)
  return(workbook)
}

##for the given set of litt results (littResults, which is returned from LITT)
##write out the all potential sources table, combining the rating and label in one table
##stcasenolab indicates whether to use stcaseno or ID in the header for the table
##outPrefix is the prefix for the output files
writeAllSourcesTable <- function(littResults, outPrefix, stcasenolab = F) {
  allSources = littResults$allPotentialSources
  cat = getScoreCategories(allSources)
  allSources[,c(3:6, 8:9)] = apply(allSources[,c(3:6, 8:9)], 2, as.numeric) #make columns numeric for Excel
  ##move label to end
  allSources = cbind(allSources[,names(allSources)!="label"], data.frame(label=allSources[,names(allSources)=="label"]))
  cat = cbind(cat[,names(cat)!="label"], data.frame(label=cat[,names(cat)=="label"]))
  ##merge categorical and numerical rating
  cat$timeRate = ifelse(is.na(allSources$timeRate), NA, paste(allSources$timeRate, " (", cat$timeRate, ")", sep=""))
  cat$infRate = ifelse(is.na(allSources$infRate), NA, paste(allSources$infRate, " (", cat$infRate, ")", sep=""))
  cat$epiRate = ifelse(is.na(allSources$epiRate), NA, paste(allSources$epiRate, " (", cat$epiRate, ")", sep=""))
  ##move rank and scores first
  cat = cat[,c(1:2, 9, 7:8, 3:6, 10)]
  
  ##write
  fileName=paste(outPrefix, psFileName, sep="")
  wb = writeExcelTable(fileName=fileName,
                       sheetName = "potential sources",
                       df = cat,
                       stcasenolab = stcasenolab,
                       snpRate = T,
                       save = F)
  wb = writeExcelTable(fileName=fileName,
                       workbook=wb,
                       sheetName = "filtered cases",
                       df = littResults$filteredSources,
                       stcasenolab = stcasenolab)
}

##for the given set of litt results (littResults, which is returned from LITT)
##write out the top ranked sources transmission table
##stcasenolab indicates whether to use stcaseno or ID in the header for the table
##outPrefix is the prefix for the output files
##returns whether wrote out file
writeTopRankedTransmissionTable <- function(littResults, outPrefix, stcasenolab = F, log) {
  if(nrow(littResults$topRanked)) {
    if(!all(is.na(littResults$topRanked$score))) {
      littResults$topRanked$scoreWoWGS[!is.na(littResults$topRanked$score)] = NA #only give w/o SNP score if score is missing for some cases
    }
    littResults$topRanked$label = as.character(littResults$topRanked$label)
    writeExcelTable(fileName=paste(outPrefix, txFileName, sep=""),
                    sheetName = "top ranked sources", 
                    df = littResults$topRanked,
                    stcasenolab = stcasenolab)
    return(T)
  } else {
    cat("There were no cases with a potential source that passed all filters.\n", file = log, append = T)
    return(F)
  }
}

##for the given set of litt results (littResults, which is returned from LITT)
##write out the epi link table
##stcasenolab indicates whether to use stcaseno or ID in the header for the table
##outPrefix is the prefix for the output files
##returns whether wrote out file
writeEpiTable <- function(littResults, outPrefix, stcasenolab = F, log) {
  if(nrow(littResults$epi)) {
    epiOut = littResults$epi
    writeExcelTable(fileName=paste(outPrefix, epiFileName, sep=""),
                    sheetName="Epi links",
                    df = epiOut,
                    stcasenolab = stcasenolab)
    return(T)
  } else {
    cat("No epi links found; no epi link table will be generated.\n", file = log, append = T)
    return(F)
  }
}


##function that takes the dataframe df from RShiny fileInput and returns NA if the file does not exist
##otherwise, returns the distance matrix result needed for LITT
##if BN is true, indicates a BioNumerics distance triangle with accession numbers
readShinyDistanceMatrix <- function(df, bn=F) {
  if(is.null(df)) {
    return(NA)
  } 
  fname = df$datapath
  # print(fname)
  if(!file.exists(fname)) {
    return(NA)
  }
  if(!bn) {
    return(formatDistanceMatrixWithoutDedupOrStno(fname))
  } else {
    return(formatDistanceMatrix(fname))
  }
}

###for the given set of epi links in epi, clean up the strengths, labels, de-duplicate, and subset to only those cases in the analysis
##epi = epi link data frame
##cases = list of cases
cleanEpi <- function(epi, cases, log) {
  if(all(is.na(epi))) {
    epi = data.frame(case1 = character(),
                     case2 = character(),
                     strength = character(),
                     label = character(),
                     stringsAsFactors = F)
  } else { #clean
    epi = epi[epi$case1 %in% cases & epi$case2 %in% cases,] #so don't print errors for unnecessary rows
    ##check strengths
    epi$strength = as.character(tolower(epi$strength))
    if(any(is.na(epi$strength) | epi$strength == "")) {
      r = which(is.na(epi$strength) | epi$strength == "")
      cat(paste("These rows are missing an epi link strength and have been set to probable:",
                paste(r, collapse = ", "), "\n"), file = log, append = T)
      for(z in r) {
        cat("\n", file = log, append = T)
        cat(paste(epi[z,], collapse="\t"), file = log, append = T)
      }
      cat("\n", file = log, append = T)
      epi$strength[r] = "probable"
    }
    if(any(epi$strength != "definite" & epi$strength != "probable" & epi$strength != "possible", na.rm = T)) {
      r = which(epi$strength != "definite" & epi$strength != "probable" & epi$strength != "possible")
      cat(paste("These rows have an invalid epi link strength and have been set to probable:",
                paste(r, collapse = ", "), "\n"), file = log, append = T)
      # print(epi[r,])
      cat(paste(names(epi), collapse ="\t"), file = log, append = T)
      for(z in r) {
        cat("\n", file = log, append = T)
        cat(paste(epi[z,], collapse="\t"), file = log, append = T)
      }
      cat("\n", file = log, append = T)
      epi$strength[r] = "probable"
    }
    
    ##remove / from labels so won't interfere with parsing later
    ##remove , from labels so can convert to cSV
    epi$label = as.character(epi$label)
    epi$label = gsub("/", "|", epi$label)
    epi$label = gsub(",", "|", epi$label)
    
    ##remove duplicates
    numDup = 0
    epi$case1 = as.character(epi$case1)
    epi$case2 = as.character(epi$case2)
    epi$label = as.character(epi$label)
    r = 1
    while(r < nrow(epi)) {
      dups = which((epi$case1==epi$case1[r] & epi$case2==epi$case2[r]) | (epi$case2==epi$case1[r] & epi$case1==epi$case2[r]))
      if(length(dups) > 1) {
        numDup = numDup + length(dups) - 1
        ##keep the one with the greatest strength
        stren = epi$strength[dups]
        if(any(stren == "definite")) {
          keep = dups[stren == "definite"]
        } else if(any(stren == "probable")) {
          keep = dups[stren == "probable"]
        } else {
          keep = dups
        }
        keep = keep[1] #if multiple pairs with same strength, keep the first in the table
        epi = epi[-dups[dups!=keep],]
        if(keep != r) {
          r = r + 1
        }
      } else {
        r = r + 1
      }
    }
    if(numDup != 0) {
      cat(paste("There were", numDup, "duplicated epi links, which were removed.\n"), file = log, append = T)
    }
  }
  return(epi)
}

snpDefaultCut = 5
expectedcolnames = c("ID", "XRAYCAV", "SPSMEAR", "ExtrapulmonaryOnly", "Pediatric", "IPStart", "IPEnd") #expected columns in case data
optionalcolnames = "UserDateData" #optional columns in case data table

##output base names
defaultLogName = "LITT_log.txt"
epiFileName = "LITT_Calculated_Epi_Data.xlsx"
txFileName = "LITT_Top_Ranked_Transmission_Network.xlsx"
caseFileName = "LITT_Calculated_Case_Data.xlsx"
psFileName = "LITT_All_Potential_Sources.xlsx"
dateFileName = "LITT_Calculated_Date_Data.xlsx"

###runs the LITT analysis
###inputs:
##caseData = data frame case line list; expected column names are "ID", "XRAYCAV", "SPSMEAR", "IPStart", "IPEnd", "Extrapulmonary", "Pediatric"; "UserDateData" is optional, for TB GIMS LITT runs
##epi = data frame with columns title case1, case2, strength (strength of the epi link between cases 1 and 2), label (label for link if known) -> optional, if not given, the epi links in GIMS will be used as definite epi links
##dist = WGS SNP distance matrix (row names and column names are state case number) (optional)
##SNPcutoff = eliminate source if the SNP distance from the target is greater than this cutoff
##addlRiskFactor = data frame where one column is ID, and the rest are risk factors with Y/N values
##log = name of log file to write messages to
##progress = progress bar for R Shiny interface (NA if not running through interface)
###outputs:
##returns named list, where xxxxxxx
litt <- function(caseData, epi=NA, dist=NA, SNPcutoff = snpDefaultCut, addlRiskFactor = NA, log = defaultLogName, progress = NA) {
  ####check case data inputs
  if(!all(expectedcolnames %in% names(caseData))) {
    cat(paste("Case data table expects the following columns:", paste(expectedcolnames, collapse=", ")), file = log, append = T)
    cat(paste("\nHowever, these required columns are missing required columns in case data table: ", 
              paste(expectedcolnames[!expectedcolnames %in% names(caseData)], collapse = ", "), 
              ". Analysis was stopped. See documentation for more details.\n"), file = log, append = T)
    stop(paste("Missing required columns in caseData:", paste(expectedcolnames[!expectedcolnames %in% names(caseData)], collapse = ", ")))
  }
  caseData$IPStart = convertToDate(caseData$IPStart)
  caseData$IPEnd = convertToDate(caseData$IPEnd)
  caseData = caseData[caseData$ID!="weight",] #do not include weight (which will be in addlRiskFactor)
  if(any(is.na(caseData$IPStart))) {
    cat(paste("IP start is required, but is missing for:", paste(caseData$ID[is.na(caseData$IPStart)], collapse = ","), "\n"), 
        file = log, append = T)
    cat("These cases have been removed from the analysis\n", file = log, append = T)
  }
  if(nrow(caseData) < 2) {
    cat("LITT requires at least two cases, but there are ", nrow(caseData), 
        ". Analysis has been stopped\n", file = log, append = T)
    stop("LITT requires at least two cases, but there are ", nrow(caseData))
  }
  cases = caseData$ID
  
  ####user date data for time cutoff; assume everything from user if this column is missing (non-TB GIMS runs)
  if(!"UserDateData" %in% names(caseData)) {
    caseData$UserDateData = T
  }
  ##check user date data is T/F
  if(class(caseData$UserDateData)=="character" | class(caseData$UserDateData)=="factor") {
    caseData$UserDateData = ifelse(caseData$UserDateData=="Y" | tolower(caseData$UserDateData)=="yes" |
                                     caseData$UserDateData=="T" | tolower(caseData$UserDateData)=="true", T, F)
  }
  caseData$UserDateData[is.na(caseData$UserDateData)] = F#assume not from user if mix of missing and present values
  
  #####check extrapulomnary and pediatric variables are Y/N
  if(class(caseData$Pediatric)=="logical") {
    caseData$Pediatric = ifelse(caseData$Pediatric==T, "Y", "N")
  }
  caseData$Pediatric = as.character(caseData$Pediatric)
  bad = is.na(caseData$Pediatric) | !caseData$Pediatric %in% c("Y", "N")
  if(any(bad)) {
    cat(paste("Pediatric results are missing from the following cases, which will be assumed to be adult:", 
              paste(caseData$ID[bad], collapse = ","), "\n"), file = log, append = T)
    caseData$Pediatric[bad] = "N"
  }
  
  if(class(caseData$ExtrapulmonaryOnly)=="logical") {
    caseData$ExtrapulmonaryOnly = ifelse(caseData$ExtrapulmonaryOnly==T, "Y", "N")
  }
  caseData$ExtrapulmonaryOnly = as.character(caseData$ExtrapulmonaryOnly)
  bad = is.na(caseData$ExtrapulmonaryOnly) | !caseData$ExtrapulmonaryOnly %in% c("Y", "N")
  if(any(bad)) {
    cat(paste("Extrapulmonary only results are missing from the following cases, which will be assumed to be pulmonary:", 
              paste(caseData$ID[bad], collapse = ","), "\n"), file = log, append = T)
    caseData$ExtrapulmonaryOnly[bad] = "N"
  }
  
  
  ####calculate infectious rating
  infRate = data.frame(ID = cases,
                       infRate = NA)
  for(i in 1:nrow(infRate)) {
    if(!is.na(caseData$XRAYCAV[caseData$ID==infRate$ID[i]]) &
       caseData$XRAYCAV[caseData$ID==infRate$ID[i]] == "Y") {
      infRate$infRate[i] = 0 #cavitary, smear positive or negative
    } else if(!is.na(caseData$SPSMEAR[caseData$ID==infRate$ID[i]]) &
              caseData$SPSMEAR[caseData$ID==infRate$ID[i]] == "POS") {
      infRate$infRate[i] = 1 #smear positive, noncavitary
    } else {
      infRate$infRate[i] = 5 #smear negative, noncavitary
    }
  }
  
  ####clean epi
  epi = fixEpiNames(epi, log)
  epi = cleanEpi(epi, caseData$ID, log)
  if(nrow(epi) > 0) {
    epi$SNPdistance = getSNPDistance(epi, dist) 
  }
  cat(paste("Number of epi links:", nrow(epi)), "\n", file = log, append = T)
  
  ####set up risk factor weights
  weights = NA
  if(any(!is.na(addlRiskFactor))) {
    addlRiskFactor = addlRiskFactor[addlRiskFactor$ID %in% c(cases, "weight"),]
    addlRiskFactor[] = lapply(addlRiskFactor, as.character)
    if(any(!cases %in% addlRiskFactor$ID)) {
      cat("These cases have no risk factor values, so will be treated as N for all risk factors:\n", file = log, append = T)
      cat(paste(cases[!cases %in% addlRiskFactor$ID], collapse = ","), "\n", file = log, append = T)
    }
    ##if any NA values, set to N
    for(c in 1:ncol(addlRiskFactor)) {
      if(any(is.na(addlRiskFactor[,c]))) {
        na = is.na(addlRiskFactor[addlRiskFactor$ID!="weight",c])
        addlRiskFactor[c(na,F),c] = "N"
      }
    }
    ##look for duplicated cases
    if(any(duplicated(addlRiskFactor$ID))) {
      dup = duplicated(addlRiskFactor$ID)
      #warning if not all same
      for(i in addlRiskFactor$ID[dup]) {
        for(c in 1:ncol(addlRiskFactor)) {
          if(length(unique(addlRiskFactor[addlRiskFactor$ID==i,c])) != 1) {
            warning(paste("additional risk factor has multiple values for", i, "in column", c, 
                          "\nThe value in the first row containing", i, "will be used."))
            cat(paste("additional risk factor has multiple values for", i, "in column", c, 
                          "\nThe value in the first row containing", i, "will be used.\n"), file = log, append = T)
          }
        }
      }
      addlRiskFactor = addlRiskFactor[!dup,]
    }
    
    ##make sure state case number is first column
    if(names(addlRiskFactor)[1] != "ID") {
      addlRiskFactor = data.frame(ID = addlRiskFactor$ID,
                                  addlRiskFactor[names(addlRiskFactor)!="ID"])
    }
    
    ##set up weights (weights should add to 1)
    weights = rep(1/(ncol(addlRiskFactor)-1), ncol(addlRiskFactor)-1) #all equal if no user input
    if(any(grepl("weight", addlRiskFactor$ID, ignore.case = T))) {
      # userWeight = addlRiskFactor[grepl("weight", addlRiskFactor$ID, ignore.case = T),-1]
      # userWeight[] = lapply(userWeight, as.character)
      # if(class(userWeight)=="factor") {
      #   userWeight = as.numeric(as.character(userWeight))
      # } else {
      #   userWeight = as.numeric(userWeight[1,])
      # }
      # userWeight = as.numeric(addlRiskFactor[grepl("weight", addlRiskFactor$ID, ignore.case = T),-1]) #won't work with factors
      userWeight = getAddlRFUserUserWeights(addlRiskFactor)
      if(!all(is.na(userWeight))) { #if no weights given, use all equal
        if(any(is.na(userWeight) | userWeight <= 0)) { #remove missing or negative weights
          miss = is.na(userWeight) | (!is.na(userWeight) & userWeight <= 0) #missing or negative weight
          cat(paste("Risk factor weight is missing or negative for: ", 
                    paste(names(addlRiskFactor)[-1][miss], collapse = ", "), "\n", sep=""), file = log, append = T)
          cat(paste(ifelse(sum(miss)==1, " This risk factor has", " These risk factors have"), 
                    " been removed from analysis.\n", sep=""), file = log, append = T)
          addlRiskFactor = addlRiskFactor[,c(T, !is.na(userWeight) & userWeight > 0)]
          # userWeight = getAddlRFUserUserWeights(addlRiskFactor)
        }
        # weights = userWeight/sum(userWeight, na.rm = T) #sum to 1
      }
    }
    if(class(addlRiskFactor)=="data.frame") { #check still have some risk factors
      if(any(grepl("weight", addlRiskFactor$ID, ignore.case = T))) {
        userWeight = getAddlRFUserUserWeights(addlRiskFactor)
        weights = userWeight/sum(userWeight, na.rm = T) #sum to 1
      } else {
        weights = rep(1/(ncol(addlRiskFactor)-1), ncol(addlRiskFactor)-1) #all equal if no user input
      }
      addlRiskFactor[] = lapply(addlRiskFactor, as.character) #convert to character
      if(any(grepl("weight", addlRiskFactor$ID, ignore.case = T))) {
        addlRiskFactor[grepl("weight", addlRiskFactor$ID, ignore.case = T),-1] = weights #update table with weights used for output
      } else {
        addlRiskFactor = rbind(addlRiskFactor, c("weight", weights))
      }
      rfWeights = data.frame(variable = names(addlRiskFactor)[-1],
                             weight = as.numeric(addlRiskFactor[addlRiskFactor[,1]=="weight", -1]))
    } else {
      rfWeights = NA
      addlRiskFactor = NA
    }
  } else {
    rfWeights = NA
  }
  
  ####set up final results
  results = data.frame(target = cases,
                       sources = NA,
                       scores = NA,
                       rank1labels = " ",
                       stringsAsFactors = F)
  allSources = data.frame(target=character(),
                          source=character(),
                          snpDistance=numeric(),
                          infRate=numeric(),
                          timeRate=numeric(),
                          epiRate=numeric(),
                          label=character(),
                          score=numeric(),
                          scoreWithoutWGS=numeric(),
                          rank=numeric()) #all source matrices for all cases that pass potential source filter
  allTx = data.frame(target=character(),
                     source=character(),
                     score=numeric(),
                     scoreWithoutWGS=numeric(),
                     label=character()) #all potential sources ranked first, for visualizing LITT predicted transmission network
  allFilt = data.frame(target=character(), 
                       filteredCase=character(), 
                       reasonFiltered=character()) #all filtered cases for each target
  
  ####get ranked potential source list for every case
  for(target in cases) {
    sources = cases[cases!=target]
    
    ###remove extrapulmonary cases as source
    extrapulm = as.character(caseData$ID[caseData$ExtrapulmonaryOnly=="Y" & caseData$ID %in% sources])
    if(length(extrapulm) > 0) {
      allFilt = rbind(allFilt,
                      data.frame(target=target, filteredCase = extrapulm, reasonFiltered = "extrapulmonary"))
    }
    sources = sources[!sources %in% extrapulm]
    ##make empty data frame to add to allSources if there are no potential sources
    noSource = as.data.frame(matrix(NA, nrow=1, ncol=8))
    names(noSource) = c("snpDistance", "infRate", "timeRate", "epiRate", "label", "score", "scoreWithoutWGS", "rank")
    noSource = cbind(data.frame(target=target,
                                source="no potential sources"),
                     noSource)
    if(length(sources)==0) { ##stop if no potential sources remaining
      allSources = rbind(allSources, noSource)
      next()
    }
    
    ###remove pediatric cases as source
    ped = as.character(caseData$ID[caseData$Pediatric=="Y"  & caseData$ID %in% sources])
    if(length(ped) > 0) {
      allFilt = rbind(allFilt,
                      data.frame(target=target, filteredCase=ped, reasonFiltered="pediatric"))
    }
    sources = sources[!sources %in% ped]
    if(length(sources)==0) { ##stop if no potential sources remaining
      allSources = rbind(allSources, noSource)
      next()
    }
    
    ###get list of potential sources that are SNPcutoff or fewer SNPs from the target if both cases are sequenced
    ###if either case is not sequenced, only keep source if it is epi linked or has shared risk factor
    tEpi = epi[epi$case1==target | epi$case2==target,]
    tEpi = unique(c(tEpi$case1, tEpi$case2))
    sShareRF = NA
    if(!all(is.na(addlRiskFactor))) {
      if(target %in% addlRiskFactor$ID) {
        rf = addlRiskFactor[addlRiskFactor$ID==target,] == "Y"
        if(any(rf)) {
          sShareRF = addlRiskFactor$ID[sapply(1:nrow(addlRiskFactor),
                                              function(r) {any(addlRiskFactor[r,rf]=="Y")})]
        }
      }
    }
    if(!all(is.na(dist)) & target %in% row.names(dist)) { #target sequenced
      for(s in sources) {
        if(s %in% colnames(dist)) {#source sequenced
          if(dist[row.names(dist)==target, colnames(dist)==s] > SNPcutoff) { #above cutoff
            sources = sources[sources != s]
            allFilt = rbind(allFilt,
                            data.frame(target=target, filteredCase=s, 
                                       reasonFiltered=paste(dist[row.names(dist)==target, colnames(dist)==s], "SNPs")))
          }
        } else {
          if(!s %in% tEpi & (all(is.na(sShareRF)) | !s %in% sShareRF)) { #not sequenced or epi linked
            sources = sources[sources != s]
            allFilt = rbind(allFilt,
                            data.frame(target=target, filteredCase=s, reasonFiltered="source not sequenced and no epi"))
          }
        }
      }
    } else {
      if(any(!sources %in% tEpi)) {
        allFilt = rbind(allFilt,
                        data.frame(target=target, filteredCase=sources[!sources %in% tEpi], 
                                   reasonFiltered="given case not sequenced and no epi"))
      }
      sources = sources[sources %in% tEpi]
    }
    if(length(sources)==0) { ##stop if no potential sources remaining
      allSources = rbind(allSources, noSource)
      next()
    }
    
    ###remove additional potential sources based on infectious period/case date (different cutoffs depending on whether Sx onset present)
    ###time cutoff; different cutoffs if have more time information (more precise if have ip or sx onset)
    MoCutoffBothAvail = 3 #eliminate source if the potential source IP start is more than this number of months after target case date and symptom onset is available
    MoCutoffNoneAvail = MoCutoffBothAvail+3 #=6;eliminate source if the potential source IP start is more than this number of months after target case date and symptom onset is not available
    MoCutoffOne = ceiling((MoCutoffBothAvail + MoCutoffNoneAvail)/2) #=5; eliminate source if the potential source IP start is more than this number of months after target case date and either source or target has date data
    tAvail = caseData$UserDateData[caseData$ID==target]
    tDate = caseData$IPStart[caseData$ID==target]
    i=1
    while(i <= length(sources)) {
      s = sources[i]
      sAvail = caseData$UserDateData[caseData$ID==s]
      timeCut = ifelse(tAvail & sAvail, MoCutoffBothAvail,
                       ifelse(tAvail | sAvail, MoCutoffOne, MoCutoffNoneAvail))
      sIPstart = caseData$IPStart[caseData$ID==s]
      if(sIPstart > tDate %m+% months(timeCut)) {
        allFilt = rbind(allFilt,
                        data.frame(target=target,
                                   filteredCase=sources[i], 
                                   reasonFiltered=paste("given case earliest date - source IP start =",
                                                        round(time_length(difftime(tDate, sIPstart), "months"), digits=2),
                                                        "months with", timeCut, "month cutoff")))
        sources = sources[-i]
      } else {
        i = i + 1
      }
    }
    
    if(length(sources)==0) { ##stop if no potential sources remaining
      allSources = rbind(allSources, noSource)
      next()
    }
    
    ###set up score matrix
    score = data.frame(target=target,
                       source=sources,
                       snpDistance = NA,
                       stringsAsFactors = F)
    ###SNP distance
    score$snpDistance = sapply(1:nrow(score),
                               function(row) {
                                 return(ifelse(!all(is.na(dist)) && target %in% row.names(dist) && score$source[row] %in% colnames(dist),
                                               dist[row.names(dist)==target, colnames(dist)==score$source[row]],
                                               NA))},
                               USE.NAMES = F)
    
    ###calculate time rating
    score$timeRate = sapply(score$source, timeScore, caseData, tDate)
    
    ###calculate infectious rating
    score$infRate = sapply(1:nrow(score),
                           function(row){return(infRate$infRate[infRate$ID==score$source[row]])})
    
    ###calculate epi link strength rating
    score$epiRate = sapply(score$source, epiLinkScore, target, epi, log, USE.NAMES = F)
    label = sapply(score$source, epiLinkLabel, target, epi, USE.NAMES = F)
    
    ###update epi link strength rating with risk factor if no epi link
    if(any(!is.na(addlRiskFactor))) {
      if(target %in% addlRiskFactor$ID) {
        for(r in 1:nrow(score)) {
          if(score$epiRate[r] == noEpiScore & score$source[r] %in% addlRiskFactor$ID) {
            for(c in 2:ncol(addlRiskFactor)) {
              if(!is.na(addlRiskFactor[addlRiskFactor$ID==target, c]) &&
                 addlRiskFactor[addlRiskFactor$ID==target, c] == "Y" &&
                 !is.na(addlRiskFactor[addlRiskFactor$ID==score$source[r], c]) &&
                 addlRiskFactor[addlRiskFactor$ID==score$source[r], c] == "Y") {
                score$epiRate[r] = score$epiRate[r] - weights[c-1]
                label[r] = addlRFLabel(label[r], names(addlRiskFactor[c]))
              }
            }
          }
        }
      }
    }
    
    ###calculate score
    ##no time rating, but include source earlier, no RF
    score$label = label
    
    ##get two scores: with and without SNP distance
    score$score = getScore(score)
    score$scoreWithoutWGS = getScoreWithoutWGS(score)
    
    ###filter by score
    filtRows = (!is.na(score$snpDistance) & score$score >= 8) |
      (is.na(score$snpDistance) & score$scoreWithoutWGS >= 5)
    if(any(filtRows)) {
      allFilt = rbind(allFilt,
                      data.frame(target=target,
                                 filteredCase = score$source[filtRows], #sub("*", "", score$source[filtRows], fixed=T),#remove asterisk from non-sequenced
                                 reasonFiltered = sapply(which(filtRows),
                                                         function(r) {
                                                           if(is.na(score$score[r])) {
                                                             return(paste("Without WGS score >= 5 (time=", score$timeRate[r],
                                                                          ";infectiousness=", score$infRate[r],
                                                                          ";epi=", score$epiRate[r],
                                                                          ";score without WGS=", score$scoreWithoutWGS[r],
                                                                          ")", sep=""))
                                                           } else {
                                                             return(paste("score >= 8 (SNP=", score$snpDistance[r],
                                                                          ";time=", score$timeRate[r],
                                                                          ";infectiousness=", score$infRate[r],
                                                                          ";epi=", score$epiRate[r],
                                                                          ";score=", score$score[r],
                                                                          ")", sep=""))
                                                           }
                                                         })))
    }
    score = score[!filtRows,]
    
    ##order and rank remaining sources
    if(nrow(score)) { #may have no sources after filtering
      if(all(is.na(score$snpDistance))) { #no sequenced isolates
        score = score[order(score$scoreWithoutWGS),]
      } else if(all(!is.na(score$snpDistance))) { #all sequenced isolates
        score = score[order(score$score),]
      } else { #have mix of non-sequenced and sequenced isolates
        seqSources = score[!is.na(score$snpDistance),]
        nonSeqSources = score[is.na(score$snpDistance),]
        
        ##sort sequenced cases by score
        seqSources = seqSources[order(seqSources$score),]
        
        ##sort non sequenced cases by score without WGS
        nonSeqSources = nonSeqSources[order(nonSeqSources$scoreWithoutWGS),]
        
        ##insert non-sequenced cases before the first sequenced case with a larger score without WGS
        seqIndex = 1
        for(r in 1:nrow(nonSeqSources)) {
          ins = which(seqSources$scoreWithoutWGS[seqIndex:nrow(seqSources)] >= nonSeqSources$scoreWithoutWGS[r])[1]+seqIndex-1
          if(length(ins) < 1 | is.na(ins)) {
            seqIndex = nrow(seqSources) + 1 #nothing matches so insert at end
          } else if(seqSources$scoreWithoutWGS[ins] > nonSeqSources$scoreWithoutWGS[r]) {
            seqIndex = ins #is larger so insert as tie with previous case
          } else {
            seqIndex = ins+1 #is equal so insert as tie with this case
          }
          if(seqIndex == 1 & nrow(seqSources) > 1) {
            seqSources = rbind(seqSources[1,], nonSeqSources[r,], seqSources[2:nrow(seqSources),])
          } else if(seqIndex == 1 & nrow(seqSources) == 1) {
            seqSources = rbind(seqSources[1,], nonSeqSources[r,])
          } else if(seqIndex <= nrow(seqSources)) {
            seqSources = rbind(seqSources[1:(seqIndex-1),], nonSeqSources[r,], seqSources[seqIndex:nrow(seqSources),])
          } else {
            seqSources = rbind(seqSources, nonSeqSources[r:nrow(nonSeqSources),])
            break
          }
          seqIndex = seqIndex + 1 
        }
        score = seqSources
      }
      
      ##add rank to score
      rank = 1
      score$rank = rank
      if(nrow(score) > 1) {
        if(!all(is.na(score$score))) {
          dup = duplicated(score$score)
          for(r in 2:nrow(score)) {
            # if(!is.na(score$score[r]) & (is.na(score$score[r-1]) | score$score[r-1] < score$score[r])) {
            if(!is.na(score$score[r]) & !dup[r]) { #source is not tied with the previous source (assumes score was sorted) or was a non-sequenced case in a mix of sequenced and non-sequenced cases not also tie
              rank = r
            }
            score$rank[r] = rank
          }
        } else { #order by score without WGS if all cases have no sequence
          for(r in 2:nrow(score)) {
            ##increase rank to row number (number of previous cases) if score is higher or previous row is NA (no sequence) but this row is not
            if(score$scoreWithoutWGS[r-1] < score$scoreWithoutWGS[r]) {
              rank = r
            }
            score$rank[r] = rank
          }
        }
      }
      score$rank[is.na(score$snpDistance)] = paste(score$rank[is.na(score$snpDistance)], "*", sep="")
      
      ###output results
      allSources = rbind(allSources, score)
      
      ##update transmission network
      r1 = score[score$rank=="1" | score$rank=="1*" | score$rank==1,]
      allTx = rbind(allTx,
                    data.frame(target = target,
                               source = r1$source,
                               score = r1$score,
                               scoreWoWGS = r1$scoreWithoutWGS,
                               label=r1$label,
                               stringsAsFactors = F),
                    stringsAsFactors = F)
      
      ###update final results table
      coll = score$source[1]#collapsed list of sources
      sc = score$score[1]#collapsed list of scores
      if(nrow(score) > 1) {
        if(!all(is.na(score$score))) {
          dup = duplicated(score$score)
          for(i in 2:nrow(score)) {
            if(is.na(score$score[i]) || dup[i]) { #source is tied with the previous source (assumes score was sorted) or was a non-sequenced case in a mix of sequenced and non-sequenced cases also tie
              coll = paste(coll, score$source[i], sep="/")
            } else {
              coll = paste(coll, score$source[i], sep=";")
              sc = paste(sc, score$score[i], sep=";")
            }
          }
        } else { #order by score without WGS if all cases have no sequence
          sc = paste(score$scoreWithoutWGS[1], "*", sep="")
          dup = duplicated(score$scoreWithoutWGS)
          for(i in 2:nrow(score)) {
            if(dup[i]) { #source is tied with the previous source (assumes score was sorted)
              coll = paste(coll, score$source[i], sep="/")
            } else {
              coll = paste(coll, score$source[i], sep=";")
              sc = paste(sc, paste(score$scoreWithoutWGS[i], "*", sep=""), sep=";")
            }
          }
        }
      } else if(all(is.na(score$score))) {
        sc = paste(score$scoreWithoutWGS[1], "*", sep="")
      }
      if(!is.na(coll)) { 
        ##rank 1 label
        sp = strsplit(coll, ";", fixed = T)[[1]][1]
        sp = strsplit(sp, "/", fixed=T)[[1]]
        r1lab = score$label[score$source==sp[1]]
        if(length(sp) > 1) {
          for(i in 2:length(sp)) {
            r1lab = paste(r1lab, score$label[score$source==sp[i]], sep="/")
          }
        }
        results$rank1labels[results$target==target] = r1lab
      }
      results$sources[results$target==target] = coll
      results$scores[results$target==target] = sc
    } else {
      allSources = rbind(allSources, noSource)
    }
  }
  
  ####return results
  results = results[order(results$target),]
  return(list(allPotentialSources = allSources, 
              filteredSources = allFilt, 
              topRanked = allTx, 
              summary=results, 
              epi = epi,
              rfWeights = rfWeights))
}
