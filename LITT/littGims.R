##sets up LITT using TB GIMS data
source("litt.R")

library(haven) #to read GIMS data

##get patient data
gimsPtGenoFolder = "\\\\cdc.gov\\project\\NCHHSTP_DTBE_GENO\\TBGIMS\\GIMS data\\"
gimsPtFiles = list.files(gimsPtGenoFolder)
gimsPtFiles = gimsPtFiles[grepl("gims_patient_[0-9]*.sas7bdat", gimsPtFiles)]
fileinfo = file.info(paste(gimsPtGenoFolder, gimsPtFiles, sep=""))
gimsPtName = row.names(fileinfo)[fileinfo$ctime==max(fileinfo$ctime)] #ctime = creation time
cat(paste("GIMS patient file: ", gimsPtName, "\r\n"))
gimsPt = read_sas(gimsPtName)
gimsPtName = sub(gimsPtGenoFolder, "", gimsPtName, fixed=T)

##get all gims export data
gimsExportFolder = "\\\\cdc.gov\\project\\NCHHSTP_DTBE_GENO\\TBGIMS\\Surveillance\\"
gimsExportFiles = list.files(gimsExportFolder)
gimsAllFiles = gimsExportFiles[grepl("^gims_export_[0-9]*.sas7bdat$", gimsExportFiles)]
fileinfo = file.info(paste(gimsExportFolder, gimsAllFiles, sep=""))
gimsAllName = row.names(fileinfo)[fileinfo$ctime==max(fileinfo$ctime)] #ctime = creation time
cat(paste("GIMS export file:", gimsAllName, "\r\n"))
gimsAll = read_sas(gimsAllName)
names(gimsPt) = tolower(names(gimsPt)) #sometimes is Stcaseno, sometimes stcaseno
gimsAll = merge(gimsAll,
                data.frame(STCASENO=gimsPt$stcaseno,
                           sp_coll_date=gimsPt$sp_coll_date),
                by="STCASENO",
                all = T)
gimsAllName = sub(gimsExportFolder, "", gimsAllName, fixed=T)

##get genotyping data
gimsPtFiles = list.files(gimsPtGenoFolder)
gimsGenoFiles = gimsPtFiles[grepl("gims_genotype_[0-9]*.sas7bdat", gimsPtFiles)]
fileinfo = file.info(paste(gimsPtGenoFolder, gimsGenoFiles, sep=""))
gimsGenoName = row.names(fileinfo)[fileinfo$ctime==max(fileinfo$ctime)] #ctime = creation time
cat(paste("GIMS genotyping file:", gimsGenoName, "\r\n"))
gimsGeno = read_sas(gimsGenoName)
gimsGenoName = sub(gimsPtGenoFolder, "", gimsGenoName, fixed=T)

###function that replaces all NA values in a vector with an empty string (used for epi links)
replaceMissing <- function(values) {
  values[is.na(values) | values == " "] = ""
  return(values)
}


###function that calculates symptom onset from infectious period start if sympmtom onset not available
##sxCases is the list of cases corresponding to the sxOnsetDates
##sxOnsetDates is the current list of sxOnset dates
##ipCases is the list of cases corresponding to the ipStart
##ipStart is the user uploaded list of infectious period start dates
calcSxOnset <- function(sxCases, sxOnsetDates, ipCases, ipStart) {
  ipStart = convertToDate(ipStart)
  for(i in 1:length(sxOnsetDates)) {
    if(is.na(sxOnsetDates[i]) & sxCases[i] %in% ipCases) {
      sxOnsetDates[i] = ipStart[ipCases==sxCases[i]] %m+% months(3)
    }
  }
  return(sxOnsetDates)
}

##function that takes the given data frame and corrects the state case number capitalization to be uniform across the program
fixStcasenoName <- function(df) {
  if(any(!is.na(df))) {
    df = df[!apply(df, 1, function(x) all(is.na(x))),
            !apply(df, 2, function(x) all(is.na(x)))] #remove rows and columns that are all NA
    names(df)[grepl("stcaseno", names(df), ignore.case = T) |
                grepl("state[ .]*case[ .]*number", names(df), ignore.case = T)] = "STCASENO"
    df$STCASENO = removeWhitespacePeriods(df$STCASENO)
  }
  return(df)
}


##check symptom onset column names; if none present, return NA
##check for duplicates; if same isolate has two different sx onset dates, give warning and take the earlier
fixSxOnsetNames <- function(df, log) {
  if(any(!is.na(df))) {
    col = grepl("sx[ ]*onset", names(df), ignore.case = T) | grepl("symptom[ ]*onset", names(df), ignore.case = T)
    if(sum(col) == 1) {
      names(df)[col] = "sxOnset"
    } else if(sum(col) > 1) {
      warning(paste("More than one column for symptom onset:", paste(names(df)[col], collapse = ", ")),
              "\r\nPlease label one column sxOnset.\r\n Symptom onset was not used in this analysis.")
      cat(paste("More than one column for symptom onset:", paste(names(df)[col], collapse = ", ")),
          "\r\nPlease label one column sxOnset.\r\n Symptom onset was not used in this analysis.\r\n", 
          file = log, append = T)
      return(NA)
    } else if(sum(col) < 1) {
      cat("No symptom onset column, so symptom onset was not used in the analysis.\r\nPlease add a column named symptom onset or sxOnset to use symptom onset in analysis.\r\n", 
          file = log, append = T)
      return(NA)
    }
    ##look for duplicates
    df$STCASENO = as.character(df$STCASENO)
    dup = duplicated(df$STCASENO)
    if(any(dup)) {
      ids = df$STCASENO[dup]
      for(i in ids) {
        rows = which(df$STCASENO==i)
        df$sxOnset[rows[1]] = getMin(df$sxOnset[df$STCASENO==i], i, "symptom onset")
        df = df[-rows[-1],]
      }
    }
  }
  return(df)
}

##function to calculate IP end for visualizations
calculateIPEnd <- function(stno, gimsCases, locIPEnd=NA) {
  ##no IP end for extrapulmonary cases
  if(!((!is.na(gimsCases$SITEPULM[gimsCases$STCASENO==stno]) & 
        gimsCases$SITEPULM[gimsCases$STCASENO==stno]=="Y") |
       (!is.na(gimsCases$SITELARYN[gimsCases$STCASENO==stno]) &
        gimsCases$SITELARYN[gimsCases$STCASENO==stno]=="Y"))) {
    return(as.Date(NA))
  } else {
    rx = gimsCases$RXDATE[gimsCases$STCASENO==stno]
    if(!is.na(locIPEnd) & (is.na(rx) | locIPEnd > rx)) {
      return(locIPEnd)
    } else if(!is.na(rx)) {
      return(rx + weeks(2))
      ##hopefully, most cases will be these first two conditions; rest are only if no rxdate available
    } else if((!is.na(gimsCases$THERREAS[gimsCases$STCASENO==stno]) & 
               gimsCases$THERREAS[gimsCases$STCASENO==stno] == "DIED") | #if patient died on treatment, use stop tx date
              ((!is.na(gimsCases$STATUS[gimsCases$STCASENO==stno]) &
                gimsCases$STATUS[gimsCases$STCASENO==stno]=="DEAD"))) { #if patient was already dead, use death date
      stopther = convertToDate(gimsCases$STOPTHER[gimsCases$STCASENO==stno])
      if(!is.na(stopther)) {
        return(stopther)
      } else {
        dd = convertToDate(gimsCases$DEATHDATE[gimsCases$STCASENO==stno])
        if(is.na(dd)) { #use report date if no death date/stop therapy date
          return(gimsCases$RPTDATE[gimsCases$STCASENO==stno])
        } else {
          return(dd)
        }
      }
    } else if(!is.na(gimsCases$THERREAS[gimsCases$STCASENO==stno]) & 
              gimsCases$THERREAS[gimsCases$STCASENO==stno] == "COMPLETED") { #if therapy completed, use latest of earliest date set of dates, plus two weeks
      row = gimsCases$STCASENO==stno
      return(max(gimsCases$ISUSDATE[row], gimsCases$RXDATE[row], gimsCases$CNTDATE[row], gimsCases$RPTDATE[row],
                 gimsCases$sp_coll_date[row], na.rm=T) + weeks(2))
    } else { #example, didn't start therapy
      return(as.Date(NA))
    }
  }
}

##function that returns whether a case (sid) is in the input sxOnset or infectiousPeriod data
##used to calculate time cutoff
caseInInputs <- function(sid, sxOnset, infectiousPeriod) {
  if(any(!is.na(sxOnset)) && 
     sid %in% sxOnset$STCASENO && 
     !is.na(sxOnset$sxOnset[sxOnset$STCASENO==sid])) {
    return(T)
  } else {
    return(any(!is.na(infectiousPeriod)) 
           && sid %in% infectiousPeriod$STCASENO && 
             !is.na(infectiousPeriod$IPStart[infectiousPeriod$STCASENO==sid]))
  }
}

##returns all accession numbers for the given state case number sno
getAccessionNumber <- function(stno) {
  acc = gimsGeno$accessionnumber[gimsGeno$StCaseNo==stno]
  acc = acc[!is.na(acc) & acc != ""]
  acc = paste(acc, collapse=",")
  return(acc)
}

##cleans up and writes case data table
##caseOut = case output data to format
##outPrefix = output prefix
cleanCaseOutput <- function(caseOut, outPrefix) {
  ###convert T/F to Y/N
  caseOut$UserDateData = ifelse(caseOut$UserDateData, "Y", "N")
  caseOut$sequenceAvailable = ifelse(caseOut$sequenceAvailable, "Y", "N")
  
  ##format date/zip columns
  caseOut$zipcode = as.numeric(as.character(caseOut$zipcode))
  # caseOut$earliestDate = format(caseOut$earliestDate, format="%m/%d/%Y")
  caseOut$IPStart = format(caseOut$IPStart, format="%m/%d/%Y")
  caseOut$IPEnd = format(caseOut$IPEnd, format="%m/%d/%Y")
  caseOut$IAS = format(caseOut$IAS, format="%m/%d/%Y")
  caseOut$IAE = format(caseOut$IAE, format="%m/%d/%Y")
  
  writeExcelTable(fileName=paste(outPrefix, caseFileName, sep=""),
                  sheetName = "case data",
                  df = caseOut,
                  wrapHeader = T, 
                  stcasenolab = T,
                  gims = T)
}

##for the given file, return a distance matrix that can be passed to the dist variable in LITT
##returns a matrix: fill in whole matrix, with row and column names the state case number and SNP distances rounded to nearest whole number
##the fileName is expected to point to a SNP distance matrix or lower triangle, such as from BioNumerics (include file path if not in working dir)
##assumption: text files may be from BioNumerics and therefore may not have the column headers; all other files formats it is safe to assume will have the header
##log is file to write results to
##if appendlog is true, append results to log file; otherwise overwrite (needed because may read distance matrix first)
formatBNDistanceMatrix <- function(fileName, log, appendlog=T) { #formerly formatDistanceMatrix
  cat("", file=log, append=appendlog)
  ##read file
  if(endsWith(fileName, ".txt") || endsWith(fileName, ".tsv")) {
    mat = as.matrix(read.table(fileName, sep="\t", header = T, row.names = 1))
    ##Bionumerics does not provide the column headers
    if(ncol(mat)==1) { #bionumerics seems to tab separate the first name then space separate the rest; does not generate columns for the upper triangle and so will get error if more columns than column names
      df = as.matrix(read.table(fileName, sep="\t", header = F, row.names = 1))
      mat = matrix(nrow = nrow(df), ncol = nrow(df), NA)
      for(r in 1:nrow(df)) {
        snps = as.numeric(as.character(strsplit(trimws(df[r,]), "\\s+")[[1]]))
        for(c in 1:length(snps)) {
          mat[r,c] = snps[c]
        }
      }
      row.names(mat) = row.names(df)
    } #else if(make.names(row.names(mat))[1] != colnames(mat)[1]) {
      # mat = as.matrix(read.table(fileName, sep="\t", header = F, row.names = 1))
    # }
    colnames(mat) = row.names(mat) #fix the . for non character spaces
  } else if(endsWith(fileName, ".xls") || endsWith(fileName, ".xlsx")) {
    df = read.xlsx(fileName, sheetIndex = 1)
    df = df[!apply(df, 1, function(x) all(is.na(x))),
              !apply(df, 2, function(x) all(is.na(x)))] #remove rows and columns that are all NA
    mat = as.matrix(df[,-1])
    row.names(mat) = df[,1]
    colnames(mat) = as.character(df[,1]) #fix the . for non character spaces
  } else if(endsWith(fileName, ".csv")) {
    mat = as.matrix(read.table(fileName, sep=",", header = T, row.names = 1))
    ##remove rows and columns that are all NA
    mat = mat[!apply(mat, 1, function(x) all(is.na(x))),
              !apply(mat, 2, function(x) all(is.na(x)))]
    colnames(mat) = row.names(mat) #fix the . for non character spaces
  } else {
    cat("Distance matrix should be in a text, Excel or CSV file format.\r\n", file = log, append = T)
    return(NA)
  }
  colnames(mat) = sub("^X", "", colnames(mat))
  colnames(mat) = removeWhitespacePeriods(colnames(mat))
  row.names(mat) = removeWhitespacePeriods(row.names(mat))
  
  ##remove rows and columns that are all NA
  mat = mat[!apply(mat, 1, function(x) all(is.na(x))),
            !apply(mat, 2, function(x) all(is.na(x)))]
  
  if(nrow(mat) != ncol(mat)) {
    cat("Distance matrix has ", nrow(mat), " rows but ", ncol(mat), 
        " columns. It must have the same number of rows and columns containing SNP distance.\r\n", file = log, append = T)
    stop("Distance matrix must have the same number of rows and columns containing SNP distance.")
  }
  
  ##fill in matrix if needed
  if(all(is.na(mat[upper.tri(mat)]))) {
    mat[upper.tri(mat)] = t(mat)[upper.tri(mat)]
  }
  
  ##round to bring the .99 to whole number
  mat = round(mat)
  
  ##convert to state case number
  sid = row.names(mat)
  acc = row.names(mat)
  for(i in 1:length(sid)) {
    if(sid[i]!="MRCA") {
      if(!sid[i] %in% gimsAll$STCASENO) { #is accession number
        id = unique(gimsGeno$StCaseNo[gimsGeno$accessionnumber==acc[i]])
        if(length(id) == 0 || id == "NA" || is.na(id) || id == "") {
          cat(paste("No state case number for", acc[i], "so it will be excluded.\r\n"), file = log, append = T)
          sid[i] = NA
        } else if(length(id) > 1) {
          cat(paste("Multiple state case numbers for", acc[i], "so it will be excluded.\r\n"), file = log, append = T)
          sid[i] = NA
        } else {
          sid[i] = id
        }
      }
    }
  }
  row.names(mat) = sid
  colnames(mat) = sid
  
  ##remove NAs
  acc = acc[!is.na(row.names(mat))]
  mat = mat[!is.na(row.names(mat)), !is.na(colnames(mat))]
  acc = acc[row.names(mat)!=""]
  mat = mat[row.names(mat)!="", colnames(mat)!=""] #no state case number available yet
  sid = row.names(mat)
  
  ##check for duplicated isolates (two isolates with same stcaseno but different accno)
  ##remove the later accession number if multiple sequences for the same case
  if(any(duplicated(sid))) {
    dup = unique(sid[duplicated(sid)])
    for(d in dup) {
      dacc = sort(acc[sid==d])
      ##check the rows are the same and give a warning otherwise
      for(i in 2:length(dacc)) {
        if(!all(mat[acc==dacc[i],] == mat[acc==dacc[i-1],])) {
          warning(dacc[i-1], " and ", dacc[i], " have different SNP distances in ", fileName, " but are the same patient (", d, "); ", 
                  dacc[i-1], " will be used.")
          cat(dacc[i-1], " and ", dacc[i], " have different SNP distances in ", fileName, " but are the same patient (", d, "); ", 
              dacc[i-1], " will be used.\r\n", file = log, append = T)
        }
      }
      mat = mat[!acc %in% dacc[2:length(dacc)],!acc %in% dacc[2:length(dacc)]]
      sid = sid[!acc %in% dacc[2:length(dacc)]]
      acc = acc[!acc %in% dacc[2:length(dacc)]]
    }
  }
  
  ##stop if there are missing values
  if(any(is.na(mat))) {
    cat("There is missing data in the distance matrix. All of the lower triangle must be filled in.\r\n", file = log, append = T)
    stop("Distance matrix is missing data.")
  }
  
  return(mat)
}

###sets up the inputs for LITT then calls LITT then formats the outputs
###inputs:
##outPrefix = prefix for writing results (default is none); if a different directory is desired, provide that here
##cases = state case number of cases to be included in the analysis (if not provided, it is the list of all cases in the inputs)
##dist = WGS SNP distance matrix (row names and column names are state case number) (optional)
##caseData = combined dataframe with symptom onset, infectious period and additional risk factor data; required field, if missing, call meaAlgorithm directly
##epi = data frame with columns title case1, case2, strength (strength of the epi link between cases 1 and 2), label (label for link if known) -> optional, if not given, the epi links in GIMS will be used as definite epi links
##gimsRiskFactor = data frame with first column titled variable and optional second column labeled weight that lists the GIMS variables to use as additional risk factors, with their weights (positive weights are inlcuded in LITT and output; zero in case line list only; negative is not output)
##SNPcutoff = eliminate source if the SNP distance from the target is greater than this cutoff
##writeDate = if true, write date table to Excel and include in list of outputs
##writeDist = if true, write distance matrix to Excel and include in list of outputs
##if appendlog is true, append results to log file; otherwise overwrite (needed because may read distance matrix first)
##progress = progress bar for R Shiny interface (NA if not running through interface)
littGims <- function(outPrefix = "", cases=NA, dist=NA, caseData, epi=NA, rfTable = NA, gimsRiskFactor = NA, 
                     writeDate = F, writeDist = F, appendlog=F, SNPcutoff = snpDefaultCut, progress = NA) {
  log = paste(outPrefix, defaultLogName, sep="")
  cat("LITT analysis with TB GIMS\r\nSNP cutoff = ", SNPcutoff, "\r\n", file = log, append=appendlog)
  cat(paste("GIMS patient file: ", gimsPtName, "\r\n"), file = log, append = T)
  cat(paste("GIMS export file:", gimsAllName, "\r\n"), file = log, append = T)
  cat(paste("GIMS genotyping file:", gimsGenoName, "\r\n"), file = log, append = T)
  
  ####check inputs
  ##fix field names
  names(caseData) = removeWhitespacePeriods(names(caseData))
  caseData = fixStcasenoName(caseData)
  caseData = fixPresumedSource(caseData)
  epi = fixEpiNames(epi, log)
  colnames(dist) = removeXFromNames(colnames(dist))
  
  ###split
  sxOnset = fixSxOnsetNames(caseData, log) #data frame with columns titled stcaseno (state case number) and sxOnset (symptom onset or earliest diagnostic finding) -> optional ideal input
  if("sxOnset" %in% names(sxOnset)) {
    sxOnset = sxOnset[,names(sxOnset) %in% c("STCASENO", "sxOnset")]
  }
  names(caseData)[names(caseData)=="STCASENO"] = "ID" #data frame with columns titled stcaseno (state case number), IPStart (date of start of infectious period) and IPEnd (date of end of infectious period) -> optional ideal input if sx onset not available
  infectiousPeriod = fixIPnames(caseData, log)
  names(caseData)[names(caseData)=="ID"] = "STCASENO"
  if("IPStart" %in% names(infectiousPeriod)) {
    names(infectiousPeriod)[names(infectiousPeriod)=="ID"] = "STCASENO"
    infectiousPeriod = infectiousPeriod[,names(infectiousPeriod) %in% c("STCASENO", "IPStart", "IPEnd")]
  }
  iae = fixIAE(caseData) #for printing
  if("IAE" %in% names(iae)) {
    names(iae)[names(iae)=="ID"] = "STCASENO"
    iae = iae[,c("STCASENO", "IAS", "IAE")]
  } else {
    iae=NA
  }
  
  ####split out additional risk factors
  addlRiskFactor=NA #data frame with first column titled stcaseno (state case number) and remaining columns represent additional variables with Y or N for additional risk variables (ex Y or N for whether the case was in a homeless shelter) -> optional
  # if(!all(names(caseData) %in% c("STCASENO", "sxOnset", "IPStart", "IPEnd"))) { #have RFs
  #   addlRiskFactor=caseData[,!names(caseData) %in% c("sxOnset", "IPStart", "IPEnd")]
  # }
  addlRiskFactor=NA #risk factor data, with row for weight
  printVars = NA #additional variables to print but not standard variables or risk factors
  if(!all(is.na(rfTable))) {
    rfTable = fixRfTable(rfTable, log)
    if(!all(is.na(rfTable))) {
      if(any(!rfTable$variable %in% names(caseData))) { #extra variables not in case data table
        miss = !rfTable$variable %in% names(caseData)
        cat("The following variables are in the risk factor table but not in the case data table, so will not be used in analysis: ",
            paste(rfTable$variable[miss], collapse=", "), "\r\n", file = log, append = T)
        rfTable = rfTable[!miss,]
      }
      if(any(!names(caseData) %in% c(expectedcolnames, optionalcolnames, rfTable$variable))) {
        miss = !names(caseData) %in% c(expectedcolnames, optionalcolnames, rfTable$variable)
        cat("There are extra columns in the case data table, which will be removed: ",
            paste(names(caseData)[miss], collapse=", "), "\r\n", file = log, append = T)
      }
      ##set up print table (weight 0)
      rfTable = rfTable[!is.na(rfTable$weight),]
      if(any(rfTable$weight==0)) {
        printVars = caseData[,names(caseData) %in% c("ID", "STCASENO", rfTable$variable[rfTable$weight==0])]
      }
      rfTable = rfTable[rfTable$weight > 0,]
      ##set up additional risk factor table
      if(nrow(rfTable) > 0) {
        addlRiskFactor = caseData[,names(caseData) %in% c("ID", "STCASENO", rfTable$variable)]
        addlRiskFactor[] = lapply(addlRiskFactor, as.character)
        for(c in 1:ncol(addlRiskFactor)) {
          if(names(addlRiskFactor)[c]!="ID") {
            addlRiskFactor[,c] = toupper(addlRiskFactor[,c])
            addlRiskFactor[grepl("^yes$", addlRiskFactor[,c], ignore.case = T),c] = "Y"
            addlRiskFactor[grepl("^no$", addlRiskFactor[,c], ignore.case = T),c] = "N"
          }
        }
        addlRiskFactor = rbind(addlRiskFactor, c("weight", rfTable$weight))
      }
    }
  }
  caseData=caseData[,names(caseData) %in% c(expectedcolnames, optionalcolnames, "STCASENO")]
  
  ####get list of cases
  if(all(is.na(cases))) {
    cases = vector()
    if(any(!is.na(dist))) {
      cases = colnames(dist)
    }
    if(any(!is.na(sxOnset))) {
      cases = c(cases, as.character(sxOnset$STCASENO))
    }
    if(any(!is.na(infectiousPeriod))) {
      cases = c(cases, as.character(infectiousPeriod$STCASENO))
    }
    if(any(!is.na(epi))) {
      cases = c(cases, as.character(epi$case1), as.character(epi$case2))
    }
    if(any(!is.na(addlRiskFactor))) {
      cases = c(cases, as.character(addlRiskFactor$STCASENO))
      cases = cases[!grepl("weight", cases, ignore.case = T)] #make sure weight of additional RF not in case list
    }
    if(!all(is.na(cases))) {
      cases = sort(unique(cases))
    }
  }
  cases = cases[!grepl("MRCA", cases, ignore.case = T) & !grepl("case[ ]*no", cases, ignore.case = T) &
                  cases != "" & !is.na(cases)] #remove MRCA and stcaseno guidance from case list
  ##check all cases are in TB GIMS, and if not remove
  if(!all(is.na(cases))) {
    if(any(!(cases %in% gimsAll$STCASENO))) {
      cat(paste("TB GIMS is missing the following cases, which will be removed from the analysis:", 
                paste(cases[!(cases %in% gimsAll$STCASENO)], collapse = ", "), "\r\n"), file = log, append = T)
      cases = cases[cases %in% gimsAll$STCASENO]
    }
  }
  if(all(is.na(cases)) || length(cases) < 2) { #no cases to analyze
    cat("There must be at least 2 cases with state case numbers in TB GIMS. Analysis has been stopped.\r\n", file = log, append = T)
    stop("There must be at least 2 cases with state case numbers in TB GIMS.\r\n")
  }
  cat(paste("Number of cases:", length(cases), "\r\n"), file = log, append = T)
  
  gimsCases = gimsAll[gimsAll$STCASENO %in% cases,]
  
  if(all(class(progress)!="logical")) {
    progress$set(value = -1)
  }
  
  ####add needed case data from TB GIMS: smear, cavitary, extrapulmonary, pediatric
  littCaseData = gimsCases[, names(gimsCases) %in% c("STCASENO", "SPSMEAR", "XRAYCAV")]
  littCaseData$SPSMEAR[is.na(littCaseData$SPSMEAR)] = ""
  littCaseData$XRAYCAV[is.na(littCaseData$XRAYCAV)] = ""
  littCaseData$ExtrapulmonaryOnly = sapply(littCaseData$STCASENO, 
                                           function(s){
                                             return(ifelse(!((!is.na(gimsCases$SITEPULM[gimsCases$STCASENO==s]) & 
                                                                gimsCases$SITEPULM[gimsCases$STCASENO==s]=="Y") |
                                                               (!is.na(gimsCases$SITELARYN[gimsCases$STCASENO==s]) &
                                                                  gimsCases$SITELARYN[gimsCases$STCASENO==s]=="Y")), 
                                                           "Y", "N"))})
  littCaseData$ExtrapulmonaryOnly[is.na(littCaseData$ExtrapulmonaryOnly)] = "N"
  littCaseData$Pediatric = sapply(littCaseData$STCASENO,
                                  function(s) {
                                    return(ifelse(gimsCases$AGE[gimsCases$STCASENO==s] < 10, "Y", "N")) #note this includes cases missing an age
                                  })
  littCaseData$Pediatric[is.na(littCaseData$Pediatric)] = "N"
  
  ####set up dates
  ##case dates
  gimsCases$ISUSDATE = convertToDate(gimsCases$ISUSDATE)
  gimsCases$RXDATE = convertToDate(gimsCases$RXDATE)
  gimsCases$CNTDATE = convertToDate(gimsCases$CNTDATE)
  gimsCases$RPTDATE = convertToDate(gimsCases$RPTDATE)
  gimsCases$sp_coll_date = convertToDate(gimsCases$sp_coll_date)
  ##sometimes specimen collection date is 1/1/1900 instead of blank if missing; set to missing if more than 1 year from report date
  gimsCases$sp_coll_date[gimsCases$sp_coll_date < gimsCases$RPTDATE %m-% years(1) | 
                           gimsCases$sp_coll_date > gimsCases$RPTDATE %m+% years(1)] = as.Date(NA)
  dates = data.frame(gimsCases$STCASENO, 
                     gimsCases$ISUSDATE, gimsCases$RXDATE, gimsCases$CNTDATE, gimsCases$RPTDATE, gimsCases$sp_coll_date,
                     gimsCases$PEDAGE,
                     stringsAsFactors = F)
  names(dates) = sub("gimsCases.", "", names(dates))
  inputIP = infectiousPeriod #for writing
  inputSx = sxOnset
  
  ##update earliest date with symptom onset if available
  if(any(!is.na(sxOnset))) {
    # sxOnset = fixSxOnsetNames(sxOnset, log)
    sxOnset = sxOnset[sxOnset$STCASENO %in% cases,]
    sxOnset$sxOnset = convertToDate(sxOnset$sxOnset)
    sxOnset = sxOnset[!is.na(sxOnset$sxOnset),]
    inputSx = sxOnset
    dates = merge(dates, sxOnset, by = "STCASENO", all.x=T)
  } else {
    dates$sxOnset = as.Date(NA)
  }
  
  # ##update earliest date with infection acquisition end if available
  # if(!all(is.na(iae))) {
  #   iae3 = iae
  #   iae3$IAE = iae$IAE %m+% months(3) #because 3 months will be subtracted if it is the earliest date
  #   dates = merge(dates, iae3, by = "STCASENO", all.x=T)
  # }
  ##merge infectious acqusition end with infectious period if available
  # if(!all(is.na(iae))) {
  #   iae = iae[iae$STCASENO %in% cases,]
  #   iae$STCASENO = as.character(iae$STCASENO)
  #   infectiousPeriod$STCASENO = as.character(infectiousPeriod$STCASENO)
  #   if(all(is.na(infectiousPeriod))) {
  #     infectiousPeriod = iae
  #     names(infectiousPeriod)[2] = "IPStart"
  #     infectiousPeriod$IPEnd = as.Date(NA)
  #   } else {
  #     for(r in 1:nrow(iae)) {
  #       if(!iae$STCASENO[r] %in% infectiousPeriod$STCASENO) { #have IAE but not in IP
  #         infectiousPeriod = rbind(infectiousPeriod,
  #                                  data.frame(STCASENO = iae$STCASENO[r],
  #                                             IPStart = iae$IAE[r],
  #                                             IPEnd = as.Date(NA)))
  #       } else if(!is.na(iae$IAE[r])) { #set IPStart to min of IAE and IP start
  #         ips = infectiousPeriod$IPStart[infectiousPeriod$STCASENO==iae$STCASENO[r]]
  #         if(!is.na(ips) & ips != iae$IAE[r]) {
  #           cat(paste0("Infectious period start and infection acquisition end were both provided for ", iae$STCASENO[r], 
  #               ", but they are not the same. The earlier of the two dates will be used as IP Start.\r\n"), file = log, append = T)
  #         }
  #         infectiousPeriod$IPStart[infectiousPeriod$STCASENO==iae$STCASENO[r]] = min(iae$IAE[r],ips, na.rm = T)
  #       }
  #     }
  #   }
  # }
  ##merge IP and infection acquisition end window if needed
  if(!all(is.na(iae))) {
    temp = merge(littCaseData, iae, by="STCASENO")
    if(!all(is.na(infectiousPeriod))) {
      temp = merge(temp, infectiousPeriod, by= "STCASENO")
    }
    infectiousPeriod = mergeIAEtoIPstart(temp, log)
    infectiousPeriod = infectiousPeriod[,c("STCASENO", "IPStart", "IPEnd")]
  }
  
  ##if infectious period start is provided but symptom onset is missing, define symptom onset as IP start + 3 months
  if((all(is.na(sxOnset)) | any(is.na(dates$sxOnset))) & #missing symptom onset dates
     any(!is.na(infectiousPeriod))) { #but have IP start
    infectiousPeriod = infectiousPeriod[infectiousPeriod$STCASENO %in% cases,]
    dates$sxOnset = calcSxOnset(dates$STCASENO, dates$sxOnset, infectiousPeriod$STCASENO, infectiousPeriod$IPStart)
  }
  # dates$earliestDate = convertToDate(apply(dates[,-1], 1, min, na.rm=T))
  dates$earliestDate = convertToDate(apply(dates[,!names(dates) %in% c("STCASENO", "PEDAGE")], 1, min, na.rm=T))
  
  ##calculate infectious period start
  if(any(!is.na(infectiousPeriod))) {
    infectiousPeriod = infectiousPeriod[infectiousPeriod$STCASENO %in% cases,]
    dates = merge(dates, infectiousPeriod, by="STCASENO", all.x = T)
    dates$IPStart = convertToDate(dates$IPStart)
    dates$IPEnd = convertToDate(dates$IPEnd)
    ##check to see if IP start is after earliest date (and therefore incorrect)
    if(any(dates$IPStart > dates$earliestDate %m-% months(1), na.rm = T)) {
      r = which(dates$IPStart > dates$earliestDate %m-% months(1))
      cat("The following isolates have an infectious period start later than 1 month prior to the earliest date; 3 months prior to the earliest date will be used instead:\r\n", 
          file = log, append = T)
      cat(paste(names(dates), collapse ="\t"), file = log, append = T)
      for(z in r) {
        cat("\r\n", dates[z,1], file = log, append = T)
        for(c in 2:ncol(dates)) {
          cat(paste0("\t", as.character(dates[z,c])), file = log, append = T)
        }
      }
      cat("\r\n\r\n", file = log, append = T)
      dates$IPStart[r] = as.Date(NA)
    }
    if(any(dates$IPEnd < dates$RXDATE, na.rm = T)) {
      r = which(dates$IPEnd < dates$RXDATE & !is.na(dates$RXDATE))
      cat("The following isolates have an infectious period end earlier than the treatment start date; 2 weeks after the RXDATE will be used instead:\r\n", 
          file = log, append = T)
      cat(paste(names(dates), collapse ="\t"), file = log, append = T)
      for(z in r) {
        cat("\r\n", dates[z,1], file = log, append = T)
        for(c in 2:ncol(dates)) {
          cat(paste0("\t", as.character(dates[z,c])), file = log, append = T)
        }
      }
      cat("\r\n\r\n", file = log, append = T)
      dates$IPEnd[r] = as.Date(NA)
    }
    dates$IPStart[is.na(dates$IPStart)] = dates$earliestDate[is.na(dates$IPStart)] %m-% months(3)
    for(r in 1:nrow(dates)) { #using sapply will return the numeric
      dates$IPEnd[r] = calculateIPEnd(dates$STCASENO[r], gimsCases, dates$IPEnd[r])
    }
  } else {
    dates$IPStart = dates$earliestDate %m-% months(3)
    dates$IPEnd = as.Date(NA)
    for(r in 1:nrow(dates)) { #using sapply will return the numeric
      dates$IPEnd[r] = calculateIPEnd(dates$STCASENO[r], gimsCases, dates$IPEnd[r])
    }
  }
  
  ##calculate infection acquisition start
  dates$IAS = dates$RPTDATE %m-% months(dates$PEDAGE)
  if(!all(is.na(iae))) { #incorporate input data if available
    if(any(!is.na(iae$IAS))) {
      temp = iae[!is.na(iae$IAS) & iae$STCASENO %in% cases,]
      for(r in 1:nrow(temp)) {
        ped = littCaseData$Pediatric[littCaseData$STCASENO==temp$STCASENO[r]]
        if(!is.na(ped) & ped == "Y") {
          if(temp$IAS[r] >= dates$IAS[dates$STCASENO==temp$STCASENO[r]] %m-% months(1)) { #allow one month buffer given uncertainty in our date
            dates$IAS[dates$STCASENO==temp$STCASENO[r]] = temp$IAS[r]
          } else {
            cat(paste0(as.character(temp$STCASENO[r]), " has an input infection acquisition start of ", format(temp$IAS[r], format="%m/%d/%Y"), 
                # " but given PEDAGE and RPTDATE for this patient, ", format(dates$IAS[dates$STCASENO==temp$STCASENO[r]] %m-% months(1), format="%m/%d/%Y"),
                # " is the earliest this date can be, so ", format(dates$IAS[dates$STCASENO==temp$STCASENO[r]], format="%m/%d/%Y"), "will be used.\r\n", 
                " but this is before RPTDATE-PEDAGE for this patient, so that date (", format(dates$IAS[dates$STCASENO==temp$STCASENO[r]], format="%m/%d/%Y"), ") will be used instead.\r\n"), 
                file = log, append = T)
          }
        } else {
        #   cat(as.character(temp$STCASENO[r]), " has an input infection acquisition start, but is not pediatric; this date will be ignored.\r\n", 
        #       file = log, append = T)
          dates$IAS[dates$STCASENO==temp$STCASENO[r]] = temp$IAS[r]
        }
      }
    }
  }
  
  littCaseData = merge(littCaseData, dates[,names(dates) %in% c("STCASENO", "IPStart", "IPEnd", "IAS")], by="STCASENO", all = T)
  # if(any(!is.na(infectiousPeriod))) {
  #   infectiousPeriod$IPStart = convertToDate(infectiousPeriod$IPStart)
  # }
  littCaseData$UserDateData = sapply(littCaseData$STCASENO, caseInInputs, sxOnset, infectiousPeriod)
  if("Presumed.Source" %in% names(caseData)) {
    if("Presumed.Source.Strength" %in% names(caseData)) {
      littCaseData = merge(littCaseData, caseData[, c("STCASENO", "Presumed.Source", "Presumed.Source.Strength")], all.x = T)
    } else {
      littCaseData = merge(littCaseData, caseData[, c("STCASENO", "Presumed.Source")], all.x = T)
    }
  }
  
  if(all(class(progress)!="logical")) {
    progress$set(value = 0)
  }
  
  ####add epi in TB GIMS
  epi = cleanEpi(epi, cases, log)
  gimsCases$LKCASE1YR = replaceMissing(gimsCases$LKCASE1YR)
  gimsCases$LKCASE1ST = replaceMissing(gimsCases$LKCASE1ST)
  gimsCases$LKCASE1NO = replaceMissing(gimsCases$LKCASE1NO)
  gimsCases$LKCASE2YR = replaceMissing(gimsCases$LKCASE2YR)
  gimsCases$LKCASE2ST = replaceMissing(gimsCases$LKCASE2ST)
  gimsCases$LKCASE2NO = replaceMissing(gimsCases$LKCASE2NO)
  for(c in cases) {
    row = gimsCases$STCASENO==c
    lk1 = paste(gimsCases$LKCASE1YR[row], gimsCases$LKCASE1ST[row], gimsCases$LKCASE1NO[row], sep="")
    lk2 = paste(gimsCases$LKCASE2YR[row], gimsCases$LKCASE2ST[row], gimsCases$LKCASE2NO[row], sep="")
    lk = c(lk1, lk2)
    reas = c(gimsCases$LKREAS1, gimsCases$LKREAS2)
    for(i in length(lk)) {
      l = lk[i]
      if(nchar(l) > 0) {
        ##check this link or its inverse is not already in epi
        if(nrow(epi) == 0 ||
           !(any((epi$case1==c & epi$case2==l) | (epi$case1==l & epi$case2==c)))) {
          epi = rbind(epi,
                      data.frame(case1=c,
                                 case2=l,
                                 strength="definite",
                                 label = ifelse(reas[i] == "1" || reas[i] == "3", "TB GIMS self", "TB GIMS")))
        }
      }
    }
  }
  epi$strength = as.character(tolower(epi$strength))
  
  ####set up GIMS risk factors
  ##expected variables:
  ## c("HOMELESS", "HIVSTAT", "CORRINST", "LONGTERM", "IDU", "NONIDU", "ALCOHOL", 
  ##   "OCCUHCW", "OCCUCORR",
  ##   "RISKTNF", "RISKORGAN", "RISKDIAB", "RISKRENAL", "RISKIMMUNO"); add AnyCorr
  if(any(!is.na(gimsRiskFactor))) {
    names(gimsRiskFactor) = tolower(names(gimsRiskFactor))
    names(gimsRiskFactor) = removeWhitespacePeriods(names(gimsRiskFactor))
    if(!any(names(gimsRiskFactor)=="variable")) {
      cat("To use GIMS variables as risk factor, input table must have a column labeled \"variable\" that contains a list of the GIMS variables to use. Input table does not contain this column so GIMS variables will not be used.", 
          file = log, append = T)
      warning("To use GIMS variables as risk factor, input table must have a column labeled \"variable\" that contains a list of the GIMS variables to use. Input table does not contain this column so GIMS variables will not be used.")
    } else {
      gimsRiskFactor$variable = removeWhitespacePeriods(as.character(gimsRiskFactor$variable))
      if("AnyCorr" %in% gimsRiskFactor$variable) {
        gimsCases$AnyCorr = ""
      }
      rf = gimsCases[,names(gimsCases) %in% c("STCASENO", gimsRiskFactor$variable)]
      rf[] = lapply(rf, as.character) #convert to character
      if(any(names(gimsRiskFactor)=="weight")) {
        gimsRiskFactor = gimsRiskFactor[match(names(rf)[names(rf)!="STCASENO"], gimsRiskFactor$variable),]
        if(all(is.na(gimsRiskFactor$weight))) {
          gimsRiskFactor$weight = 1
        }
      } else {
        gimsRiskFactor$weight = 1
      }
      rf = rbind(rf, c("weight", gimsRiskFactor$weight))
      rf = rf[,c(T, !is.na(gimsRiskFactor$weight))] 
      gimsRiskFactor = gimsRiskFactor[!is.na(gimsRiskFactor$weight),]
      if(any(gimsRiskFactor$weight == 0)) {
        if(all(is.na(printVars))) {
          printVars = rf[,c(T, gimsRiskFactor$weight==0)]
        } else {
          printVars = merge(printVars, rf[,c(T, gimsRiskFactor$weight==0)], by="STCASENO")
        }
      }
      rf = rf[,c(T, gimsRiskFactor$weight[!is.na(gimsRiskFactor$weight)] > 0)] #only include positive values; 0 is for vis only and -1 will not be in any output
      
      if(any(gimsRiskFactor$weight > 0)) {
        ##for HIVSTAT, convert POS to Y
        if("HIVSTAT" %in% names(rf)) {
          rf$HIVSTAT[!is.na(rf$HIVSTAT) & rf$HIVSTAT=="POS"] = "Y"
        }
        
        ##for the occupations, need to also check the primaryocc column
        if("OCCUHCW" %in% names(rf)) {
          rf$OCCUHCW[rf$STCASENO!="weight"] = sapply(rf$STCASENO[rf$STCASENO!="weight"], function(sid) {
            return(ifelse((!is.na(gimsCases$OCCUHCW[gimsCases$STCASENO==sid]) & gimsCases$OCCUHCW[gimsCases$STCASENO==sid]=="Y") | 
                            (!is.na(gimsCases$PRIMARYOCC[gimsCases$STCASENO==sid]) & gimsCases$PRIMARYOCC[gimsCases$STCASENO==sid]=="HCW"),
                          "Y", "N"))})
        }
        if("OCCUCORR" %in% names(rf)) {
          rf$OCCUCORR[rf$STCASENO!="weight"] = sapply(rf$STCASENO[rf$STCASENO!="weight"], function(sid) {
            return(ifelse((!is.na(gimsCases$OCCUCORR[gimsCases$STCASENO==sid]) & gimsCases$OCCUCORR[gimsCases$STCASENO==sid]=="Y") | 
                            (!is.na(gimsCases$PRIMARYOCC[gimsCases$STCASENO==sid]) & gimsCases$PRIMARYOCC[gimsCases$STCASENO==sid]=="CORR"),
                          "Y", "N"))})
        }
        
        ##for AnyCorr, combine OCCURCORR and CORRINST
        if("AnyCorr" %in% names(rf)) {
          rf$AnyCorr[rf$STCASENO!="weight"] = sapply(rf$AnyCorr[rf$STCASENO!="weight"], function(sid){
            ifelse((!is.na(gimsCases$OCCUCORR[gimsCases$STCASENO==sid]) & gimsCases$OCCUCORR[gimsCases$STCASENO==sid]=="Y") | 
                     (!is.na(gimsCases$PRIMARYOCC[gimsCases$STCASENO==sid]) & gimsCases$PRIMARYOCC[gimsCases$STCASENO==sid]=="CORR") |
                     (!is.na(gimsCases$CORRINST[gimsCases$STCASENO==sid]) & gimsCases$CORRINST[gimsCases$STCASENO==sid]=="Y"),
                   "Y", "N")
          })
        }
        
        ## merge with addlRiskFactor
        if(any(!is.na(addlRiskFactor))) {
          addlRiskFactor = merge(addlRiskFactor, rf, by="STCASENO", all=T)
        } else {
          addlRiskFactor = rf
        }
      }
    }
  }
  
  if(all(class(progress)!="logical")) {
    progress$set(value = 1)
  }
  
  
  ####run litt
  names(littCaseData)[names(littCaseData) == "STCASENO"] = "ID"
  names(addlRiskFactor)[names(addlRiskFactor) == "STCASENO"] = "ID"
  littResults = litt(caseData = littCaseData, epi = epi, dist = dist, SNPcutoff = SNPcutoff, addlRiskFactor = addlRiskFactor,
                     progress = progress, log = log)
  littCaseData = littResults$caseData
  names(littCaseData)[names(littCaseData) == "ID"] = "STCASENO"
  names(addlRiskFactor)[names(addlRiskFactor) == "ID"] = "STCASENO"
  
  ####write results
  ##if Excel spreadsheet already exists, delete file; otherwise will note generate file, and if old table is bigger, will get extra rows from old table
  outputExcelFiles = ""
  outputExcelFiles = paste(outPrefix, c(epiFileName, dateFileName, caseFileName, txFileName, psFileName, rfFileName,
                                        distFileName, heatmapFileName), sep="")
  if(any(file.exists(outputExcelFiles))) {
    outputExcelFiles = outputExcelFiles[file.exists(outputExcelFiles)]
    file.remove(outputExcelFiles)
  }

  ####write results
  ###write out the dates
  write = dates[dates$STCASENO %in% littCaseData$STCASENO,]
  ##merge inputs
  if(all(is.na(iae))) {
    iae = data.frame(STCASENO = write$STCASENO,
                     inputIAS = as.Date(NA),
                     inputIAE = as.Date(NA))
  } else {
    iae = data.frame(STCASENO = iae$STCASENO,
                     inputIAS = convertToDate(iae$IAS),
                     inputIAE = convertToDate(iae$IAE))
  }
  write = merge(iae, write, by = "STCASENO", all.y=T)
  if(!all(is.na(inputIP))) {
    loclIP = data.frame(STCASENO = inputIP$STCASENO,
                        inputIPStart = convertToDate(inputIP$IPStart),
                        inputIPEnd = convertToDate(inputIP$IPEnd))
  } else {
    loclIP = data.frame(STCASENO = write$STCASENO,
                        inputIPStart = as.Date(NA),
                        inputIPEnd = as.Date(NA))
  }
  write = merge(loclIP, write, by = "STCASENO", all.y=T)
  if(!all(is.na(sxOnset))) {
    so = data.frame(STCASENO = inputSx$STCASENO,
                    inputSxOnset = inputSx$sxOnset)
  } else {
    so = data.frame(STCASENO = write$STCASENO,
                    inputSxOnset = as.Date(NA))
  }
  write = merge(so, write, by="STCASENO", all.y=T)
  ##split IAE and IP
  temp = splitPedEPDates(merge(write, littCaseData[,c("STCASENO", "ExtrapulmonaryOnly", "Pediatric")], by = "STCASENO", all=T))
  write$IAE = as.Date(NA)
  for(r in 1:nrow(write)) { #sapply won't return a date, and not always in the same order
    write$IPStart[r] = temp$IPStart[temp$STCASENO==write$STCASENO[r]]
    write$IPEnd[r] = temp$IPEnd[temp$STCASENO==write$STCASENO[r]]
    write$IAE[r] = temp$IAE[temp$STCASENO==write$STCASENO[r]]
  }
  ##write dates separately
  if(writeDate) {
    datewrite = write
    for(c in 2:ncol(datewrite)) {
      if(names(datewrite)[c] != "PEDAGE") {
        datewrite[,c] = format(datewrite[,c], "%m/%d/%Y")
      }
    }
    writeExcelTable(fileName=paste(outPrefix, dateFileName, sep=""),
                    sheetName = "dates",
                    df = datewrite,
                    stcasenolab = T,
                    wrapHeader = T)
  } else {
    outputExcelFiles = outputExcelFiles[!grepl(dateFileName, outputExcelFiles)]
  }
  if(all(class(progress)!="logical")) {
    progress$set(value = 7)
  }

  ###write out other case data variables
  ##remove all dates but earliest and IP
  # write = write[,names(write) %in% c("STCASENO", "earliestDate")] 
  ##add variables of interest (included in all output, but not in risk factors)
  gimsVars = data.frame(STCASENO = gimsCases$STCASENO,
                        accessionNumber = sapply(gimsCases$STCASENO, getAccessionNumber),
                        zipcode = gimsCases$ZIPCODE,
                        county = gimsCases$COUNTY,
                        state = gimsCases$STATE,
                        gender = gimsCases$SEX,
                        RACEHISP = gimsCases$RACEHISP,
                        age = gimsCases$AGE3)
  # write = merge(gimsVars, write, by="STCASENO")
  # write = merge(write, littCaseData, by = "STCASENO", all.y=T)
  littCaseData$sequenceAvailable = littCaseData$STCASENO %in% colnames(dist)
  write = merge(littCaseData, gimsVars, by = "STCASENO", all.x=T)
  write$STCASENO = as.character(write$STCASENO)
  write = write[,c(which(names(write)=="STCASENO"), which(names(write)=="accessionNumber"),
                   which(names(write)=="IPStart"),
                   which(names(write)=="IPEnd"),
                   which(!names(write) %in% c("STCASENO", "accessionNumber", "IPStart", "IPEnd")))]#do dates first
  # write$sequenceAvailable = write$STCASENO %in% colnames(dist)


  # if(any(!is.na(gimsRiskFactor))) {
  #   if(any(names(gimsRiskFactor)=="weight")) {
  #     gimsRiskFactor = gimsRiskFactor[gimsRiskFactor$weight == 0,] #negative numbers indicate do not want variable in output, postive values will be in addlRiskFactor
  #   }
  #   if(nrow(gimsRiskFactor)) {
  #     write = merge(write, gimsCases[,names(gimsCases) %in% c("STCASENO", as.character(gimsRiskFactor$variable))], by="STCASENO", all.x=T)
  #   }
  # }

  if(any(!is.na(littResults$rfWeights))) {
    # write[] = lapply(write, as.character) #convert to character
    # write = rbind(write,
    #               c("weight", sapply(names(write)[-1], function(v) {
    #                 ifelse(v %in% littResults$rfWeights$variable,
    #                        littResults$rfWeights$weight[littResults$rfWeights$variable==v], "")
    #               })))
    if(any(!is.na(gimsRiskFactor))) { #clean gims risk factor names so it will be consistent with case data table
      # gRFs = littResults$rfWeights$variable %in% gimsRiskFactor$variable
      # cleanNames = cleanGimsRiskFactorNames(littResults$rfWeights$variable[gRFs])
      # littResults$rfWeights$variable[gRFs] = cleanNames
      littResults$rfWeights$variable = cleanGimsRiskFactorNames(littResults$rfWeights$variable)
    }
    littResults$rfWeights$variable = gsub(".", " ", littResults$rfWeights$variable, fixed=T)
    writeExcelTable(fileName=paste(outPrefix, rfFileName, sep=""),
                    df = littResults$rfWeights,
                    filter=F)
    if(all(class(progress)!="logical")) {
      progress$set(value = 8)
    }
  } else {
    outputExcelFiles = outputExcelFiles[!grepl(rfFileName, outputExcelFiles)]
  }

  if(nrow(littResults$epi)) {
    write$numEpiLinks = sapply(write$STCASENO,
                               function(sno) {
                                 return(sum(littResults$epi$case1==sno | littResults$epi$case2==sno))
                               })
  } else {
    write$numEpiLinks = 0
  }
  if(nrow(littResults$topRanked)) {
    write$numTimesIsRank1 = sapply(write$STCASENO,
                                   function(sno) {
                                     return(sum(littResults$topRanked$source==sno))
                                   })
    write$numPotSourcesRanked1 = sapply(write$STCASENO,
                                        function(sno) {
                                          return(sum(littResults$topRanked$target==sno))
                                        })
  } else {
    write$numTimesIsRank1 = 0
    write$numPotSourcesRanked1 = 0
  }
  if(nrow(littResults$allPotentialSources)) {
    write$totNumPotSources = sapply(write$STCASENO,
                                    function(sno) {
                                      return(sum(littResults$allPotentialSources$target==sno))
                                    })
  } else {
    write$totNumPotSources = 0
  }
  # if(any(write$STCASENO=="weight")) {
  #   write[write$STCASENO=="weight", names(write) %in% c("numEpiLinks", "numTimesIsRank1", "numPotSourcesRanked1", "totNumPotSources", "sequenceAvailable")] = NA #do not give numbers for weight
  # }

  ###add risk factors and weights
  if(any(!is.na(addlRiskFactor))) {
    addlRiskFactor = addlRiskFactor[addlRiskFactor$STCASENO != "weight",]
    write = merge(write, addlRiskFactor, by="STCASENO", all.x=T)
  }
  if(any(!is.na(printVars))) {
    write = merge(write, printVars, by="STCASENO", all.x=T)
  }

  ###write case data
  write = write[order(as.character(write$STCASENO)),] #without as.character, order will be weird because factors are out of order
  cleanCaseOutput(caseOut=splitPedEPDates(write), outPrefix=outPrefix)
  if(all(class(progress)!="logical")) {
    progress$set(value = 9) #skip 8 unless have risk factors
  }

  ###write epi links
  w = writeEpiTable(littResults, outPrefix, stcasenolab = T, log = log)
  if(!w) { #table not written, so do not include in output list
    outputExcelFiles = outputExcelFiles[!grepl(epiFileName, outputExcelFiles)]
  }
  if(all(class(progress)!="logical")) {
    progress$set(value = 10)
  }

  ###write distance matrix if writeDist is true
  if(!all(is.na(dist)) & writeDist) {
    writeDistTable(dist, outPrefix)
  } else {
    outputExcelFiles = outputExcelFiles[!grepl(distFileName, outputExcelFiles)]
  }

  ###clean up and output transmission network
  w = writeTopRankedTransmissionTable(littResults, outPrefix, stcasenolab = T, log = log)
  if(!w) { #table not written, so do not include in output list
    outputExcelFiles = outputExcelFiles[!grepl(txFileName, outputExcelFiles)]
  }
  if(all(class(progress)!="logical")) {
    progress$set(value = 11)
  }

  ###get categorical labels and combine source matrix and reason filtered into one Excel spreadsheet
  writeAllSourcesTable(littResults, outPrefix, stcasenolab = T)
  if(all(class(progress)!="logical")) {
    progress$set(value = 12)
  }

  ###generate heatmap if any cases are ranked
  if(nrow(littResults$topRanked)) {
    littHeatmap(outPrefix = outPrefix, all = littResults$allPotentialSources)
  } else {
    outputExcelFiles = outputExcelFiles[!grepl(heatmapFileName, outputExcelFiles)]
  }
  if(all(class(progress)!="logical")) {
    progress$set(value = 13)
  }

  write.table(littResults$summary, paste(outPrefix, "PotentialSources.txt", sep=""),
              row.names = F, col.names = T, quote = F, sep = "\t") ##for consistency in other validation work
  littResults$outputFiles = c(log, outputExcelFiles)
  littResults$caseData = write
  return(littResults)
}
