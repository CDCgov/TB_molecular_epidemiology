##sets up LITT without TB GIMS data
source("litt.R")

##cleans up and writes case data table
##caseOut = case output data to format
##outPrefix = output prefix
cleanCaseOutput <- function(caseOut, outPrefix) {
  ###convert T/F to Y/N
  caseOut = caseOut[,names(caseOut) != "UserDateData"] #remove user date data column since all samples will be Y
  caseOut$sequenceAvailable = ifelse(caseOut$sequenceAvailable, "Y", "N")
  
  ##format date/zip columns
  caseOut$IPStart = format(caseOut$IPStart, format="%m/%d/%Y")
  caseOut$IPEnd = format(caseOut$IPEnd, format="%m/%d/%Y")
  
  names(caseOut) = gsub(".", "", names(caseOut), fixed=T) #clean up risk factors
  
  writeExcelTable(fileName=paste(outPrefix, caseFileName, sep=""),
                  sheetName = "case data",
                  df = caseOut,
                  wrapHeader=T)
}

##function that takes in the outputs of the TB GIMS version and formats it for running the no GIMS version
##prefix = prefix used in the GIMS run
##rfs = column names of risk factors used in the GIMS run
formatLittGimsCaseTable <- function(prefix, rfs=NA) {
  caseData = read.xlsx(paste(prefix, caseFileName, sep=""), sheetName = 1)
  return(cleanedCaseDataHeadersToVarNames(caseData, rfs=rfs))
}

##for the given data frame df, convert the cleaned header name to the needed variable name
##for using case data table from previous LITT runs
##rfs = column names of risk factors to keep
cleanedCaseDataHeadersToVarNames <- function(caseData, rfs = NA) {
  ##fix names
  names(caseData)[tolower(names(caseData))=="stcaseno"] = "ID"
  names(caseData)[names(caseData)=="State.Case.Number"] = "ID"
  names(caseData)[names(caseData)=="Case.ID"] = "ID"
  names(caseData)[names(caseData)=="Evidence.of.Cavity.by.X.Ray"] = "XRAYCAV"
  names(caseData)[names(caseData)=="Sputum.Smear"] = "SPSMEAR"
  names(caseData)[names(caseData)=="Extrapulmonary.Only"] = "ExtrapulmonaryOnly"
  names(caseData)[names(caseData)=="User.Input.Date.Data.Available"] = "UserDateData"
  ##subset to needed variables
  vars = c(expectedcolnames, "Calculated.Infectious.Period.Start", "Calculated.Infectious.Period.End")
  if(!any(is.na(rfs))) {
    vars = c(vars, rfs)
  }
  caseData = caseData[,names(caseData) %in% vars]
  return(caseData)
}

###sets up the inputs for LITT then calls LITT then formats the outputs
###inputs:
##outPrefix = prefix for writing results (default is none); if a different directory is desired, provide that here
##dist = WGS SNP distance matrix (row names and column names are state case number) (optional)
##caseData = combined dataframe with ID, infectious period start and end and additional risk factor data; required field
##epi = data frame with columns title case1, case2, strength (strength of the epi link between cases 1 and 2), label (label for link if known) -> optional, if not given, the epi links in GIMS will be used as definite epi links
##SNPcutoff = eliminate source if the SNP distance from the target is greater than this cutoff
##progress = progress bar for R Shiny interface (NA if not running through interface)
littNoGims <- function(outPrefix = "", caseData, dist=NA, epi=NA, SNPcutoff = snpDefaultCut, progress = NA) {
  log = paste(outPrefix, defaultLogName, sep="")
  cat("LITT analysis\n", file = log)
  
  ####check inputs
  if(all(is.na(caseData))) {
    cat("A case data table is required\n", file = log, append = T)
    stop("A case data table is required")
  }
  if(!"ID" %in% names(caseData)) {
    caseData = cleanedCaseDataHeadersToVarNames(caseData)
  }
  if(!"ID" %in% names(caseData)) {
    cat("A case data table with one column labeled ID or STATECASNO is required\n", file = log, append = T)
    stop("Case data table must have a column with the case ID, named ID")
  }
  if(nrow(caseData) < 2) {
    cat("LITT requires at least two cases, but there ", ifelse(nrow(caseData)==1, "is ", "are "),
        nrow(caseData), ifelse(nrow(caseData)==1, "case", "cases"), "\n", file = log, append = T)
    stop("LITT requires at least two cases, but there are ", nrow(caseData))
  }
  
  ###check IP start present
  caseData = fixIPnames(caseData, log)
  if(!"IPStart" %in% names(caseData) || !"IPEnd" %in% names(caseData)) {
    cat("The case data table must contain columns called IPStart and IPEnd, which indicate infectious period start and end for each case\n", 
        file = log, append = T)
    stop("Infectious period (columns named IPStart and IPEnd) is required in case data table")
  }
  caseData$IPStart = convertToDate(caseData$IPStart)
  caseData$IPEnd = convertToDate(caseData$IPEnd)

  ####split out additional risk factors
  addlRiskFactor=NA 
  if(!all(names(caseData) %in% c(expectedcolnames, optionalcolnames))) { #have RFs
    addlRiskFactor=caseData[,!names(caseData) %in% c(expectedcolnames[expectedcolnames!="ID"], optionalcolnames)]
    caseData=caseData[caseData$ID!="weight",names(caseData) %in% c(expectedcolnames, optionalcolnames)]
  }
  
  ###list of cases = cases in caseData
  cases = caseData$ID
  cat(paste("Number of cases:", length(cases), "\n"), file = log, append = T)
  
  ####set up epi
  if(!all(is.na(epi))) {
    epi = fixEpiNames(epi, log)
    epi = cleanEpi(epi, cases, log)
    epi$strength = as.character(tolower(epi$strength))
  }
  
  ##clean up distance matrix
  colnames(dist) = removeXFromNames(colnames(dist))
  
  if(all(class(progress)!="logical")) {
    progress$set(value = 1)
  }
  
  ####run litt
  littResults = litt(caseData = caseData, epi = epi, dist = dist, SNPcutoff = SNPcutoff, addlRiskFactor = addlRiskFactor,
                     progress = progress, log = log)
  
  ####write results
  ###if Excel spreadsheet already exists, delete file; otherwise will note generate file, and if old table is bigger, will get extra rows from old table
  outputExcelFiles = paste(outPrefix, c(epiFileName, caseFileName, txFileName, psFileName), sep="")
  if(any(file.exists(outputExcelFiles))) {
    outputExcelFiles = outputExcelFiles[file.exists(outputExcelFiles)]
    file.remove(outputExcelFiles)
  }
  
  ###write out other case data variables
  write = caseData
  write$sequenceAvailable = write$ID %in% colnames(dist)
  
  ##add risk factors and weights
  if(any(!is.na(addlRiskFactor))) {
    addlRiskFactor = addlRiskFactor[addlRiskFactor$ID!="weight",] #this weight is pre-calculation; correct weight added below
    write = merge(write, addlRiskFactor, by="ID", all = T)
  }
  
  write = write[order(as.character(write$ID)),] #without as.character, order will be weird because factors are out of order
  
  if(any(!is.na(littResults$rfWeights))) {
    write[] = lapply(write, as.character) #convert to character
    write = rbind(write,
                  c("weight", sapply(names(write)[-1], function(v) {
                    ifelse(v %in% littResults$rfWeights$variable,
                           littResults$rfWeights$weight[littResults$rfWeights$variable==v], "")
                  })))
    # writeExcelTable(fileName=paste(outPrefix, "LITT_Risk_Factor_Weights.xlsx", sep=""),
    #                 df = littResults$rfWeights)
  }
  
  if(nrow(littResults$epi)) {
    write$numEpiLinks = sapply(write$ID,
                               function(sno) {
                                 return(sum(littResults$epi$case1==sno | littResults$epi$case2==sno))
                               })
  } else {
    write$numEpiLinks = 0
  }
  if(nrow(littResults$topRanked)) {
    write$numTimesIsRank1 = sapply(write$ID, 
                                   function(sno) {
                                     return(sum(littResults$topRanked$source==sno))
                                   })
  } else {
    write$numTimesIsRank1 = 0
  }
  write[write$ID=="weight", names(write) %in% c("numEpiLinks", "numTimesIsRank1", "sequenceAvailable")] = NA #do not give numbers for weight
  cleanCaseOutput(caseOut=write, outPrefix=outPrefix)
  if(all(class(progress)!="logical")) {
    progress$set(value = 8) #skip 7 unless have rfs
  }
  
  ###write epi links
  w = writeEpiTable(littResults, outPrefix, log = log)
  if(!w) { #table not written, so do not include in output list
    outputExcelFiles = outputExcelFiles[!grepl(epiFileName, outputExcelFiles)]
  }
  if(all(class(progress)!="logical")) {
    progress$set(value = 9)
  }
  
  ###clean up and output transmission network
  w = writeTopRankedTransmissionTable(littResults, outPrefix, log = log)
  if(!w) { #table not written, so do not include in output list
    outputExcelFiles = outputExcelFiles[!grepl(txFileName, outputExcelFiles)]
  }
  if(all(class(progress)!="logical")) {
    progress$set(value = 10)
  }
  
  ###get categorical labels and combine source matrix and reason filtered into one Excel spreadsheet
  writeAllSourcesTable(littResults, outPrefix, stcasenolab = F)
  if(all(class(progress)!="logical")) {
    progress$set(value = 11)
  }
  
  littResults$outputFiles = c(log, outputExcelFiles)
  return(littResults)
}
