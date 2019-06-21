##sets up LITT without TB GIMS data
source("litt.R")

##cleans up and writes case data table
##caseOut = case output data to format
##outPrefix = output prefix
##keepUserDate = if true, keep user date data column
cleanCaseOutput <- function(caseOut, outPrefix, keepUserDate = F) {
  ###convert T/F to Y/N
  if(!keepUserDate) {
    caseOut = caseOut[,names(caseOut) != "UserDateData"] #remove user date data column since all samples will be Y
  } else {
    caseOut$UserDateData = ifelse(caseOut$UserDateData, "Y", "N")
  }
  caseOut$sequenceAvailable = ifelse(caseOut$sequenceAvailable, "Y", "N")
  
  ##format date/zip columns
  caseOut$IPStart = format(caseOut$IPStart, format="%m/%d/%Y")
  caseOut$IPEnd = format(caseOut$IPEnd, format="%m/%d/%Y")
  caseOut$IAS = format(caseOut$IAS, format="%m/%d/%Y")
  caseOut$IAE = format(caseOut$IAE, format="%m/%d/%Y")
  
  writeExcelTable(fileName=paste(outPrefix, caseFileName, sep=""),
                  sheetName = "case data",
                  df = caseOut,
                  wrapHeader=T)
}

##function that takes in the outputs of the TB GIMS version and formats it for running the no GIMS version
##prefix = prefix used in the GIMS run
##rfs = column names of risk factors used in the GIMS run
formatLittGimsCaseTable <- function(prefix) {
  caseData = read.xlsx(paste(prefix, caseFileName, sep=""), sheetName = 1)
  return(cleanedCaseDataHeadersToVarNames(caseData))
}

##for the given data frame df, convert the cleaned header name to the needed variable name
##for using case data table from previous LITT runs
##rfs = column names of risk factors to keep
cleanedCaseDataHeadersToVarNames <- function(caseData) {
  ##fix names
  names(caseData)[names(caseData)=="Case.ID" | names(caseData)=="Case ID"] = "ID"
  if(!"ID" %in% names(caseData)) {
    names(caseData)[tolower(names(caseData))=="stcaseno"] = "ID"
    names(caseData)[names(caseData)=="State.Case.Number"] = "ID"
  }
  names(caseData)[names(caseData)=="Evidence.of.Cavity.by.X.Ray..XRAYCAV."] = "XRAYCAV"
  names(caseData)[names(caseData)=="Sputum.Smear..SPSMEAR."] = "SPSMEAR" #make.names(spName)
  names(caseData)[names(caseData)=="Extrapulmonary.Only"] = "ExtrapulmonaryOnly"
  names(caseData)[names(caseData)=="User.Input.Date.Data.Available"] = "UserDateData"
  ##remove previously calculated LITT summary
  caseData = caseData[,!names(caseData) %in%
                        c("Number.of.Epi.Links", "Number.of.Times.is.Ranked.1st.in.Potential.Source.List",
                          "Number.of.Potential.Sources.Tied.for.Rank.1", "Total.Number.of.Potential.Sources",
                          "Sequence.Available.In.Analysis")]
  return(caseData)
}

###sets up the inputs for LITT then calls LITT then formats the outputs
###inputs:
##outPrefix = prefix for writing results (default is none); if a different directory is desired, provide that here
##dist = WGS SNP distance matrix (row names and column names are state case number) (optional)
##caseData = combined dataframe with ID, infectious period start and end and additional risk factor data; required field
##epi = data frame with columns title case1, case2, strength (strength of the epi link between cases 1 and 2), label (label for link if known) -> optional, if not given, the epi links in GIMS will be used as definite epi links
##SNPcutoff = eliminate source if the SNP distance from the target is greater than this cutoff
##rfTable = dataframe with variable (list of additional risk factors, which correspond to column names in caseData) and weight
##keepExtraCDcol = if true, keep extra columns in case data, e.g. if case data is from a run of LITT with TB GIMS and the user wants to keep the extra surveillance columns (will remove number of epi links and number of times ranked 1st)
##writeDist = if true, write distance matrix to Excel and include in list of outputs
##if appendlog is true, append results to log file; otherwise overwrite (needed because may read distance matrix first)
##progress = progress bar for R Shiny interface (NA if not running through interface)
littNoGims <- function(outPrefix = "", caseData, dist=NA, epi=NA, SNPcutoff = snpDefaultCut, rfTable= NA, 
                       keepExtraCDcol = F, writeDist = F, appendlog = F, progress = NA) {
  log = paste(outPrefix, defaultLogName, sep="")
  cat("LITT analysis\r\nSNP cutoff = ", SNPcutoff, "\r\n", file = log, append=appendlog)
  
  ####check inputs
  caseData = caseData[!apply(caseData, 1, function(x) all(is.na(x))),
                      !apply(caseData, 2, function(x) all(is.na(x)))] #remove rows and columns that are all NA
  names(caseData) = paste0(toupper(substring(names(caseData),1,1)), substring(names(caseData), 2)) #make sure have correct capitalization (e.g. Pediatric, not pediatric)
  if(all(is.na(caseData))) {
    cat("A case data table is required\r\n", file = log, append = T)
    stop("A case data table is required")
  }
  if(!"ID" %in% names(caseData)) {
    caseData = cleanedCaseDataHeadersToVarNames(caseData)
  }
  if(!"ID" %in% names(caseData)) {
    cat("A case data table with one column labeled ID or STATECASNO is required.\r\n", file = log, append = T)
    stop("Case data table must have an ID column.")
  }
  if("weight" %in% caseData$ID) {
    cat("Case ID cannot be weight; this row has been removed from case data table.\r\n", file = log, append = T)
    caseData = caseData[caseData$ID!="weight",]
  }
  if(nrow(caseData) < 2) {
    cat("LITT requires at least two cases, but there ", ifelse(nrow(caseData)==1, "is ", "are "),
        nrow(caseData), ifelse(nrow(caseData)==1, "case", "cases"), "\r\n", file = log, append = T)
    stop("LITT requires at least two cases, but there are ", nrow(caseData))
  }
  
  caseData = fixPresumedSource(caseData)
  keepUserDate = "UserDateData" %in% names(caseData)
  
  ###check IP start present
  caseData = fixIPnames(caseData, log)
  if(!"IPStart" %in% names(caseData) || !"IPEnd" %in% names(caseData)) {
    cat("The case data table must contain columns called IPStart and IPEnd, with dates, which indicate infectious period start and end for each case.\r\n", 
        file = log, append = T)
    stop("Infectious period (columns named IPStart and IPEnd) is required in case data table.")
  }
  caseData$IPStart = convertToDate(caseData$IPStart)
  caseData$IPEnd = convertToDate(caseData$IPEnd)
  
  ###merge IP and infection acquisition window if needed
  caseData = fixIAE(caseData)
  caseData = mergeIAEtoIPstart(caseData, log)
  caseData = caseData[,names(caseData)!="IAE"] #remove IAE (will be added in later for printing)
  # if("IAS" %in% names(caseData)) {
  #   caseData$Pediatric[is.na(caseData$Pediatric)] = "N"
  #   if(any(!is.na(caseData$IAS) & caseData$Pediatric!="Y")) {
  #     bad = caseData$ID[!is.na(caseData$IAS) & caseData$Pediatric!="Y"]
  #     cat("The following cases have an infection acquisition start but are not pediatric: ",
  #         paste(bad, collapse=", "), ". Infection acquisition start will be ignored for these cases.\r\n",
  #         file = log, append = T)
  #     caseData$IAS[caseData$Pediatric!="Y"] = as.Date(NA) #IAS should be missing for IAS
  #   }
  # }
  
  ####split out additional risk factors
  # addlRiskFactor=NA 
  # if(!all(names(caseData) %in% c(expectedcolnames, optionalcolnames))) { #have RFs
  #   addlRiskFactor=caseData[,!names(caseData) %in% c(expectedcolnames[expectedcolnames!="ID"], optionalcolnames)]
  #   caseData=caseData[caseData$ID!="weight",names(caseData) %in% c(expectedcolnames, optionalcolnames)]
  # }
  addlRiskFactor=NA #risk factor data, with row for weight
  printVars = NA #additional variables to print but not standard variables or risk factors
  if(!all(is.na(rfTable))) {
    rfTable = fixRfTable(rfTable, log)
    if(!all(is.na(rfTable))) {
      rfTable$variable = paste0(toupper(substring(rfTable$variable,1,1)), substring(rfTable$variable, 2)) #capitalized the first letter in case data table
      if(any(!rfTable$variable %in% names(caseData))) { #extra variables not in case data table
        miss = !rfTable$variable %in% names(caseData)
        cat("The following variables are in the risk factor table but not in the case data table, so will not be used in analysis: ",
            paste(rfTable$variable[miss], collapse=", "), "\r\n", file = log, append = T)
        rfTable = rfTable[!miss,]
      }
      if(any(!names(caseData) %in% c(expectedcolnames, optionalcolnames, rfTable$variable)) & !keepExtraCDcol) {
        miss = !names(caseData) %in% c(expectedcolnames, optionalcolnames, rfTable$variable)
        cat("There are extra columns in the case data table, which will be removed: ",
            paste(names(caseData)[miss], collapse=", "), "\r\n", file = log, append = T)
      }
      ##set up print table (weight 0)
      rfTable = rfTable[!is.na(rfTable$weight),]
      if(any(rfTable$weight==0)) {
        # printVars = caseData[,names(caseData) %in% c("ID", rfTable$variable[rfTable$weight==0])]
        printVars = caseData[,c("ID", rfTable$variable[rfTable$weight==0])]
      }
      rfTable = rfTable[rfTable$weight > 0,]
      ##set up additional risk factor table
      if(nrow(rfTable) > 0) {
        # addlRiskFactor = caseData[,names(caseData) %in% c("ID", rfTable$variable)]
        addlRiskFactor = caseData[,c("ID", rfTable$variable)]
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
  ##if user wants to include extra surveillance columns from gims, merge with printVars
  if(keepExtraCDcol) {
    gims = caseData[,!names(caseData) %in% 
                      c(expectedcolnames[expectedcolnames!="ID"], optionalcolnames, 
                        "Calculated.Infectious.Period.Start", "Calculated.Infectious.Period.End")]
    if(!all(is.na(rfTable))) {
      gims = gims[,!names(gims) %in% rfTable$variable]
    }
    if(!class(gims)=="character") { #if there were not extra columns gims ends up being character vector of ID
      if(all(is.na(printVars))) {
        printVars = gims
      } else {
        gims = gims[,!names(gims) %in% names(printVars)[names(printVars)!="ID"]]
        printVars = merge(gims, printVars, by="ID")
      }
    }
  }
  caseData=caseData[,names(caseData) %in% c(expectedcolnames, optionalcolnames)]
  
  ###list of cases = cases in caseData
  cases = caseData$ID
  cat(paste("Number of cases:", length(cases), "\r\n"), file = log, append = T)
  
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
  caseData = littResults$caseData
  
  ####write results
  ###if Excel spreadsheet already exists, delete file; otherwise will note generate file, and if old table is bigger, will get extra rows from old table
  outputExcelFiles = paste(outPrefix, c(epiFileName, caseFileName, txFileName, psFileName, rfFileName, distFileName, heatmapFileName), sep="")
  if(any(file.exists(outputExcelFiles))) {
    outputExcelFiles = outputExcelFiles[file.exists(outputExcelFiles)]
    file.remove(outputExcelFiles)
  }
  
  ###write out other case data variables
  write = caseData
  write$sequenceAvailable = write$ID %in% colnames(dist)
  
  if(any(!is.na(littResults$rfWeights))) {
    # write[] = lapply(write, as.character) #convert to character
    # write = rbind(write,
    #               c("weight", sapply(names(write)[-1], function(v) {
    #                 ifelse(v %in% littResults$rfWeights$variable,
    #                        littResults$rfWeights$weight[littResults$rfWeights$variable==v], "")
    #               })))
    littResults$rfWeights$variable = gsub(".", " ", littResults$rfWeights$variable, fixed=T)
    writeExcelTable(fileName=paste(outPrefix, rfFileName, sep=""),
                    df = littResults$rfWeights,
                    filter=F)
    if(all(class(progress)!="logical")) {
      progress$set(value = 7) 
    }
  } else {
    outputExcelFiles = outputExcelFiles[!grepl(rfFileName, outputExcelFiles)]
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
  # write[write$ID=="weight", names(write) %in% c("numEpiLinks", "numTimesIsRank1", "sequenceAvailable")] = NA #do not give numbers for weight
  
  ###add risk factors and extra variables
  if(any(!is.na(addlRiskFactor))) {
    addlRiskFactor = addlRiskFactor[addlRiskFactor$ID != "weight",]
    write = merge(write, addlRiskFactor, by="ID", all.x=T)
  }
  if(any(!is.na(printVars))) {
    write = merge(write, printVars, by="ID", all.x=T)
  }
  
  ###write case table
  write = write[order(as.character(write$ID)),] #without as.character, order will be weird because factors are out of order
  cleanCaseOutput(caseOut=splitPedEPDates(write), outPrefix=outPrefix, keepUserDate = keepUserDate)
  
  if(all(class(progress)!="logical")) {
    progress$set(value = 8) #skip 7 unless have rfs
  }
  
  ###write distance matrix if writeDist is true
  if(!all(is.na(dist)) & writeDist) {
    writeDistTable(dist, outPrefix)
  } else {
    outputExcelFiles = outputExcelFiles[!grepl(distFileName, outputExcelFiles)] 
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
  
  ###generate heatmap if any cases are ranked
  if(nrow(littResults$topRanked)) {
    littHeatmap(outPrefix = outPrefix, all = littResults$allPotentialSources)
  } else {
    outputExcelFiles = outputExcelFiles[!grepl(heatmapFileName, outputExcelFiles)]
  }
  if(all(class(progress)!="logical")) {
    progress$set(value = 12)
  }
  
  littResults$outputFiles = c(log, outputExcelFiles)
  littResults$caseData = write
  return(littResults)
}
