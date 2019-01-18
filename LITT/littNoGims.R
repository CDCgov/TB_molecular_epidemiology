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
  
  writeExcelTable(fileName=paste(outPrefix, "LITT_Calculated_Case_Data.xlsx", sep=""),
                  sheetName = "case data",
                  df = caseOut,
                  wrapHeader=T)
}

##function that takes in the outputs of the TB GIMS version and formats it for running the no GIMS version
##prefix = prefix used in the GIMS run
##rfs = column names of risk factors used in the GIMS run
formatLittGimsCaseTable <- function(prefix, rfs=NA) {
  caseData = read.xlsx(paste(prefix, "LITT_Calculated_Case_Data.xlsx", sep=""), sheetName = 1)
  vars = c("State.Case.Number", 
           "Calculated.Infectious.Period.Start", "Calculated.Infectious.Period.End",
           "Evidence.of.Cavity.by.X.Ray", "Sputum.Smear",
           "Extrapulmonary.Only", "Pediatric", "User.Input.Date.Data.Available")
  if(!any(is.na(rfs))) {
    vars = c(vars, rfs)
  }
  caseData = caseData[,names(caseData) %in% vars]
  names(caseData)[names(caseData)=="State.Case.Number"] = "ID"
  names(caseData)[names(caseData)=="Evidence.of.Cavity.by.X.Ray"] = "XRAYCAV"
  names(caseData)[names(caseData)=="Sputum.Smear"] = "SPSMEAR"
  names(caseData)[names(caseData)=="Extrapulmonary.Only"] = "ExtrapulmonaryOnly"
  names(caseData)[names(caseData)=="User.Input.Date.Data.Available"] = "UserDateData"
  return(caseData)
}

###sets up the inputs for LITT then calls LITT then formats the outputs
###inputs:
##outPrefix = prefix for writing results (default is none); if a different directory is desired, provide that here
##dist = WGS SNP distance matrix (row names and column names are state case number) (optional)
##caseData = combined dataframe with ID, infectious period start and end and additional risk factor data; required field
##epi = data frame with columns title case1, case2, strength (strength of the epi link between cases 1 and 2), label (label for link if known) -> optional, if not given, the epi links in GIMS will be used as definite epi links
##SNPcutoff = eliminate source if the SNP distance from the target is greater than this cutoff
littNoGims <- function(outPrefix = "", caseData, dist=NA, epi=NA, SNPcutoff = snpDefaultCut) {
  ####check inputs
  if(all(is.na(caseData)) || !"ID" %in% names(caseData)) {
    if("stcaseno" %in% tolower(names(caseData))) {
      names(caseData)[tolower(names(caseData))=="stcaseno"] = "ID"
    } else {
      stop("Case data is required, and must have a column named ID")
    }
  }
  if(nrow(caseData) < 2) {
    stop("LITT requires at least two cases, but there are ", nrow(caseData))
  }
  
  ###check IP start present
  caseData = fixIPnames(caseData)
  if(!"IPStart" %in% names(caseData) || !"IPEnd" %in% names(caseData)) {
    stop("Infectious period (IPStart and IPEnd) is required")
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
  cat(paste("Number of cases:", length(cases), "\n"))
  
  ####set up epi
  epi = fixEpiNames(epi)
  epi = cleanEpi(epi, cases)
  epi$strength = as.character(tolower(epi$strength))
  # if(nrow(epi) > 0) {
  #   epi$SNPdistance = getSNPDistance(epi, dist) 
  # }
  # cat(paste("Number of epi links:", nrow(epi)), "\n")
  
  ##clean up distance matrix
  colnames(dist) = removeXFromNames(colnames(dist))
  
  ####run litt
  littResults = litt(caseData = caseData, epi = epi, dist = dist, SNPcutoff = SNPcutoff, addlRiskFactor = addlRiskFactor)
  
  ####write results
  ###if Excel spreadsheet already exists, delete file; otherwise will note generate file, and if old table is bigger, will get extra rows from old table
  outputExcelFiles = paste(outPrefix, c("LITT_Calculated_Epi_Data.xlsx", 
                                        "LITT_Calculated_Case_Data.xlsx", "LITT_Transmission_Network.xlsx",
                                        "LITT_All_Potential_Sources.xlsx"), sep="")
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
  write[write$ID=="weight", names(write) %in% c("numEpiLinks", "numTimesIsRank1")] = NA #do not give numbers for weight
  cleanCaseOutput(caseOut=write, outPrefix=outPrefix)
  
  ###write epi links
  # if(nrow(littResults$epi)) {
  #   epiOut = littResults$epi
  #   writeExcelTable(fileName=paste(outPrefix, "LITT_Calculated_Epi_Data.xlsx", sep=""),
  #                   sheetName="Epi links",
  #                   df = epiOut)
  # } else {
  #   cat("No epi links\n")
  # }
  writeEpiTable(littResults, outPrefix)
  
  ####clean up and output transmission network
  # if(nrow(littResults$topRanked)) {
  #   writeExcelTable(fileName=paste(outPrefix, "LITT_Top_Ranked_Transmission_Network.xlsx", sep=""),
  #                   sheetName = "top ranked sources", 
  #                   df = littResults$topRanked)
  # } else {
  #   cat("There were no cases with a potential source that passed all filters.\n")
  # }
  writeTopRankedTransmissionTable(littResults, outPrefix)
  
  ####get categorical labels and combine source matrix and reason filtered into one Excel spreadsheet
  # allSources = littResults$allPotentialSources
  # cat = getScoreCategories(allSources)
  # allSources[,c(3:6, 8:9)] = apply(allSources[,c(3:6, 8:9)], 2, as.numeric) #make columns numeric for Excel
  # ##move label to end
  # allSources = cbind(allSources[,names(allSources)!="label"], data.frame(label=allSources[,names(allSources)=="label"]))
  # cat = cbind(cat[,names(cat)!="label"], data.frame(label=cat[,names(cat)=="label"]))
  # wb = writeExcelTable(fileName=paste(outPrefix, "LITT_All_Potential_Sources.xlsx", sep=""),
  #                      sheetName = "numeric potential sources", 
  #                      df = allSources, 
  #                      wrapHeader=T)
  # wb = writeExcelTable(workbook=wb,
  #                      sheetName = "categorical potential sources", 
  #                      df = cat, 
  #                      wrapHeader=T)
  writeAllSourcesTable(littResults, outPrefix, stcasenolab = F)
}
