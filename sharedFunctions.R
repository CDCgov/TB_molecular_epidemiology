####generic methods that can be used by multiple algorithms, including both LITT and LATTE

##clean dateList by removing invalid dates, e.g. 1/1/1900 is sometimes used for missing data
cleanDate <- function(dateList) {
  dateList[dateList==as.Date("1900-1-1")] = as.Date(NA)
  return(dateList)
}

##function that converts the given dateList to a vector of dates, allowing different date format options
convertToDate <- function(dateList) {
  dateList = as.character(dateList)
  dateList[grepl("[A-Z]|[a-z]", dateList)] = NA #set anything with letters to be NA (this should include "NA")
  ##figure out pattern (need to skip NA)
  pattern = dateList[1]
  index=2
  while(is.na(pattern) || pattern == "") {
    if(index > length(dateList)) { #everything is NA or empty string
      return(rep(as.Date(NA), length(dateList)))
    }
    pattern = dateList[index]
    index=index+1
  }
  if(grepl("[0-9]+/[0-9]+/[0-9][0-9][0-9][0-9]", pattern)) {
    return(cleanDate(as.Date(dateList, format = "%m/%d/%Y")))
  } else if(grepl("[0-9][0-9][0-9][0-9]-[0-9]+-[0-9]+", pattern)) {
    return(cleanDate(as.Date(dateList, format = "%Y-%m-%d")))
  } else if(grepl("[0-9]+-[0-9]+-[0-9][0-9][0-9][0-9]", pattern)) {
    return(cleanDate(as.Date(dateList, format = "%m-%d-%Y")))
  } else if(grepl("[A-Z][a-z][a-z] [0-9]+, [0-9][0-9][0-9][0-9]", pattern)) { #abbreviated month
    return(cleanDate(as.Date(dateList, format = "%B %d, %Y")))
  } else if(grepl("[A-Z][a-z]* [0-9]+, [0-9][0-9][0-9][0-9]", pattern)) { #full month
    return(cleanDate(as.Date(dateList, format = "%b %d, %Y")))
  } else {
    stop("Invalid date format:", pattern)
  }
}

##function that returns a vector of the SNP distance of all epi links in epi
##dist = SNP distance matrix
##epi = list of epi links
getSNPDistance <- function(epi, dist) {
  if(all(is.na(dist)) | all (is.na(epi))) {
    return(NA)
  } else {
    snp = rep(NA, nrow(epi))
    for(r in 1:nrow(epi)) {
      if(epi$case1[r] %in% row.names(dist) & epi$case2[r] %in% row.names(dist)) {
        snp[r] = dist[row.names(dist)==epi$case1[r], colnames(dist)==epi$case2[r]]
      }
    }
  }
  return(snp)
}

##for the given file, return a distance matrix that has filled in the upper triangle and rounded to the nearest whole number
##compared to LITT version, does not de-duplicate or convert to state case number
##returns a matrix: fill in whole matrix, with row and column names the state case number and SNP distances rounded to nearest whole number
##the fileName is expected to point to a SNP distance matrix or lower triangle, such as from BioNumerics (include file path if not in working dir)
##log is file to write results to
##if appendlog is true, append results to log file; otherwise overwrite (needed because may read distance matrix first)
formatDistanceMatrixWithoutStno <- function(fileName, log, appendlog=T) { #formerly formatDistanceMatrixWithoutDedup; formatDistanceMatrixWithoutDedupOrStno
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
    } else if(make.names(row.names(mat))[1] != colnames(mat)[1]) {
      mat = as.matrix(read.table(fileName, sep="", header = F, row.names = 1))
    }
    colnames(mat) = row.names(mat) #fix the . for non character spaces
  } else if(endsWith(fileName, ".xls") || endsWith(fileName, ".xlsx")) {
    df = read.xlsx(fileName, sheetIndex = 1)
    mat = as.matrix(df[,-1])
    row.names(mat) = df[,1]
    colnames(mat) = as.character(df[,1]) #fix the . for non character spaces
  } else if(endsWith(fileName, ".csv")) {
    mat = as.matrix(read.table(fileName, sep=",", header = T, row.names = 1))
    colnames(mat) = row.names(mat) #fix the . for non character spaces
  } else {
    cat("Distance matrix should be in a text, Excel or CSV file format. SNP distances will not be used.\r\n", file = log, append = T)
    return(NA)
  }
  colnames(mat) = sub("^X", "", colnames(mat))
  
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
  
  ##check for duplicates
  ids = rownames(mat)
  dup = duplicated(ids)
  if(any(dup)) {
    cat("Distance matrix contains multiple distances for these IDs: ", paste(unique(ids[dup]), collapse=", "), "\r\n", file = log, append = T)
    
    ##warning if any are different
    for(d in unique(ids[dup])) {
      sub = mat[rownames(mat)==d,]
      same = sapply(2:nrow(sub), function(r) {all(sub[1,]==sub[r,])})
      if(!all(same)) {
        cat("WARNING: ", d, " is present multiple times in the distance matrix, with different distances.\r\n", file = log, append = T)
      }
    }
    cat("The first appearance of these cases in the matrix will be used.\n", file = log, append = T)
    
    mat = mat[!dup, !dup]
  }
  
  ##stop if there are missing values
  if(any(is.na(mat))) {
    cat("There is missing data in the distance matrix. All of the lower triangle must be filled in.\r\n", file = log, append = T)
    stop("Distance matrix is missing data.")
  }
  
  return(mat)
}


##for the given list of dates, return the earliest, with a warning if dates are not all the same
##to be used to remove duplicated cases in inputs
##warnID is the ID of the case associated with the dates, warnVar is the variable being tested
##if min is true, return the minumum case, otherwise return the max
##if message is true, print a warning
getMin <- function(dates, warnID, warnVar, min=T, message=T) {
  all = dates
  dates = convertToDate(dates)
  if(all(is.na(dates))) {
    return(NA)
  } else {
    if(any(is.na(dates))) { #some but not all missing
      if(message)
        warning(paste(warnID, "has multiple", warnVar, "rows, some of which are missing data; rows missing data removed"))
      dates = dates[!is.na(dates)]
    }
    if(length(dates) == 1) {
      return(all[convertToDate(all) == dates][1])
    } else {
      val = dates[1]
      if(length(unique(dates)) != 1) {
        if(message) {
          if(min) {
            warning(paste(warnID, "has multiple", warnVar, "values; earliest date will be used"))
          } else {
            warning(paste(warnID, "has multiple", warnVar, "values; latest date will be used"))
          }
        }
        if(min) {
          val = min(dates)
        } else {
          val = max(dates)
        }
      }
      return(all[convertToDate(all) == val][1])
    }
  }
}

##function that takes the dataframe df from Rshiny fileInput and returns NA if the file does not exist exists or the data frame from the file otherwise
##per https://shiny.rstudio.com/reference/shiny/0.14/fileInput.html : 
# after the user selects and uploads a file, it will be a data frame with 'name', 'size', 'type', and 'datapath' columns. The 'datapath' column will contain the local filenames where the data can be found.
readShinyInputFile <- function(df) {
  if(is.null(df)) {
    return(NA)
  } 
  fname = df$datapath
  # print(fname)
  if(!file.exists(fname)) {
    return(NA)
  }
  if(endsWith(fname, "xlsx")) {
    return(read.xlsx(fname, sheetIndex = 1))
  } else if(endsWith(fname, "csv")) {
    return(read.table(fname, header = T, sep = ","))
  } else if(endsWith(fname, "txt")) {
    return(read.table(fname, header = T, sep = "\t"))
  }
}