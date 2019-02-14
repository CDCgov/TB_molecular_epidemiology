library(shiny)
library(shinyjs)
source("littGims.R")

# gimsVars = c("HOMELESS", "HIVSTAT", "CORRINST", "LONGTERM", "IDU", "NONIDU", "ALCOHOL", 
#              "OCCUHCW", "OCCUCORR",
#              "RISKTNF", "RISKORGAN", "RISKDIAB", "RISKRENAL", "RISKIMMUNO")
gimsVars = c("IDU", "NONIDU", "ALCOHOL", "HOMELESS", "AnyCorr")

# Define UI ----
ui <- fluidPage(
  titlePanel("LITT with TB GIMS"),
  useShinyjs(),
  fluidRow(
    column(3, align="center",
           h3("Input files")),
    column(3, align="center",
           h3("Set up")),
    column(1),
    column(5, align="center",
           h3("Advanced Options"),
           h4("GIMS risk factors to include with weights"))
  ),
  fluidRow(
    column(3,
           fileInput("caseData", "Case data", accept=c(".xlsx", ".csv")),
           fileInput("epi", "Epi link list", accept=c(".xlsx", ".csv")),
           fileInput("distMatrix", "SNP distance matrix", accept=c(".xlsx", ".csv", ".txt")),
           checkboxInput("BNmat",
                         "Distance matrix from BioNumerics", value=T),
           checkboxInput("writeDist",
                         "Output cleaned distance matrix", value=T)),
    column(3,
           textInput("prefix", "Name prefix for output files"),
           sliderInput("snpCutoff", "SNP cutoff", min=0, max=10, value=5, step=1, round=T),
           checkboxInput("writeDate", "Output dates used in IP calculations"),
           br(),
           h3("Included cases"),
           fileInput("caseListFile", "Upload list of state case numbers to analyze", accept=c(".xlsx", ".csv")),
           textAreaInput("caseListManual", "Type list of state case numbers", rows=5)),
    column(1),
    column(width=5, 
           fluidRow(width=12,
                    column(6, align="center",
                           lapply(gimsVars[1:3], function(g) {
                             sliderInput(g, g, min=-1, max=10, value=-1, step=1, round=T)
                           })),
                    column(6, align="center",
                           lapply(gimsVars[4:5], function(g) {
                             sliderInput(g, g, min=-1, max=10, value=-1, step=1, round=T)
                           }))),
           fluidRow(
             column(12, 
                    helpText("Weight > 0 means RF included in ranking. Weight = 0 means RF not in ranking but in outputs. Weight = -1 means RF not in ranking or outputs.",
                             align="left"),
                    br(),
                    h4("Additional Risk Factors", align="center"),
                    fileInput("rfTable", "Table of risk factor weights", accept=c(".xlsx", ".csv")))))),
  fluidRow(column(12, align="center",
                  actionButton("clear", "Clear inputs"),
                  actionButton("run", "Run", style='background-color:royalblue; color:white; padding:20px 40px'), #https://www.w3schools.com/css/css3_buttons.asp
                  downloadButton("downloadData", "Download Results"),                
                  br(),
                  br(),
                  htmlOutput("message")))
)

##text to format html string to increase message font size
outputfontsizestart = "<font font-size=\"30px\"><b>"
outputfontsizeend = "</b></font>"

##set up for downloading results
dash <- .Platform$file.sep #file separator
tmpdir = tempdir() #temporary directory to write outputs
tmpdir = paste(tmpdir, dash, sep="")
outfiles = NA #list of output files
incCases = "" #list of cases to include in analysis
delCases = "MRCA" #list of cases that were deleted by the user in the manual text field

# Define server logic ----
server <- function(input, output, session) {
  shinyjs::hide("downloadData")
  
  rv <- reactiveValues(clCaseData = F,
                       clEpi = F,
                       clDist = F,
                       clRF = F,
                       clCaseList = F, #variables that if true, indicate that clear has been hit but a new table has not been uploaded
                       # incCases = "", #list of cases to include in analysis
                       # delCases = "", #list of cases that were deleted by the user in the manual text field
                       inputCaseList = F, #if true, case list was input, so don't update list when other files are uploaded
                       caseData = NA, #add the caseData, dist, and epi in case they get read to get case list so don't read twice
                       dist = NA,
                       epi = NA)
  
  ##run LATTE when action button hit
  observeEvent(input$run, {
    ##case data
    if(rv$clCaseData) {
      caseData = NA
    } else if(all(is.na(rv$caseData))) {
      caseData = readShinyInputFile(input$caseData)
    } else {
      caseData = rv$caseData
    }
    if(all(is.na(caseData)) | rv$clCaseData) {
      output$message <- renderText({paste(outputfontsizestart, "No case data. Please input a case data table.", outputfontsizeend, sep="")})
      return(NULL)
    }
    ##set up progress and output prefix
    progress <- Progress$new(session, min=-2, max=13)
    on.exit(progress$close())
    progress$set(message = "Running LITT")
    output$message <- renderText({paste(outputfontsizestart, "Analyzing data", outputfontsizeend, sep="")})
    progress$set(value=0)
    outPrefix = paste(tmpdir, input$prefix, sep="")
    ##distance matrix, epi and rf table
    if(rv$clDist) {
      dist = NA
    } else if(all(is.na(rv$dist))){
      dist = readShinyDistanceMatrix(input$distMatrix, bn=input$BNmat, log = paste(outPrefix, defaultLogName, sep=""))
    } else {
      dist = rv$dist
    }
    if(rv$clEpi) {
      epi = NA
    } else if(all(is.na(rv$epi))){
      epi = readShinyInputFile(input$epi)
    } else {
      epi = rv$epi
    }
    if(rv$clRF) {
      rf = NA
    } else {
      rf = readShinyInputFile(input$rfTable)
    }
    ##gims risk factors
    gimsRF = data.frame(variable = gimsVars,
                        weight = sapply(1:length(gimsVars), function(i){input[[gimsVars[i]]]}))
    gimsRF = gimsRF[gimsRF$weight >= 0,]
    if(nrow(gimsRF) < 1) {
      gimsRF = NA
    }
    
    ##case list
    cases = NA
    if(input$caseListManual != "") {
      cases = strsplit(input$caseListManual, "\n")[[1]]
    }
    
    res = tryCatch({
      littres = littGims(outPrefix = outPrefix,
                         cases = cases,
                         caseData = caseData,
                         dist = dist,
                         epi = epi,
                         SNPcutoff = input$snpCutoff,
                         rfTable = rf,
                         gimsRiskFactor = gimsRF,
                         writeDate = input$writeDate,
                         writeDist = input$writeDist,
                         appendlog = F,
                         progress = progress)
      outfiles <<- littres$outputFiles
      output$message <- renderText({paste(outputfontsizestart, "Analysis complete", outputfontsizeend, sep="")})
      shinyjs::show("downloadData")
    }, error = function(e) {
      outfiles <<- paste(tmpdir, input$prefix, defaultLogName, sep="")
      output$message <- renderText({paste(outputfontsizestart, "Error detected:<br/>", geterrmessage(),
                                          "<br/>Download and view log for more details.", outputfontsizeend, sep="")})
      cat(geterrmessage(), file = outfiles, append = T)
      shinyjs::show("downloadData")
    })
  })
  
  ##zip and download outputs
  output$downloadData <- downloadHandler(
    filename = function() {
      paste(input$prefix, "LITT.zip", sep="")
    },
    content = function(fname) {
      outfiles = sub(tmpdir, "", outfiles, fixed=T)
      currdir = getwd()
      setwd(tmpdir)
      zip(zipfile = fname, files = outfiles)
      if(file.exists(paste0(fname, ".zip"))) {
        file.rename(paste0(fname, ".zip"), fname)
      }
      # output$message <- renderText({paste(outputfontsizestart, "Download complete", outputfontsizeend, sep="")})
      setwd(currdir)
    },
    contentType = "application/zip"
  )
  
  ##clear inputs if clear button is clicked
  observeEvent(input$clear, {
    updateTextInput(session, "prefix", value="")
    updateSliderInput(session, "snpCutoff", value=snpDefaultCut)
    lapply(gimsVars, function(g) {updateSliderInput(session, g, value=-1)})
    output$message <- renderText({""})
    shinyjs::hide("downloadData")
    reset("caseData")
    reset("epi")
    reset("distMatrix")
    reset("rfTable")
    reset("caseListFile")
    rv$clCaseData <- T
    rv$clEpi <- T
    rv$clDist <- T
    rv$clRF <- T
    rv$clCaseList <- T
    incCases <<- ""
    delCases <<- "MRCA"
    rv$inputCaseList <- F
    updateCheckboxInput(session, "BNmat", value=T)
    updateCheckboxInput(session, "writeDist", value=T)
    updateCheckboxInput(session, "writeDate", value=F)
    updateTextInput(session,"caseListManual",value="") 
  })
  
  ##if any inputs change, hide download button and remove output message
  observe({
    input$prefix
    output$message <- renderText({""})
    shinyjs::hide("downloadData")
  })
  observe({
    input$snpCutoff
    output$message <- renderText({""})
    shinyjs::hide("downloadData")
  })
  observe({
    input$writeDate
    output$message <- renderText({""})
    shinyjs::hide("downloadData")
  })
  observe({
    input$caseData
    if(!is.null(input$caseData) & !rv$inputCaseList & !rv$clCaseData) {
      rv$caseData = fixStcasenoName(readShinyInputFile(input$caseData))
      incases = as.character(rv$caseData$STCASENO)
      if(all(incCases=="")) {
        incCases <<- sort(unique(incases))
      } else {
        incCases <<- sort(unique(c(incCases[incCases!=""], incases)))
      }
      incCases <<- incCases[!incCases %in% delCases] #remove previously deleted cases
      updateTextInput(session, "caseListManual", value=paste(incCases, collapse="\n"))
    }
    rv$clCaseData <- F
    output$message <- renderText({""})
    shinyjs::hide("downloadData")
  })
  observe({
    input$epi
    if(!is.null(input$epi) & !rv$inputCaseList & !rv$clEpi) {
      rv$epi = fixEpiNames(fixStcasenoName(readShinyInputFile(input$epi)), log = paste(tmpdir, input$prefix, defaultLogName, sep=""))
      if(all(incCases=="")) {
        incCases <<- sort(unique(as.character(c(rv$epi$case1, rv$epi$case2))))
      } else {
        incCases <<- sort(unique(as.character(c(incCases[incCases!=""], rv$epi$case1, rv$epi$case2))))
      }
      incCases <<- incCases[!incCases %in% delCases] #remove previously deleted cases
      updateTextInput(session, "caseListManual", value=paste(incCases, collapse="\n"))
    }
    rv$clEpi <- F
    output$message <- renderText({""})
    shinyjs::hide("downloadData")
  })
  observe({
    input$distMatrix
    if(!is.null(input$distMatrix) & !rv$inputCaseList & !rv$clDist) {
      rv$dist = fixStcasenoName(readShinyDistanceMatrix(input$distMatrix, bn=input$BNmat, 
                                                        log = paste(tmpdir, input$prefix, defaultLogName, sep="")))
      if(all(incCases=="")) {
        incCases <<- sort(unique(row.names(rv$dist)))
      } else {
        incCases <<- sort(unique(c(incCases[incCases!=""], row.names(rv$dist))))
      }
      incCases <<- incCases[!incCases %in% delCases] #remove previously deleted cases
      updateTextInput(session, "caseListManual", value=paste(incCases, collapse="\n"))
    }
    rv$clDist <- F
    output$message <- renderText({""})
    shinyjs::hide("downloadData")
  })
  observe({
    input$writeDist
    output$message <- renderText({""})
    shinyjs::hide("downloadData")
  })
  observe({
    input$BNmat
    output$message <- renderText({""})
    shinyjs::hide("downloadData")
  })
  observe({
    input$rfTable
    rv$clRF = F
    output$message <- renderText({""})
    shinyjs::hide("downloadData")
  })
  observe({
    input$caseListFile
    if(!is.null(input$caseListFile) & !rv$clCaseList) {
      rv$inputCaseList = T
      res = fixStcasenoName(readShinyInputFile(input$caseListFile))
      incCases <<- res$STCASENO
      updateTextInput(session, "caseListManual", value=paste(incCases, collapse="\n"))
    }
    rv$clCaseList = F
    output$message <- renderText({""})
    shinyjs::hide("downloadData")
  })
  observe({
    input$caseListManual
    if(input$caseListManual != "") {
      rvcases = incCases
      mancases = strsplit(input$caseListManual, "\n")[[1]]
      delCases <<- unique(c(delCases, rvcases[!rvcases %in% mancases])) #list of cases deleted by user, to not include with future uploads
    }
    output$message <- renderText({""})
    shinyjs::hide("downloadData")
  })
  observe({
    input$IDU
    output$message <- renderText({""})
    shinyjs::hide("downloadData")
  })
  observe({
    input$NONIDU
    output$message <- renderText({""})
    shinyjs::hide("downloadData")
  })
  observe({
    input$ALCOHOL
    output$message <- renderText({""})
    shinyjs::hide("downloadData")
  })
  observe({
    input$HOMELESS
    output$message <- renderText({""})
    shinyjs::hide("downloadData")
  })
  observe({
    input$AnyCorr
    output$message <- renderText({""})
    shinyjs::hide("downloadData")
  })
}

# Run the app ----
shinyApp(ui = ui, server = server)