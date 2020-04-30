library(shiny)
library(shinyjs)
library(shinyWidgets)
library(zip)
source("littNoGims.R")

# Define UI ----
ui <- fixedPage(
  useShinyjs(),
  titlePanel(tagList(span("Logically Inferred Tuberculosis Transmission (LITT)",
                          # span(actionButton('help', 'help'),
                          span(dropdownButton(tags$style(".btn-custom {background-color: white; color: black; border-color: black;}"), #https://github.com/dreamRs/shinyWidgets/issues/126 
                                              # circle = F,
                                              # label = "?",
                                              size = "sm",
                                              status="custom",
                                              icon = icon("question"), #https://shiny.rstudio.com/reference/shiny/0.14/icon.html
                                              tooltip=tooltipOptions(title="Help"),
                                              right = T,
                                              h4("Help"),
                                              tags$style(HTML("#help{border-color: white; width:260px; text-align:left}")),
                                              actionButton("help", "User's manual & training presentation",
                                                           onclick="window.open('https://github.com/CDCgov/TB_molecular_epidemiology/tree/master/LITT_Documentation/LITT%20user%20manual%20and%20training%20presentation')"),
                                              actionButton("help", "Input file templates",
                                                           onclick="window.open('https://github.com/CDCgov/TB_molecular_epidemiology/tree/master/LITT_Documentation/LITT%20input%20file%20templates')"),
                                              actionButton("help", "Training datasets",
                                                           onclick="window.open('https://github.com/CDCgov/TB_molecular_epidemiology/tree/master/LITT_Documentation/LITT%20training%20datasets')"),
                                              actionButton("help", "Reference")),
                               style = "position:absolute;right:2em;"))), #https://stackoverflow.com/questions/54523349/place-actionbutton-on-right-side-of-titlepanel-in-shiny
             windowTitle = "LITT"),
  fluidRow(
    column(4, #align="center",
           h3("Input files", align="center"),
           p("Warning: do not upload personally identifiable information (PII)", style="color:red;text-align:center"),
           fileInput("caseData", "Case data table", accept=c(".xlsx", ".csv")),
           checkboxInput("keepExtraCDcol",
                         # "Output extra columns from the TB GIMS version"),
                         "Output extra columns in case data table", value=T),
           br(),
           fileInput("epi", "Epi link table", accept=c(".xlsx", ".csv")),
           fileInput("distMatrix", "SNP distance matrix", accept=c(".xlsx", ".csv", ".txt")),
           checkboxInput("writeDist",
                         "Include distance matrix in outputs", value=F)),
    column(4, #align="center",
           h3("Set up outputs", align="center"),
           textInput("prefix", "Name prefix for output files"),#, width="90%"),
           # sliderInput("snpCutoff", "SNP cutoff", min=0, max=7, value=5, step=1, round=T)),
           sliderInput("snpCutoff", tags$div("SNP cutoff", tags$br(), 
                                             tags$p("Maximum number of SNPs for two cases to be linked by transmission.", style="font-size: 85%; font-weight:100;")),
                                             min=0, max=7, value=5, step=1, round=T)), 
    column(4,
           h3("Advanced options", align="center"),
           fileInput("rfTable", "Table of risk factor weights", accept=c(".xlsx", ".csv")),
           helpText("This table contains a list of the columns in the case data table to use as risk factors, with their weights.",
                    "Variable names must exactly match the name of the column in the case data table."))),
  fluidRow(),
  fluidRow(column(12, align="center",
                  actionButton("clear", "Clear inputs"),
                  br(),
                  br(),
                  actionButton("run", "Run", style='background-color:royalblue; color:white; padding:20px 40px'))), #https://www.w3schools.com/css/css3_buttons.asp
  fluidRow(column(12, align="center",
                  br(),
                  br(),
                  htmlOutput("message"))), #use instead of textOutput so can change font of returning string: https://stackoverflow.com/questions/24049159/change-the-color-and-font-of-text-in-shiny-app
  fluidRow(column(12, align="center",
                  br(),
                  downloadButton("downloadData", "Download Results")))
)

##text to format html string to increase message font size
outputfontsizestart = "<font font-size=\"30px\"><b>"
outputfontsizeend = "</b></font>"

##set up for downloading results
dash <- .Platform$file.sep #file separator
tmpdir = tempdir() #temporary directory to write outputs
tmpdir = paste(tmpdir, dash, sep="")
outfiles = NA #list of output files

# Define server logic ----
server <- function(input, output, session) {
  shinyjs::hide("downloadData")
  
  rv <- reactiveValues(clCaseData = F,
                       clEpi = F,
                       clDist = F,
                       clRF = F) #variables that if true, indicate that clear has been hit but a new table has not been uploaded
  
  ##run LATTE when action button hit
  observeEvent(input$run, {
    if(rv$clCaseData) {
      caseData = NA
    } else {
      caseData = readShinyInputFile(input$caseData)
    }
    if(all(is.na(caseData)) | rv$clCaseData) {
      output$message <- renderText({paste(outputfontsizestart, "No case data. Please input a case data table.", outputfontsizeend, sep="")})
      return(NULL)
    }
    progress <- Progress$new(session, min=0, max=12) #max=6) (6 if no Excel tables)
    on.exit(progress$close())
    progress$set(message = "Running LITT")
    output$message <- renderText({paste(outputfontsizestart, "Analyzing data", outputfontsizeend, sep="")})
    progress$set(value=0)
    outPrefix = paste(tmpdir, input$prefix, sep="")
    cat("", file = paste(outPrefix, defaultLogName, sep=""), append=F) #re-initialize log on new runs
    if(rv$clDist) {
      dist = NA
    } else {
      dist = readShinyDistanceMatrix(input$distMatrix, bn=F, log = paste(outPrefix, defaultLogName, sep=""))
    }
    if(rv$clEpi) {
      epi = NA
    } else {
      epi = readShinyInputFile(input$epi)
    }
    if(rv$clRF) {
      rf = NA
    } else {
      rf = readShinyInputFile(input$rfTable)
    }
    res = tryCatch({
      littres = littNoGims(outPrefix = outPrefix,
                           caseData = caseData,
                           dist = dist, # readShinyDistanceMatrix(input$distMatrix, bn=F), #input$BNdist),
                           epi = epi, #readShinyInputFile(input$epi),
                           SNPcutoff = input$snpCutoff,
                           rfTable = rf, #readShinyInputFile(input$rfTable),
                           keepExtraCDcol = input$keepExtraCDcol,
                           writeDist = input$writeDist,
                           appendlog = T,
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
      zipr(zipfile = fname, files = outfiles)
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
    output$message <- renderText({""})
    shinyjs::hide("downloadData")
    reset("caseData")
    reset("epi")
    reset("distMatrix")
    reset("rfTable")
    rv$clCaseData <- T
    rv$clEpi <- T
    rv$clDist <- T
    rv$clRF <- T
    updateCheckboxInput(session, "keepExtraCDcol", value=T)
    updateCheckboxInput(session, "writeDist", value=F)
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
    input$caseData
    rv$clCaseData <- F
    output$message <- renderText({""})
    shinyjs::hide("downloadData")
  })
  observe({
    input$epi
    rv$clEpi <- F
    output$message <- renderText({""})
    shinyjs::hide("downloadData")
  })
  observe({
    input$distMatrix
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
    input$rfTable
    rv$clRF = F
    output$message <- renderText({""})
    shinyjs::hide("downloadData")
  })
  observe({
    input$keepExtraCDcol
    output$message <- renderText({""})
    shinyjs::hide("downloadData")
  })
}

# Run the app ----
shinyApp(ui = ui, server = server)