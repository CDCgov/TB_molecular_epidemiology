library(shiny)
library(shinyjs)
source("littNoGims.R")

# Define UI ----
ui <- fluidPage(
  titlePanel("LITT"),
  useShinyjs(),
  fluidRow(
    column(4, align="center",
           h3("Set up")),
    column(4, align="center",
           h3("Input files")),
    column(4, align="center",
           h3("Advanced options"))
  ),
  fluidRow(
    column(4,
           textInput("prefix", "Name prefix for output files"),
           sliderInput("snpCutoff", "SNP cutoff", min=0, max=10, value=5, step=1, round=T)),
    column(4,
           fileInput("caseData", "Case data table"),
           checkboxInput("cdFromGimsRun",
                         "Check if the case data comes from the TB GIMS version of LITT and you want to include the additional surveillance columns in this output"),
           br(),
           # helpText("Must include: sputum smear results, cavitation status, ",
           #          "whether a case is extrapulmonary only or pediatric, and ",
           #          "infectious period start and end.", 
           #          "Any additional columns will be treated as risk factors.",
           #          "See documentation for more details."),
           fileInput("epi", "Epi link table"),
           fileInput("distMatrix", "SNP distance matrix")),
           # checkboxInput("BNdist", 
           #               "Check if this is a table from BioNumerics with accession numbers that must be converted to state case number"), value=F),
    column(4,
           fileInput("rfTable", "Table of risk factor weights"),
           helpText("This table contains a list of the columns in the case data table to use a risk factors, with their weights.",
                    "Variable names must exactly match the name of the column in the case data table."))),
  fluidRow(),
  fluidRow(column(12, align="center",
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
  
  ##run LATTE when action button hit
  observeEvent(input$run, {
    caseData = readShinyInputFile(input$caseData)
    if(all(is.na(caseData))) {
      output$message <- renderText({paste(outputfontsizestart, "No case data; please input a case data table", outputfontsizeend, sep="")})
      return(NULL)
    }
    progress <- Progress$new(session, min=0, max=11) 
    on.exit(progress$close())
    progress$set(message = "Running LITT")
    output$message <- renderText({paste(outputfontsizestart, "Analyzing data", outputfontsizeend, sep="")})
    progress$set(value=0)
    outPrefix = paste(tmpdir, input$prefix, sep="")
    res = tryCatch({
      littres = littNoGims(outPrefix = outPrefix,
                           caseData = caseData,
                           dist = readShinyDistanceMatrix(input$distMatrix, bn=F), #input$BNdist),
                           epi = readShinyInputFile(input$epi),
                           SNPcutoff = input$snpCutoff,
                           rfTable = readShinyInputFile(input$rfTable),
                           cdFromGimsRun = input$cdFromGimsRun,
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
      output$message <- renderText({paste(outputfontsizestart, "Download complete", outputfontsizeend, sep="")})
      setwd(currdir)
    },
    contentType = "application/zip"
  )
  
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
    output$message <- renderText({""})
    shinyjs::hide("downloadData")
  })
  observe({
    input$epi
    output$message <- renderText({""})
    shinyjs::hide("downloadData")
  })
  observe({
    input$distMatrix
    output$message <- renderText({""})
    shinyjs::hide("downloadData")
  })
  observe({
    input$BNdist
    output$message <- renderText({""})
    shinyjs::hide("downloadData")
  })
  observe({
    input$rfTable
    output$message <- renderText({""})
    shinyjs::hide("downloadData")
  })
  observe({
    input$cdFromGimsRun
    output$message <- renderText({""})
    shinyjs::hide("downloadData")
  })
}

# Run the app ----
shinyApp(ui = ui, server = server)