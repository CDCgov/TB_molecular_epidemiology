library(shiny)
library(shinyjs)

source("latte.R")

# Define UI ----
ui <- fluidPage(
  titlePanel("LATTE"),
  fluidRow(
    column(6,
           h3("Set up inputs"),
           fileInput("locTab", "Table of dates in locations", accept=c(".xlsx", ".csv")),
           fileInput("ipTab", "Table of infectious periods", accept=c(".xlsx", ".csv")),
           br(),
           br(),
           h3("Set up outputs"),
           textInput("prefix", "Name prefix for output files")),
    useShinyjs(),
    column(6,
           h3("Set up link definitions"),
           radioButtons("linkType", "Type of link", 
                        choiceNames = list("Epi link", "IP epi link"),
                        choiceValues = list("epi", "ipepi"),
                        selected = "epi"),
           sliderInput("epicutoff", "Number of days cases must overlap in a location to form a definite or probable epi link", min=0, max=30, value=defaultCut, step=1, round=T),
           sliderInput("ipepicutoff", "Number of days cases must overlap each other and an IP to form an IP epi link", min=0, max=30, value=defaultCut, step=1, round=T),
           checkboxInput("removeAfter", "Check if want to include overlaps that occur after either IP end as potential IP epi links (e.g. to identify potential re-exposure during a contact investigation)")),
    
    fluidRow(column(12, align="center",
                    actionButton("run", "Run", style='background-color:royalblue; color:white; padding:20px 40px'))), #https://www.w3schools.com/css/css3_buttons.asp
    fluidRow(column(12, align="center",
                    br(),
                    br(),
                    htmlOutput("message"))), #use instead of textOutput so can change font of returning string: https://stackoverflow.com/questions/24049159/change-the-color-and-font-of-text-in-shiny-app
    fluidRow(column(12, align="center",
                    br(),
                    downloadButton("downloadData", "Download Results")))
    
  ))

##text to format html string to increase message font size
outputfontsizestart = "<font font-size=\"30px\"><b>"
outputfontsizeend = "</b></font>"

dash <- .Platform$file.sep #file separator
tmpdir = tempdir() #temporary directory to write outputs
tmpdir = paste(tmpdir, dash, sep="")
outfiles = NA #list of output files

# Define server logic ----
server <- function(input, output, session) {
  shinyjs::hide("downloadData")
  
  ##run LATTE when action button hit
  observeEvent(input$run, {
    loc = readShinyInputFile(input$locTab)
    if(all(is.na(loc))) {
      output$message <- renderText({paste(outputfontsizestart, "No location data; please input a location table", outputfontsizeend, sep="")})
      return(NULL)
    }
    progress <- Progress$new(session, min=-1, max=11) 
    on.exit(progress$close())
    progress$set(message = "Running LATTE")
    output$message <- renderText({paste(outputfontsizestart, "Starting analysis", outputfontsizeend, sep="")})
    ipEpiLink = ifelse(input$linkType == "ipepi", T, F)
    cutoff = ifelse(input$linkType == "ipepi", input$ipepicutoff, input$epicutoff)
    progress$set(value=0)
    outPrefix = paste(tmpdir, input$prefix, sep="")
    res = tryCatch({
      latteres = latteWithOutputs(outPrefix = outPrefix,
                                  loc = loc,
                                  ip = readShinyInputFile(input$ipTab),
                                  cutoff = cutoff,
                                  ipEpiLink = ipEpiLink,
                                  removeAfter = input$removeAfter,
                                  progress = progress)
      outfiles <<- latteres$outputFiles
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
      paste(input$prefix, "LATTE.zip", sep="")
    },
    content = function(fname) {
      outfiles = sub(tmpdir, "", outfiles, fixed=T)
      setwd(tmpdir)
      zip(zipfile = fname, files = outfiles)
      # print(fname)
      # if(!endsWith(fname, "zip")) {
      #   file.rename(fname, paste(fname, ".zip", sep=""))
      # }
      # print(fname)
      if(file.exists(paste0(fname, ".zip"))) {
        file.rename(paste0(fname, ".zip"), fname)
      }
      output$message <- renderText({paste(outputfontsizestart, "Download complete", outputfontsizeend, sep="")})
    },
    contentType = "application/zip"
  )
  
  ##if any inputs change, hide download button and remove output message
  observe({ 
    input$locTab
    output$message <- renderText({""})
    shinyjs::hide("downloadData")
  })
  observe({
    input$ipTab
    output$message <- renderText({""})
    shinyjs::hide("downloadData")
  })
  observe({
    input$prefix
    output$message <- renderText({""})
    shinyjs::hide("downloadData")
  })
  observe({
    input$linkType
    output$message <- renderText({""})
    shinyjs::hide("downloadData")
  })
  observe({
    input$epicutoff
    output$message <- renderText({""})
    shinyjs::hide("downloadData")
  })
  observe({
    input$ipepicutoff
    output$message <- renderText({""})
    shinyjs::hide("downloadData")
  })
  observe({
    input$removeAfter
    output$message <- renderText({""})
    shinyjs::hide("downloadData")
  })
  
  ##show correct cutoff and check box, depending on link type
  observe({
    if(input$linkType == "ipepi") {
      shinyjs::hide("epicutoff")
      shinyjs::show("ipepicutoff")
      shinyjs::show("removeAfter")
    } else {
      shinyjs::show("epicutoff")
      shinyjs::hide("ipepicutoff")
      shinyjs::hide("removeAfter")
    }
  })
}

# Run the app ----
shinyApp(ui = ui, server = server)
