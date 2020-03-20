###LATTE Shiny interface

library(shiny)
library(shinyjs)
library(shinyWidgets)
library(zip)

source("latte.R")
source("latteNoTime.R")

# Define UI ----
ui <- fluidPage(
  # titlePanel("LATTE"),
  # fluidRow(column(10, offset=2, titlePanel("LATTE"))),
  # fluidRow(column(12, align="center", titlePanel("LATTE"))),
  fluidRow(column(12, 
                  titlePanel(tagList(span("Location And Time To Epi (LATTE)",
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
                  windowTitle = "LATTE"))),
  useShinyjs(),
  tabsetPanel(
    tabPanel("With Date Data",
             # fluidRow(column(1),
             #          column(11,
             #                 br(),
             #          p("Identify overlaps in time and location"))),
             fluidRow(br(),
                      h4("Identify overlaps in time and location", align="center")),
             fluidRow(
               column(2),
               column(4,
                      h3("Set up inputs"),
                      p("Warning: do not upload personally identifiable information (PII)", style="color:red"),
                      fileInput("locTab", "Table of dates in locations (required)", accept=c(".xlsx", ".csv")),
                      fileInput("ipTab", "Table of infectious periods (IP)", accept=c(".xlsx", ".csv")),
                      br(),
                      br(),
                      h3("Set up outputs"),
                      textInput("prefix", "Name prefix for output files")),
               column(4,
                      h3("Set up link definitions"),
                      radioButtons("linkType", "Type of link",
                                   choiceNames = list("Epi link (all possible overlaps)", "IP epi link (only overlaps during an IP)"),
                                   choiceValues = list("epi", "ipepi"),
                                   selected = "epi"),
                      sliderInput("epicutoff", tags$div("Overlap threshold", tags$br(), 
                                                        tags$p("This is the number of days people must overlap in a location to form a definite or probable epi link.", style="font-size: 85%; font-weight:100;")),  
                                  min=0, max=30, value=defaultCut, step=1, round=T),
                      sliderInput("ipepicutoff", tags$div("Overlap threshold", tags$br(), 
                                                          tags$p("This is the number of days people must overlap each other and an IP to form an IP epi link.", style="font-size: 85%; font-weight:100;")), 
                                  min=0, max=30, value=defaultCut, step=1, round=T),
                      # checkboxInput("ipCasesOnly", "Only include overlaps with people that have an IP (do not look for overlaps between people when neither person has an IP)", value=T),
                      checkboxInput("removeAfter", tags$div(tags$b("Include re-infection"), tags$br(), 
                                                            tags$p("For case-case overlaps, consider re-infection regardless of which IP comes first", style="font-size: 85%; font-weight:100;")))),
               # actionButton("clear", "Clear inputs")),
               column(2)),
             
             fluidRow(column(12, align="center",
                             br(),
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
                             downloadButton("downloadData", "Download Results"))
                      
             )),
             # ))))
    tabPanel("Without Date Data",
             fluidRow(br(),
                      h4("Identify all possible pairs from lists of people", align="center")),
             fluidRow(
               column(2),
               column(4,
                      h3("Set up inputs"),
                      p("Warning: do not upload personally identifiable information (PII)", style="color:red"),
                      fileInput("noTimeTab", "Table of people in locations (required)", accept=c(".xlsx", ".csv")),
                      br(),
                      br(),
                      h3("Set up outputs"),
                      textInput("noTimePrefix", "Name prefix for output files")),
               column(4,
                      h3("Set up link definitions"),
                      helpText("Strength option will be applied to all columns (no strength means strength will be left blank) unless user chooses a column-specific strength. Valid strengths are definite, probable, possible,  or none (none will leave strength blank)."),
                      radioButtons("noTimeStrength", "Strength of link",
                                   choiceNames = list("Definite epi link", "Probable epi link", "Possible epi link",
                                                      "No strength", "Column-specific strength"),
                                   choiceValues = list("definite", "probable", "possible", "none", "custom"),
                                   selected = "probable"),
                      column(2),
                      column(10,
                             fileInput("noTimeCustomStrength", "Custom strength table", accept=c(".xlsx", ".csv")))),
               column(2)),

             fluidRow(column(12, align="center",
                             br(),
                             actionButton("noTimeClear", "Clear inputs"),
                             br(),
                             br(),
                             actionButton("noTimeRun", "Run", style='background-color:royalblue; color:white; padding:20px 40px'))), #https://www.w3schools.com/css/css3_buttons.asp
             fluidRow(column(12, align="center",
                             br(),
                             br(),
                             htmlOutput("noTimeMessage"))), #use instead of textOutput so can change font of returning string: https://stackoverflow.com/questions/24049159/change-the-color-and-font-of-text-in-shiny-app
             fluidRow(column(12, align="center",
                             br(),
                             downloadButton("noTimeDownloadData", "Download Results"))))))


##text to format html string to increase message font size
outputfontsizestart = "<font font-size=\"30px\"><b>"
outputfontsizeend = "</b></font>"

##set up for downloading results
dash <- .Platform$file.sep #file separator
tmpdir = tempdir() #temporary directory to write outputs
tmpdir = paste(tmpdir, dash, sep="")
outfiles = NA #list of output files
notime.outfiles = NA #list of output files

# Define server logic ----
server <- function(input, output, session) {
  shinyjs::hide("downloadData")
  
  rv <- reactiveValues(clLoc = F,
                       clIP = F,
                       clNTtab = F,
                       clNTcust = F) #variables that if true, indicate that clear has been hit but a new table has not been uploaded
  
  ##run LATTE when action button hit
  observeEvent(input$run, {
    if(rv$clLoc) {
      loc = NA
    } else {
      loc = readShinyInputFile(input$locTab)
    }
    if(all(is.na(loc))) {
      output$message <- renderText({paste(outputfontsizestart, "No location data; please input a location table", outputfontsizeend, sep="")})
      return(NULL)
    }
    progress <- Progress$new(session, min=-1, max=12)
    on.exit(progress$close())
    progress$set(message = "Running LATTE")
    output$message <- renderText({paste(outputfontsizestart, "Starting analysis", outputfontsizeend, sep="")})
    # ipEpiLink = ifelse(input$linkType == "ipepi", T, F)
    cutoff = ifelse(input$linkType == "ipepi", input$ipepicutoff, input$epicutoff)
    progress$set(value=0)
    outPrefix = paste(tmpdir, input$prefix, sep="")
    if(rv$clIP) {
      ip = NA
    } else {
      ip = readShinyInputFile(input$ipTab)
    }
    res = tryCatch({
      latteres = latteWithOutputs(outPrefix = outPrefix,
                                  loc = loc,
                                  ip = ip, #readShinyInputFile(input$ipTab),
                                  cutoff = cutoff,
                                  ipEpiLink = input$linkType == "ipepi",
                                  removeAfter = !input$removeAfter,
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
  
  ##run LATTE without time when action button hit
  observeEvent(input$noTimeRun, {
    if(rv$clNTtab) {
      tab = NA
    } else {
      tab = input$noTimeTab$datapath
    }
    if(all(is.na(tab))) {
      output$noTimeMessage <- renderText({paste(outputfontsizestart, "No people table; please input a table of people in locations.", outputfontsizeend, sep="")})
      return(NULL)
    }
    progress <- Progress$new(session, min=0, max=7)
    on.exit(progress$close())
    progress$set(message = "Running LATTE")
    output$noTimeMessage <- renderText({paste(outputfontsizestart, "Starting analysis", outputfontsizeend, sep="")})
    cutoff = ifelse(input$linkType == "ipepi", input$ipepicutoff, input$epicutoff)
    progress$set(value=0)
    outPrefix = paste(tmpdir, input$noTimePrefix, sep="")
    if(rv$clNTcust) {
      custom = NA
    } else {
      custom = readShinyInputFile(input$noTimeCustomStrength)
    }
    res = tryCatch({
      latteres = latteNoTimeWithOutputs(outPrefix = outPrefix,
                                        fname = tab,  
                                        strength = input$noTimeStrength,
                                        custom = custom,
                                        progress = progress)
      notime.outfiles <<- latteres$outputFiles
      output$noTimeMessage <- renderText({paste(outputfontsizestart, "Analysis complete", outputfontsizeend, sep="")})
      shinyjs::show("noTimeDownloadData")
    }, error = function(e) {
      notime.outfiles <<- paste(tmpdir, input$noTimePrefix, defaultNoTimeLogName, sep="")
      output$noTimeMessage <- renderText({paste(outputfontsizestart, "Error detected:<br/>", geterrmessage(),
                                          "<br/>Download and view log for more details.", outputfontsizeend, sep="")})
      cat(geterrmessage(), file = notime.outfiles, append = T)
      shinyjs::show("noTimeDownloadData")
    })
  })
  
  ##zip and download outputs
  output$downloadData <- downloadHandler(
    filename = function() {
      paste(input$prefix, "LATTE.zip", sep="")
    },
    content = function(fname) {
      outfiles = sub(tmpdir, "", outfiles, fixed=T)
      currdir = getwd()
      setwd(tmpdir)
      zipr(zipfile = fname, files = outfiles)
      if(file.exists(paste0(fname, ".zip"))) {
        file.rename(paste0(fname, ".zip"), fname)
      }
      output$message <- renderText({paste(outputfontsizestart, "Download complete", outputfontsizeend, sep="")})
      setwd(currdir)
    },
    contentType = "application/zip"
  )
  
  output$noTimeDownloadData <- downloadHandler(
    filename = function() {
      paste(input$noTimePrefix, "LATTE.zip", sep="")
    },
    content = function(fname) {
      notime.outfiles = sub(tmpdir, "", notime.outfiles, fixed=T)
      currdir = getwd()
      setwd(tmpdir)
      zipr(zipfile = fname, files = notime.outfiles)
      if(file.exists(paste0(fname, ".zip"))) {
        file.rename(paste0(fname, ".zip"), fname)
      }
      output$noTimeMessage <- renderText({paste(outputfontsizestart, "Download complete", outputfontsizeend, sep="")})
      setwd(currdir)
    },
    contentType = "application/zip"
  )
  
  ##clear inputs if clear button is clicked
  observeEvent(input$clear, {
    updateTextInput(session, "prefix", value="")
    updateTextInput(session, "noTimePrefix", value="")
    updateSliderInput(session, "epicutoff", value=defaultCut)
    updateSliderInput(session, "ipepicutoff", value=defaultCut)
    updateCheckboxInput(session, "removeAfter", value=F)
    # updateRadioButtons(session, "linkType", selected = epi)
    reset("linkType")
    reset("noTimeStrength")
    output$message <- renderText({""})
    output$noTimeMessage <- renderText({""})
    shinyjs::hide("downloadData")
    shinyjs::hide("noTimeDownloadData")
    reset("locTab")
    reset("ipTab")
    reset("noTimeTab")
    reset("noTimeCustomStrength")
    rv$clLoc <- T
    rv$clIP <- T
    rv$clNTtab <- T
    rv$clNTcust <- T
  })
  observeEvent(input$noTimeClear, {
    updateTextInput(session, "prefix", value="")
    updateTextInput(session, "noTimePrefix", value="")
    updateSliderInput(session, "epicutoff", value=defaultCut)
    updateSliderInput(session, "ipepicutoff", value=defaultCut)
    updateCheckboxInput(session, "removeAfter", value=F)
    reset("linkType")
    reset("noTimeStrength")
    output$message <- renderText({""})
    output$noTimeMessage <- renderText({""})
    shinyjs::hide("downloadData")
    shinyjs::hide("noTimeDownloadData")
    reset("locTab")
    reset("ipTab")
    reset("noTimeTab")
    reset("noTimeCustomStrength")
    rv$clLoc <- T
    rv$clIP <- T
    rv$clNTtab <- T
    rv$clNTcust <- T
  })
  
  ##if any inputs change, hide download button and remove output message
  observe({
    input$locTab
    rv$clLoc <- F
    output$message <- renderText({""})
    output$noTimeMessage <- renderText({""})
    shinyjs::hide("downloadData")
    shinyjs::hide("noTimeDownloadData")
  })
  observe({
    input$ipTab
    rv$clIP <- F
    output$message <- renderText({""})
    output$noTimeMessage <- renderText({""})
    shinyjs::hide("downloadData")
    shinyjs::hide("noTimeDownloadData")
  })
  observe({
    input$prefix
    output$message <- renderText({""})
    output$noTimeMessage <- renderText({""})
    shinyjs::hide("downloadData")
    shinyjs::hide("noTimeDownloadData")
  })
  observe({
    input$linkType
    output$message <- renderText({""})
    output$noTimeMessage <- renderText({""})
    shinyjs::hide("downloadData")
    shinyjs::hide("noTimeDownloadData")
  })
  observe({
    input$epicutoff
    output$message <- renderText({""})
    output$noTimeMessage <- renderText({""})
    shinyjs::hide("downloadData")
    shinyjs::hide("noTimeDownloadData")
  })
  observe({
    input$ipepicutoff
    output$message <- renderText({""})
    output$noTimeMessage <- renderText({""})
    shinyjs::hide("downloadData")
    shinyjs::hide("noTimeDownloadData")
  })
  observe({
    input$removeAfter
    output$message <- renderText({""})
    output$noTimeMessage <- renderText({""})
    shinyjs::hide("downloadData")
    shinyjs::hide("noTimeDownloadData")
  })
  observe({
    input$noTimeTab
    rv$clNTtab <- F
    output$message <- renderText({""})
    output$noTimeMessage <- renderText({""})
    shinyjs::hide("downloadData")
    shinyjs::hide("noTimeDownloadData")
  })
  observe({
    input$noTimePrefix
    output$message <- renderText({""})
    output$noTimeMessage <- renderText({""})
    shinyjs::hide("downloadData")
    shinyjs::hide("noTimeDownloadData")
  })
  observe({
    input$noTimeStrength
    output$message <- renderText({""})
    output$noTimeMessage <- renderText({""})
    shinyjs::hide("downloadData")
    shinyjs::hide("noTimeDownloadData")
  })
  observe({
    input$noTimeCustomStrength
    rv$clNTcust <- F
    output$message <- renderText({""})
    output$noTimeMessage <- renderText({""})
    shinyjs::hide("downloadData")
    shinyjs::hide("noTimeDownloadData")
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
  
  ##show custom strength upload only when selected
  observe({
    if(input$noTimeStrength == "custom") {
      shinyjs::show("noTimeCustomStrength")
    } else {
      shinyjs::hide("noTimeCustomStrength")
    }
  })
}

# Run the app ----
shinyApp(ui = ui, server = server)
