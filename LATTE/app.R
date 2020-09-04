###LATTE Shiny interface
###Author: Kathryn Winglee

library(shiny)
library(shinyjs)
library(shinyWidgets)
library(zip)

source("latte.R")
source("latteNoTime.R")

# Define UI ----
ui <- fluidPage(
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
    tabPanel("Link analysis with date data",
             fluidRow(br(),
                      h4("Identify epi links based on overlaps in dates of people in locations", align="left", style="margin-left:30px;")),
             fluidRow(
               column(4,
                      h3("Set up inputs"),
                      p("Warning: do not upload personally identifiable information (PII)", style="color:red"),
                      fileInput("locTab", "Table of dates in locations (required)", accept=c(".xlsx", ".csv")),
                      fileInput("ipTab", "Table of infectious periods (IP) (required for IP epi link and IP Gantt chart)", accept=c(".xlsx", ".csv")),
                      br(),
                      br(),
                      h3("Set up outputs"),
                      textInput("prefix", "Name prefix for output files")),
               column(4,
                      h3("Set up link definitions"),
                      radioButtons("linkType", "Type of link",
                                   choiceNames = list("Epi link (all overlaps)", "IP epi link (only overlaps during an IP)"),
                                   choiceValues = list("epi", "ipepi"),
                                   selected = "epi"),
                      sliderInput("epicutoff", tags$div("Overlap threshold", tags$br(), 
                                                        tags$p("Number of days two people must overlap in a location to form a definite or probable epi link.", style="font-size: 85%; font-weight:100;")),  
                                  min=0, max=30, value=defaultCut, step=1, round=T),
                      sliderInput("ipepicutoff", tags$div("Overlap threshold", tags$br(), 
                                                          tags$p("Number of days two people must overlap each other and an IP to form a definite or probable IP epi link.", style="font-size: 85%; font-weight:100;")), 
                                  min=0, max=30, value=defaultCut, step=1, round=T),
                      # checkboxInput("ipCasesOnly", "Only include overlaps with people that have an IP (do not look for overlaps between people when neither person has an IP)", value=T),
                      checkboxInput("removeAfter", tags$div(tags$b("Include re-exposure"), tags$br(), 
                                                            tags$p("For case-case overlaps, consider re-exposure regardless of which IP comes first", style="font-size: 85%; font-weight:100;")))),
               column(4,
                      h3("Set up Gantt chart outputs"),
                      checkboxInput("locGantt", tags$div(tags$p("Generate location Gantt chart", style="font-size: 125%; font-weight:100;")), value = T),
                      div(style = "padding: 0px 0px; margin-top:-2em; margin-left:2em",#remove space between rows
                          fluidRow(#column(2),
                            column(12, 
                                   checkboxGroupInput("locGanttTime", 
                                                      tags$div(tags$p("Select time interval(s):", style="font-size: 100%; font-weight:100; padding: 0px 0px; margin-bottom:-2em")), 
                                                      c("Day"="day", "Week"="week", "Month"="month"),
                                                      selected="day")))),
                      checkboxInput("ipGantt", tags$div(tags$p("Generate IP Gantt chart", style="font-size: 125%; font-weight:100;")), value = T),
                      div(style = "padding: 0px 0px; margin-top:-2em; ; margin-left:2em",#remove space between rows
                          fluidRow(column(10, 
                                          checkboxGroupInput("ipGanttTime", 
                                                             tags$div(tags$p("Select time interval(s):", style="font-size: 100%; font-weight:100; padding: 0px 0px; margin-bottom:-2em")), 
                                                             c("Day"="day", "Week"="week", "Month"="month"),
                                                             selected="week")))))),
             
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
                             downloadButton("downloadData", "Download Results")))),
    tabPanel("Link analysis without date data",
             fluidRow(br(),
                      h4("Identify all possible links from grouped list(s) of people without date data", align="left", style="margin-left:30px;")),
             fluidRow(
               column(2),
               column(4,
                      h3("Set up inputs"),
                      p("Warning: do not upload personally identifiable information (PII)", style="color:red"),
                      fileInput("noTimeTab", "Table of grouped list(s) of people (required)", accept=c(".xlsx", ".csv")),
                      br(),
                      br(),
                      h3("Set up outputs"),
                      textInput("noTimePrefix", "Name prefix for output files")),
               column(4,
                      h3("Specify epi link strength"),
                      helpText("Strength options will be applied to all columns unless user selects \"custom strength(s) specified\"."),
                      radioButtons("noTimeStrength", "Strength of link",
                                   choiceNames = list("Definite epi link", "Probable epi link", "Possible epi link",
                                                      "No strength specified (strength left blank)", "Custom strength(s) specified"),
                                   choiceValues = list("definite", "probable", "possible", "none", "custom"),
                                   selected = "probable"),
                      column(2),
                      column(10,
                             fileInput("noTimeCustomStrength", tags$div("Custom strength(s) table", tags$br(), 
                                                                        tags$p("Valid strength values are definite, probable, possible, or no strength specified (strength left blank).", style="font-size: 85%; font-weight:100;")), 
                                       accept=c(".xlsx", ".csv")))),
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
                             downloadButton("noTimeDownloadData", "Download Results")))),
    tabPanel("Gantt charts only",
             fluidRow(br(),
                      h4("Generate Gantt charts without any link analysis", align="left", style="margin-left:30px;")),
             fluidRow(
               column(2),
               column(4,
                      h3("Set up inputs"),
                      p("Warning: do not upload personally identifiable information (PII)", style="color:red"),
                      fileInput("ganttLocTab", "Table of dates in locations (required for location Gantt chart)", accept=c(".xlsx", ".csv")),
                      fileInput("ganttIPTabGantt", "Table of infectious periods (IP) (required for IP Gantt chart)", accept=c(".xlsx", ".csv")),
                      br(),
                      br(),
                      h3("Set up outputs"),
                      textInput("GanttPrefix", "Name prefix for output files")),
               column(4,
                      h3("Set up Gantt chart outputs"),
                      checkboxInput("ganttLocGantt", tags$div(tags$p("Generate location Gantt chart", style="font-size: 125%; font-weight:100;")), value = T),
                      div(style = "padding: 0px 0px; margin-top:-2em; margin-left:2em",#remove space between rows
                          fluidRow(#column(2),
                            column(12, 
                                   checkboxGroupInput("ganttLocGanttTime", 
                                                      tags$div(tags$p("Select time interval(s):", style="font-size: 100%; font-weight:100; padding: 0px 0px; margin-bottom:-2em")), 
                                                      c("Day"="day", "Week"="week", "Month"="month"),
                                                      selected="day")))),
                      checkboxInput("ganttIPGantt", tags$div(tags$p("Generate IP Gantt chart", style="font-size: 125%; font-weight:100;")), value = T),
                      div(style = "padding: 0px 0px; margin-top:-2em; ; margin-left:2em",#remove space between rows
                          fluidRow(column(10, 
                                          checkboxGroupInput("ganttIPGanttTime", 
                                                             tags$div(tags$p("Select time interval(s):", style="font-size: 100%; font-weight:100; padding: 0px 0px; margin-bottom:-2em")), 
                                                             c("Day"="day", "Week"="week", "Month"="month"),
                                                             selected="week"))))),
               column(2)),
             fluidRow(column(12, align="center",
                             br(),
                             actionButton("clear", "Clear inputs"),
                             br(),
                             br(),
                             actionButton("ganttRun", "Run", style='background-color:royalblue; color:white; padding:20px 40px'))), #https://www.w3schools.com/css/css3_buttons.asp
             fluidRow(column(12, align="center",
                             br(),
                             br(),
                             htmlOutput("ganttMessage"))), #use instead of textOutput so can change font of returning string: https://stackoverflow.com/questions/24049159/change-the-color-and-font-of-text-in-shiny-app
             fluidRow(column(12, align="center",
                             br(),
                             downloadButton("ganttDownloadData", "Download Results"))))))


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
  shinyjs::hide("noTimeDownloadData")
  shinyjs::hide("ganttDownloadData")
  
  rv <- reactiveValues(clLoc = F,
                       clIP = F,
                       clNTtab = F,
                       clNTcust = F)#variables that if true, indicate that clear has been hit but a new table has not been uploaded
                       
  ##run LATTE when action button hit
  observeEvent(input$run, {
    #check for location file (required)
    if(rv$clLoc) {
      loc = NA
    } else {
      loc = readShinyInputFile(input$locTab)
    }
    if(all(is.na(loc))) {
      output$message <- renderText({paste(outputfontsizestart, "No location data; please input a location table", outputfontsizeend, sep="")})
      return(NULL)
    }
    
    #check for IP file (required sometimes)
    if(rv$clIP) {
      ip = NA
    } else {
      ip = readShinyInputFile(input$ipTab)
    }
    if(all(is.na(ip)) & input$linkType == "ipepi") {
      output$message <- renderText({paste(outputfontsizestart, "No IP data; please input an IP table or select epi link instead of IP epi link", outputfontsizeend, sep="")})
      return(NULL)
    }
    if(all(is.na(ip)) & !is.null(input$ipGanttTime)) {
      output$message <- renderText({paste(outputfontsizestart, "No IP data; please input an IP table or deselect the IP Gantt chart options", outputfontsizeend, sep="")})
      return(NULL)
    }
    
    progress <- Progress$new(session, min=-1, max=15)
    on.exit(progress$close())
    progress$set(message = "Running LATTE")
    output$message <- renderText({paste(outputfontsizestart, "Starting analysis", outputfontsizeend, sep="")})
    # ipEpiLink = ifelse(input$linkType == "ipepi", T, F)
    cutoff = ifelse(input$linkType == "ipepi", input$ipepicutoff, input$epicutoff)
    progress$set(value=0)
    outPrefix = paste(tmpdir, input$prefix, sep="")
    res = tryCatch({
      latteres = latteWithOutputs(outPrefix = outPrefix,
                                  loc = loc,
                                  ip = ip, #readShinyInputFile(input$ipTab),
                                  cutoff = cutoff,
                                  ipEpiLink = input$linkType == "ipepi",
                                  removeAfter = !input$removeAfter,
                                  progress = progress,
                                  drawLocGantt = input$locGanttTime,
                                  drawIPGantt = input$ipGanttTime)
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
    updateCheckboxInput(session, "locGantt", value=T)
    updateCheckboxInput(session, "ipGantt", value=T)
    updateCheckboxGroupInput(session, "locGanttTime", selected="day")
    updateCheckboxGroupInput(session, "ipGanttTime", selected="week")
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
    updateCheckboxInput(session, "locGantt", value=T)
    updateCheckboxInput(session, "ipGantt", value=T)
    updateCheckboxGroupInput(session, "locGanttTime", selected="day")
    updateCheckboxGroupInput(session, "ipGanttTime", selected="week")
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
    input$locGantt
    output$message <- renderText({""})
    output$noTimeMessage <- renderText({""})
    shinyjs::hide("downloadData")
    shinyjs::hide("noTimeDownloadData")
  })
  observe({
    input$locGanttTime
    output$message <- renderText({""})
    output$noTimeMessage <- renderText({""})
    shinyjs::hide("downloadData")
    shinyjs::hide("noTimeDownloadData")
  })
  observe({
    input$ipGantt
    output$message <- renderText({""})
    output$noTimeMessage <- renderText({""})
    shinyjs::hide("downloadData")
    shinyjs::hide("noTimeDownloadData")
  })
  observe({
    input$ipGanttTime
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
  
  ##for Gantt chart timing, ensure upper checkbox is selected/unselected based on checkbox group
  observeEvent(input$locGantt, {
    if(input$locGantt) {
      if(is.null(input$locGanttTime)) {
        updateCheckboxGroupInput(session, "locGanttTime", selected="day")
      }
    } else {
      updateCheckboxGroupInput(session, "locGanttTime", selected=character(0))
    }
  })
  observe({
    input$locGanttTime
    if(is.null(input$locGanttTime)) {
      updateCheckboxInput(session, "locGantt", value = F)
    } else {
      updateCheckboxInput(session, "locGantt", value = T)
    }
  })
  observeEvent(input$ipGantt, {
    if(input$ipGantt) {
      if(is.null(input$ipGanttTime)) {
        updateCheckboxGroupInput(session, "ipGanttTime", selected="week")
      }
    } else {
      updateCheckboxGroupInput(session, "ipGanttTime", selected=character(0))
    }
  })
  observe({
    input$ipGanttTime
    if(is.null(input$ipGanttTime)) {
      updateCheckboxInput(session, "ipGantt", value = F)
    } else {
      updateCheckboxInput(session, "ipGantt", value = T)
    }
  })
}

# Run the app ----
shinyApp(ui = ui, server = server)
