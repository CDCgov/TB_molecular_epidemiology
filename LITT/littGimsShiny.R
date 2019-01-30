library(shiny)
library(shinyjs)

gimsVars = c("HOMELESS", "HIVSTAT", "CORRINST", "LONGTERM", "IDU", "NONIDU", "ALCOHOL", 
             "OCCUHCW", "OCCUCORR",
             "RISKTNF", "RISKORGAN", "RISKDIAB", "RISKRENAL", "RISKIMMUNO")

# Define UI ----
ui <- fluidPage(
  titlePanel("LITT with TB GIMS"),
  useShinyjs(),
  fluidRow(
    column(3, align="center",
           h3("Set up")),
    column(3, align="center",
           h3("Input files")),
    column(1),
    column(5, align="center",
           h3("Advanced Options"),
           h4("GIMS risk factors to include with weights"))
  ),
  fluidRow(
    column(3,
           textInput("prefix", "Name prefix for output files"),
           sliderInput("snpCutoff", "SNP cutoff", min=0, max=10, value=5, step=1, round=T),
           checkboxInput("incDateTab", "Check if to output dates used in IP calculations"),
           br(),
           h3("Included cases"),
           helpText("Upload or type state case numbers to analyze. ",
                    "This is optional and if not provided, LITT will analyze all cases in the inputs."),
           fileInput("caseListFile", "Upload list of state case numbers to analyze"),
           textAreaInput("caseListManual", "Type list of state case numbers", rows=5)),
    column(3,
           # h3("Input files"),
           fileInput("distMatrix", "SNP distance matrix"),
           checkboxInput("BNdist",
                         "Check if this contains accession numbers that must be converted to state case number", value=F),
           br(),
           fileInput("epi", "Epi link list")),
           fileInput("caseData", "Case data"),
           # br(),
           # br(),
           # actionButton("clear", "Clear inputs"),
           # br(),
           # br(),
           # actionButton("run", "Run", style='background-color:royalblue; color:white; padding:20px 40px')),
    # helpText("Case data includes, if available, columns for symptom onset, ",
    #          "infectious period start and end, and risk factors."),
    column(1),
    column(width=5, #https://stackoverflow.com/questions/37472915/r-shiny-how-to-generate-this-layout-with-nested-rows-in-column-2
           fluidRow(
             column(4, align="center",
                    lapply(gimsVars[1:7], function(g) {
                      numericInput(g, g, 0, min = 0, step = 1, width="70px")
                      # sliderInput(g, g, min=0, max=3, value=0, step=1, round=T)
                    })),
             column(4, align="center", 
                    # br(),
                    lapply(gimsVars[8:14], function(g) {
                      numericInput(g, g, 0, min = 1, step = 1, width="70px")
                    })),
             column(4, 
                    helpText("Weight > 0 means RF included in ranking. Weight = 0 means RF not in ranking but in outputs. Weight < 0 means RF not in ranking or outputs.", 
                             align="left"))),
           fluidRow(
             column(12, 
                    # helpText("Weight > 0 means RF included in ranking. Weight = 0 means RF not in ranking but in outputs. Weight < 0 means RF not in ranking or outputs.", 
                    #          align="left"),
                    br(),
                    h4("Additional Risk Factors", align="center"),
                    fileInput("rfTable", "Table of risk factor weights", accept=c(".xlsx", ".csv")))))),
                    # helpText("This table contains a list of the columns in the case data table to use a risk factors, with their weights.",
                    #          "Variable names must exactly match the name of the column in the case data table."))))),
    fluidRow(),
    fluidRow(column(12, align="center",
                    actionButton("clear", "Clear inputs"),
                    br(),
                    br(),
                    actionButton("run", "Run", style='background-color:royalblue; color:white; padding:20px 40px'))), #https://www.w3schools.com/css/css3_buttons.asp
    fluidRow(column(12, align="center",
                    br(),
                    br(),
                    htmlOutput("message"))),
    fluidRow(column(12, align="center",
                    br(),
                    downloadButton("downloadData", "Download Results")))
  )

# Define server logic ----
server <- function(input, output, session) {
  output$message <- renderText({"test"})
}

# Run the app ----
shinyApp(ui = ui, server = server)