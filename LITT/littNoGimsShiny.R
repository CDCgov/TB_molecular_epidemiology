library(shiny)

# Define UI ----
ui <- fluidPage(
  titlePanel("LITT"),
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
           # helpText("Must include: sputum smear results, cavitation status, ",
           #          "whether a case is extrapulmonary only or pediatric, and ",
           #          "infectious period start and end.", 
           #          "Any additional columns will be treated as risk factors.",
           #          "See documentation for more details."),
           fileInput("epi", "Epi link table"),
           fileInput("distMatrix", "SNP distance matrix"),
           checkboxInput("distIsTri", 
                         "Check if only the lower triangle is present in distance matrix (e.g. a table from BioNumerics)"), value=F),
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

# Define server logic ----
server <- function(input, output) {
  
}

# Run the app ----
shinyApp(ui = ui, server = server)