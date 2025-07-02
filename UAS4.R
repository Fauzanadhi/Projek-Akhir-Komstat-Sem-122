# Load library
library(shiny)
library(readr)
library(readxl)
library(dplyr)
library(ggplot2)
library(plotly)
library(rstatix)
library(bslib)
library(DT)
library(reshape2)
library(shinyFeedback)
library(rmarkdown)
library(knitr)
library(tools)
library(tinytex)

# UI
ui <- fluidPage(
  useShinyFeedback(),
  theme = bs_theme(bootswatch = "superhero"),
  titlePanel("Uji Korelasi Spearman dan Asosiasi Cramér's V"),
  p("Aplikasi ini digunakan untuk mengevaluasi hubungan antara dua variabel numerik (Spearman) atau dua variabel kategorik (Cramér's V). Unggah data Anda dan pilih variabel yang ingin diuji."),
  sidebarLayout(
    sidebarPanel(
      fileInput("file", "Unggah File CSV / Excel", accept = c(".csv", ".xlsx")),
      h4("Uji Korelasi Spearman (Numerik vs Numerik)"),
      uiOutput("spearman_ui"),
      h4("Uji Asosiasi Cramér's V (Kategori vs Kategori)"),
      uiOutput("cramer_ui"),
      actionButton("analyze", label = tagList(icon("play"), "Lakukan Analisis"), class = "btn btn-success"),
      br(), br(),
      uiOutput("download_ui")
    ),
    mainPanel(
      tabsetPanel(
        tabPanel("Preview Data", DT::dataTableOutput("preview")),
        tabPanel("Ringkasan Data", verbatimTextOutput("summary")),
        tabPanel("Spearman",
                 verbatimTextOutput("spearman_result"),
                 plotlyOutput("spearman_plot"),
                 uiOutput("spearman_interpretation"),
                 uiOutput("spearman_detail")),
        tabPanel("Cramér's V",
                 verbatimTextOutput("cramer_result"),
                 plotOutput("cramer_plot"),
                 uiOutput("cramer_interpretation"),
                 uiOutput("cramer_detail")),
        tabPanel("Heatmap Korelasi", 
                 plotOutput("correlation_heatmap"),
                 uiOutput("heatmap_description"))
      )
    )
  )
)

```{r correlation-matrix}
if (!is.null(params$correlation)) {
  kable(params$correlation, 
        caption = "Matrix Korelasi Spearman untuk semua variabel numerik",
        digits = 3,
        align = "c")
} else {
  cat("Matrix korelasi tidak tersedia\\n")
}
```
