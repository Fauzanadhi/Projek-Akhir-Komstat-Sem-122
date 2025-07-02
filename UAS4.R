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

output$download_report <- downloadHandler(
    filename = function() {
      paste("Laporan_Analisis_", Sys.Date(), ".docx", sep = "")
    },
    content = function(file) {
      # Show progress message menggunakan showNotification
      notification_id <- showNotification(
        "🔄 Sedang mempersiapkan laporan Word...", 
        duration = NULL, 
        type = "message"
      )
      
      tryCatch({
        # Create a temporary Rmd file
        tempReport <- file.path(tempdir(), "report.Rmd")
        
        # Write the RMD content to temp file
        writeLines(rmd_content, tempReport)
        
        # Render the document with proper parameters
        rmarkdown::render(
          input = tempReport,
          output_format = rmarkdown::word_document(),
          output_file = file,
          params = list(
            spearman = values$spearman_result,
            cramer = values$cramer_result,
            correlation = values$correlation_matrix
          ),
          envir = new.env(parent = globalenv()),
          quiet = TRUE
        )
        
        # Remove progress notification
        removeNotification(notification_id)
        
        # Show success notification
        showNotification(
          "✅ Laporan Word berhasil diunduh!", 
          duration = 5, 
          type = "message"
        )
        
      }, error = function(e) {
        # Remove progress notification
        removeNotification(notification_id)
        
        # Show error notification
        showNotification(
          paste("❌ Error membuat laporan:", e$message), 
          duration = 10, 
          type = "error"
        )
        
        print(paste("Debug error:", e$message))  # For debugging
      })
    }
  )
