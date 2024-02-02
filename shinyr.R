# Install required packages if not already installed
# install.packages(c("shiny", "shinydashboard", "readxl"))

library(shiny)
library(shinydashboard)
library(readxl)
library(DT)
library(ggplot2)
# Source the functions
source("D:/ANNISA_APLIKASI GUI30.R")

# Load data and results
result <- execute_all_steps("D:/IGRADE/datareturn.xlsx")
result
# Define UI
ui <- dashboardPage(
  dashboardHeader(title = tags$div(
    style = "padding-top: 1px; padding-left: 1px;width: 200px;white-space: nowrap; text-overflow: ellipsis;",
    "Model Indeks Tunggal"
  )
  ),
  dashboardSidebar(
    sidebarMenu(
      menuItem("Input Data Return", tabName = "tab1"),
      menuItem("Expected return dan RBR", tabName = "tab2"),
      menuItem("Kovarian", tabName = "tab3"),
      menuItem("Nilai alpha dan beta tiap saham", tabName = "tab4"),
      menuItem("ERB dan Ci Saham", tabName = "tab5"),
      menuItem("Saham di Portofolio", tabName = "tab6"),
      menuItem("Uji Normalitas Multivariat", tabName = "tab7"),
      menuItem("Value at Risk", tabName = "tab8"),
      menuItem("Kinerja Portofolio", tabName = "tab9"),
      menuItem("About", tabName = "tab10")
    )
  ),
  dashboardBody(
    tabItems(
      # Tab 1: Input Data Return
      tabItem(tabName = "tab1",
              fluidRow(
                box(
                  title = "Data",
                  width = 12,
                  dataTableOutput("data")
                )
              )
      ),
      
      # Tab 2: Expected return
      tabItem(tabName = "tab2",
              fluidRow(
                box(
                  width = 12,
                  textOutput("RBR")
                ),
                box(
                  title = "Expected return tiap variabel yaitu:",
                  width = 12,
                  tableOutput("Tabel_RBR")
                )
              )
      ),
      
      # Tab 3: Pemenuhan Asumsi
      tabItem(tabName = "tab3",
              fluidRow(
                box(
                  title = "Kovarian antar Residual Saham",
                  width = 12,
                  tableOutput("table_cov_residual_saham")
                ),
                box(
                  title = "Kovarian antara Residual Saham dengan IHSG",
                  width = 12,
                  tableOutput("table_cov_residual_saham_ihsg")
                )
              )
      ),
      
      # Tab 4
      tabItem(tabName = "tab4",
              fluidRow(
                box(
                  title = "Nilai Karakteristik Alpha dan Beta tiap Saham",
                  width = 12,
                  tableOutput("alphabeta")
                )
              )
      ),
      
      # Tab 5
      tabItem(tabName = "tab5",
              fluidRow(
                box(
                  title = "Nilai ERB dan Ci tiap Saham",
                  width = 12,
                  tableOutput("ERB_ci")
                ),
                box(
                  width = 12,
                  textOutput("Ci")
                )
              )
      ),
      
      # Tab 6: Saham di Portofolio
      tabItem(tabName = "tab6",
              fluidRow(
                box(
                  title = "Saham di Portofolio",
                  width = 12,
                  tableOutput("table_saham_portofolio")
                )
              )
      ),
      
      # Tab 7: Uji Normalitas Multivariat
      tabItem(tabName = "tab7",
              fluidRow(
                box(
                  title = "Uji Normalitas Multivariat",
                  width = 12,
                  plotOutput("plot_multivariate_normality"),
                  textOutput("text_multivariate_normality1"),
                  textOutput("text_multivariate_normality2"),
                  textOutput("text_multivariate_normality3"),
                  textOutput("text_multivariate_normality4"),
                  textOutput("text_multivariate_normality5"),
                  textOutput("text_multivariate_normality6")
                )
              )
      ),
      
      # Tab 8: Value at Risk
      tabItem(tabName = "tab8",
              fluidRow(
                box(
                  title = "Value at Risk",
                  width = 12,
                  textOutput("table_value_at_risk1"),
                  textOutput("table_value_at_risk2")
                )
              )
      ),
      
      # Tab 9: Kinerja Portofolio
      tabItem(tabName = "tab9",
              fluidRow(
                box(
                  title = "Kinerja Portofolio",
                  width = 12,
                  textOutput("text_kinerja_portofolio")
                )
              )
      ),
      
      #Tab 10:About
      tabItem(tabName = "tab10",
              fluidRow(
                box(
                  width = 12,
                  align = "Center",
                  tags$h1("Disusun oleh", style = "font-size: 32px;"), # Mengatur ukuran teks
                  tags$p("Nama : Annisa Muharromah", style = "font-size: 26px;"),
                  tags$p("NIM: 24050120130125", style = "font-size: 26px;"),
                  tags$p("Prodi: Statistika", style = "font-size: 26px;")
                )
              )
      )
    )
  )
)

# Define server
server <- function(input, output, session) {
  # Tab 1: Input Data Return
  output$data <- renderDataTable({
    result$step_1$data
  }, options = list(scrollX = TRUE, pageLength = 10))
  
  #Tab 2:
  output$RBR <- renderText({
    step_1 <- read_and_create_variables("D:/IGRADE/datareturn.xlsx")
    paste("Return Bebas Resiko(RBR) sebesar",step_1$Rbr)
  })
  output$Tabel_RBR <- renderTable({
    variabel_tersisa$result_df
  }, options = list(scrollX = TRUE, pageLength = 10),digits = 10)
  
  
  # Tab 3: Pemenuhan Asumsi
  output$table_cov_residual_saham <- renderTable({
    test_covariance_assumption(step_3$residuals_list)
  }, options = list(scrollX = TRUE, pageLength = 10))
  
  output$table_cov_residual_saham_ihsg <- renderTable({
    test_covariance_ihsg
  }, options = list(scrollX = TRUE, pageLength = 10))
  
  # Tab 4
  output$alphabeta <- renderTable({
    alpha_beta_list$df_alpha_beta
  }, options = list(scrollX = TRUE, pageLength = 10),digits = 10)
  
  #Tab 5
  output$ERB_ci <- renderTable({
    sort_variabel_tersisa(nilai_erb_urut$variabel_tersisa, nilai_erb_urut$data_ERB,ai_bi_list$ai_list,ai_bi_list$bi_list, ci_list)
  }, options = list(scrollX = TRUE, pageLength = 10),digits = 10)
  output$Ci <- renderText({
    result_determine_vars <- determine_variables_in_portfolio(nilai_erb_urut$variabel_tersisa, ci_list, erb_list)
    C_bintang <- determine_cut_off_point(ci_list, result_determine_vars$variabel_max_ci)
    paste("Cut Off pointnya yaitu",result_determine_vars$variabel_max_ci,"dengan nilai",C_bintang)
  })
  
  # Tab 6: Saham di Portofolio
  output$table_saham_portofolio <- renderTable({
    result$weights_df
  }, digits = 6)
  
  # Tab 7: Uji Normalitas Multivariat
  output$plot_multivariate_normality <- renderPlot({
    result <- calculate_mahalanobis_test(result_determine_vars$variabel_tersisa, step_1$data)
    par(mfrow = c(1, 1))
    plot(result$cdf_ds, pch = 15, main = "Cumulative Distribution Function", ylim = c(0, 1), col = "blue")
    lines(pchisq(result$dj_kuadrat, result$df), col = "purple")
  })
  output$text_multivariate_normality1 <- renderText({
    paste("H0: Jarak mahalanobis berdistribusi chi kuadrat")
  })
  
  output$text_multivariate_normality2 <- renderText({
    paste("H1: Jarak mahalanobis tidak berdistribusi chi kuadrat")
  })
  output$text_multivariate_normality3 <- renderText({
    result <- calculate_mahalanobis_test(result_determine_vars$variabel_tersisa, step_1$data)
    paste("Statistik Uji:", result$statistikuji)
  })
  output$text_multivariate_normality4 <- renderText({
    result <- calculate_mahalanobis_test(result_determine_vars$variabel_tersisa, step_1$data)
    paste("D_bintang =", round(result$D_bintang, 6))
  })
  output$text_multivariate_normality5 <- renderText({
    result <- calculate_mahalanobis_test(result_determine_vars$variabel_tersisa, step_1$data)
    paste(result$conclusion)
  })
  output$text_multivariate_normality6 <- renderText({
    result <- calculate_mahalanobis_test(result_determine_vars$variabel_tersisa, step_1$data)
    if (result$statistikuji < result$D_bintang) {
      conclusion <- "H0 diterima"
      comparison <- "lebih kecil dari"
    } else {
      conclusion <- "H0 ditolak"
      comparison <- "lebih besar dari"
    }
    
    paste(sprintf("Dengan taraf signifikansi alpha=5%%, diketahui bahwa %s karena D[statistik uji] %s dari Dbintang.", conclusion, comparison))
  })
  
  # Tab 8: Value at Risk
  output$table_value_at_risk1 <-renderText({
    paste("Besar Value At Risk | Simulasi dengan alpha=5% adalah",result$var_1_simulation$VaR) 
  })
  output$table_value_at_risk2 <-renderText({
    paste("Setelah pembangkitan return selama m kali, diketahui besar rata-rata Value at Risk sebesar: ", result$result_step_17$averageX) 
  })
  
  # Tab 9: Kinerja Portofolio
  output$text_kinerja_portofolio <- renderText({
    paste("Dengan Indeks Treynor, diperoleh besar kinerjanya yaitu: ", result$treynor_values$treynor)
  })
}

# Run Shiny app
shinyApp(ui, server)

