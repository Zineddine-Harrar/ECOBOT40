library(shiny)
library(shinydashboard)
library(readxl)
library(dplyr)
library(lubridate)
library(ggplot2)

# Chemin vers les fichiers
planning_path <- "ECOBOT40.xlsx"
task_report_path <- "TaskReport_GS142-0230-J4P-P000_20240518162446.xlsx"

# Lire les données
planning_data <- read_excel(planning_path)
task_data <- read_excel(task_report_path)

names(task_data)[grepl("Remaining life of consumables", names(task_data))] <- c("Brush (%)", "Filter (%)", "Squeegee (%)")
# Convertir les dates de début des tâches en jours de la semaine et semaines de l'année
task_data$`Task start time` <- ymd_hms(task_data$`Task start time`)
task_data$DayOfWeek <- wday(task_data$`Task start time`, label = TRUE, abbr = FALSE)
task_data$Week <- week(task_data$`Task start time`)

# Convertir les colonnes en numérique
task_data$`Total time (h)` <- as.numeric(task_data$`Total time (h)`)
task_data$`Actual cleaning area(㎡)` <- as.numeric(task_data$`Actual cleaning area(㎡)`)

# Liste des jours de la semaine pour correspondre avec le planning
days_of_week <- c("lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi", "dimanche")

# Nommer les programmes selon les indices
program_names <- c(
  "Coursive_F1_Fusion",
  "Coursives_F2_Fusion",
  "Correspondance_F1_K1",
  "K1_K2_Cage_CorrespondanceF2"
)

# Interface utilisateur
ui <- dashboardPage(
  dashboardHeader(title = tags$div(img(src = "https://atalian.fr/wp-content/uploads/sites/4/2013/05/atalian-logo.png", height = "40px"), "KPI ECOBOT 40")),
  dashboardSidebar(
    sidebarMenu(
      menuItem("KPIs", tabName = "kpis", icon = icon("tachometer-alt")),
      menuItem("Histogrammes", tabName = "histograms", icon = icon("chart-bar"))
    ),
    selectInput("week", "Sélectionnez la semaine :", choices = unique(task_data$Week))
  ),
  dashboardBody(
    tabItems(
      tabItem(tabName = "kpis",
              fluidRow(
                valueBoxOutput("suiviBox"),
                valueBoxOutput("completionBox"),
                valueBoxOutput("areaBox"),
                valueBoxOutput("timeBox")
              ),
              fluidRow(
                box(title = "Planning des Tâches", status = "primary", solidHeader = TRUE, width = 12, tableOutput("planning_table"))
              )
      ),
      tabItem(tabName = "histograms",
              fluidRow(
                box(title = "Horaires Cumulés par Semaine", status = "primary", solidHeader = TRUE, plotOutput("timeHistogram")),
                box(title = "Surface Nettoyée Cumulée par Semaine", status = "primary", solidHeader = TRUE, plotOutput("areaHistogram"))
              )
      )
    )
  )
)

# Serveur
server <- function(input, output) {
  planning_status_reactive <- reactive({
    # Filtrer les tâches effectuées pour la semaine sélectionnée
    week_data <- task_data %>% filter(Week == input$week)
    
    # Initialiser le tableau de planning avec état des tâches
    planning_status <- as.data.frame(matrix(NA, nrow=length(program_names), ncol=length(days_of_week)))
    colnames(planning_status) <- days_of_week
    rownames(planning_status) <- program_names
    
    # Boucle pour chaque jour de la semaine
    for (day in days_of_week) {
      # Filtrer les tâches planifiées pour ce jour
      planned_tasks <- planning_data[[day]]
      
      # Filtrer les tâches effectuées pour ce jour
      executed_tasks <- week_data %>% filter(DayOfWeek == day)
      
      # Comparer les tâches planifiées et effectuées
      for (i in seq_along(planned_tasks)) {
        planned_task <- planned_tasks[i]
        if (planned_task %in% program_names) {
          if (any(executed_tasks$`Cleaning plan` == planned_task, na.rm = TRUE)) {
            planning_status[planned_task, day] <- "fait"
          } else {
            planning_status[planned_task, day] <- "pas fait"
          }
        }
      }
    }
    
    # Ajustement du tableau
    planning_status <- data.frame(Program=rownames(planning_status), planning_status)
    rownames(planning_status) <- NULL
    
    return(planning_status)
  })
  
  output$planning_table <- renderTable({
    planning_status_reactive()
  }, rownames = FALSE)
  
  # Calcul du taux de suivi basé sur le tableau de planning
  output$suiviBox <- renderValueBox({
    planning_status <- planning_status_reactive()
    total_tasks <- nrow(planning_status) * (ncol(planning_status) - 1) # -1 for Program column
    tasks_done <- sum(planning_status == "fait", na.rm = TRUE)
    taux_suivi <- (tasks_done / total_tasks) * 100
    
    valueBox(
      value = paste0(round(taux_suivi, 2), "%"),
      subtitle = "Taux de Suivi",
      icon = icon("check-circle"),
      color = "green"
    )
  })
  
  output$completionBox <- renderValueBox({
    # Filtrer les tâches effectuées pour la semaine sélectionnée
    week_data <- task_data %>% filter(Week == input$week)
    
    # Initialiser les compteurs
    tasks_done <- 0
    tasks_completed_90 <- 0
    
    # Boucle pour chaque jour de la semaine
    for (day in days_of_week) {
      # Filtrer les tâches planifiées pour ce jour
      planned_tasks <- planning_data[[day]]
      
      # Filtrer les tâches effectuées pour ce jour
      executed_tasks <- week_data %>% filter(DayOfWeek == day)
      
      # Comparer les tâches planifiées et effectuées
      for (planned_task in planned_tasks) {
        if (any(executed_tasks$`Cleaning plan` == planned_task, na.rm = TRUE)) {
          tasks_done <- tasks_done + 1
          task_completed <- executed_tasks %>% filter(`Cleaning plan` == planned_task)
          if (any(task_completed$`Task completion (%)` >= 90, na.rm = TRUE)) {
            tasks_completed_90 <- tasks_completed_90 + 1
          }
        }
      }
    }
    
    # Calcul du taux de complétion
    taux_completion <- (tasks_completed_90 / tasks_done) * 100
    
    valueBox(
      value = paste0(round(taux_completion, 2), "%"),
      subtitle = "Taux de Complétion",
      icon = icon("tasks"),
      color = "blue"
    )
  })
  
  output$areaBox <- renderValueBox({
    # Filtrer les tâches effectuées pour la semaine sélectionnée
    week_data <- task_data %>% filter(Week == input$week)
    
    # Calculer la surface nettoyée cumulée
    total_area <- sum(week_data$`Actual cleaning area(㎡)`, na.rm = TRUE)
    
    valueBox(
      value = paste0(round(total_area, 2), " m²"),
      subtitle = "Surface Nettoyée Cumulée",
      icon = icon("broom"),
      color = "orange"
    )
  })
  
  output$timeBox <- renderValueBox({
    # Filtrer les tâches effectuées pour la semaine sélectionnée
    week_data <- task_data %>% filter(Week == input$week)
    
    # Calculer les heures cumulées
    total_time <- sum(week_data$`Total time (h)`, na.rm = TRUE)
    
    valueBox(
      value = paste0(round(total_time, 2), " h"),
      subtitle = "Heures Cumulées",
      icon = icon("clock"),
      color = "purple"
    )
  })
  
  output$timeHistogram <- renderPlot({
    # Calculer les horaires cumulés par semaine
    time_data <- task_data %>%
      group_by(Week) %>%
      summarise(total_time = sum(`Total time (h)`, na.rm = TRUE))
    
    ggplot(time_data, aes(x = Week, y = total_time)) +
      geom_bar(stat = "identity", fill = "steelblue") +
      labs(title = "Horaires Cumulés par Semaine",
           x = "Semaine", y = "Heures Cumulées") +
      theme_minimal()
  })
  
  output$areaHistogram <- renderPlot({
    # Calculer la surface nettoyée cumulée par semaine
    area_data <- task_data %>%
      group_by(Week) %>%
      summarise(total_area = sum(`Actual cleaning area(㎡)`, na.rm = TRUE))
    
    ggplot(area_data, aes(x = Week, y = total_area)) +
      geom_bar(stat = "identity", fill = "darkorange") +
      labs(title = "Surface Nettoyée Cumulée par Semaine", x = "Semaine", y = "Surface Cumulée (㎡)") +
      theme_minimal()
  })
}

# Lancer l'application
shinyApp(ui = ui, server = server)