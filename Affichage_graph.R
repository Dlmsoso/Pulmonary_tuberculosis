library(readxl)
library(tidyverse)
library(writexl)
library(openxlsx)
library(purrr)

library(ggplot2)
library(gridExtra)


source("Traitement_de_donnees.R")



extraction<- function(){
  extract_effectif(all_files, "Actes_classants", 
                   "Results_GHM/Actes_classants_entre_2018_2023.xlsx")
  extract_effectif(all_files, "Diag_associe", 
                   "Results_GHM/Disgnostics_associes_entre_2018_2023.xlsx")
  extract_effectif(all_files, "Actes_non_classants", 
                   "Results_GHM/Actes_non_classants_entre_2018_2023.xlsx")
  extract_effectif(all_files, "Diag_principal", 
                   "Results_GHM/Disgnostics_principaux_entre_2018_2023.xlsx")
  extract_stats(all_files, "Results_GHM/Stats_entre_2018_2023.xlsx")
}

# 5 - Étudier l'évolution dans le temps (nombres de cas, d'actes)

print_graph_evolution_cas <- function(){
  
  data <- read_excel("Results_GHM/Stats_entre_2018_2023.xlsx", sheet = "Toute la période")
  data <- data %>% select(. , -Moyenne)
  
  data <- as.data.frame(t(data))
  
  
  colnames(data) <- data[1, ]
  data <- data[-1, ]
  data$Year <- rownames(data)
  data$Year <- as.factor(data$Year)
  data$`Pourcentage de Décès` <- as.numeric(as.character(data$`Pourcentage de Décès`))
  data$`Nombre moyen d'actes` <- as.numeric(as.character(data$`Nombre moyen d'actes`))
  data$`Nombre moyen d'actes classants` <- as.numeric(as.character(data$`Nombre moyen d'actes classants`))
  data$`Nombre moyen d'actes non classants` <- as.numeric(as.character(data$`Nombre moyen d'actes non classants`))
  
  ggplot(data = data, aes(x = Year, y = Effectif, fill = `Pourcentage de Décès`)) +
    geom_col() +
    geom_text(aes(label = sprintf("%.2f%%", `Pourcentage de Décès`)), 
              vjust = +2, size = 4) +
    scale_fill_gradient(low = "lightblue", high = "salmon") +
    labs(x = "Année", y = "Effectif", title = "Tendance annuelle de l'effectif des cas de Tuberculose pulmonaire 
    avec un gradient colorimétrique représentant le taux de décès") +
    theme(
      panel.background = element_rect(fill = "white", colour = "white"), # Définir un fond blanc pour le panneau
      plot.background = element_rect(fill = "white", colour = "white")
      )
    theme_minimal()
  
  ggsave("Graph/Effectif/Graph_evolution_effectif_deces.png")
  
}

print_graph_evolution_niveaux <- function(){

  extract_severity_level(all_files, 
                         "Results_GHM/Niveaux_severite_entre_2018_2023.xlsx")
  
  file_path <- "Results_GHM/Niveaux_severite_entre_2018_2023.xlsx"
  
  plots <- list()
  
  for (sheet_name in names(sheets_dates)) {
  
    data <- read_excel(file_path, sheet = sheet_name)
    data <- data %>% 
      select(-Moyennes)
  
    data_long <- pivot_longer(data, cols = -Niveaux, names_to = "Année", values_to = "Pourcentage")
    
    
    p <- ggplot(data_long, aes(x = Année, y = Pourcentage, fill = Niveaux)) +
      geom_bar(stat = "identity", position = "dodge") +
      geom_text(aes(label = sprintf("%0.1f%%", Pourcentage)), position = position_dodge(width = 0.9), vjust = +2, size = 4) +
      scale_fill_brewer(palette = "Pastel1") +
      labs(title = sheet_name, x = "", y = "Pourcentages") +
      theme_minimal() +
      theme(text = element_text(size = 50),  # Augmente la taille de la police globale
            panel.background = element_rect(fill = "white", colour = "white"), # Définir un fond blanc pour le panneau
            plot.background = element_rect(fill = "white", colour = "white"),
            axis.title = element_text(size = 30),  # Augmente la taille de la police des titres des axes
            axis.text = element_text(size = 30),  # Augmente la taille de la police des textes des axes
            legend.title = element_text(size = 30),  # Augmente la taille de la police du titre de la légende
            legend.text = element_text(size = 30))  # Augmente la taille de la police du texte de la légende
    
      
    graph_name <- paste0("Graph/Niveaux/Evolution_Niveaux_", sheet_name, ".png")
    ggsave(graph_name, p, width = 20, height = 20)
    
    if (sheet_name != "Toute la période")
    {
      plots[[sheet_name]] <- p
    }
  }
  

  grid_plot <- grid.arrange(grobs = plots, nrow = 2, ncol = 2)
  
  ggsave("Graph/Niveaux/Evolution_Niveaux.png", grid_plot, width = 20, height = 20)
  
}

find_type<- function(type){
  if (type == "Actes classant")
  {
    file_path <- "Results_GHM/Actes_classants_entre_2018_2023.xlsx"
    output_path_graph <- "Graph/Actes_classants/Evolution_Actes_classants_"
    output_path_graphs <- "Graph/Actes_classants/Evolution_Actes_classants.png"
  }
  if (type == "Actes non classant")
  {
    file_path <- "Results_GHM/Actes_non_classants_entre_2018_2023.xlsx"
    output_path_graph <- "Graph/Actes_non_classants/Evolution_Actes_non_classants_"
    output_path_graphs <- "Graph/Actes_non_classants/Evolution_Actes_non_classants.png"
  }
  if (type == "Diagnostics principaux")
  {
    file_path <- "Results_GHM/Disgnostics_principaux_entre_2018_2023.xlsx"
    output_path_graph <- "Graph/Disgnostics_principaux/Evolution_Disgnostics_principaux_"
    output_path_graphs <- "Graph/Disgnostics_principaux/Evolution_Disgnostics_principaux.png"
  }
  if (type == "Diagnostics associes")
  {
    
    file_path <- "Results_GHM/Disgnostics_associes_entre_2018_2023.xlsx"
    output_path_graph <- "Graph/Disgnostics_associes/Evolution_Disgnostics_associes"
    output_path_graphs <- "Graph/Disgnostics_associes/Evolution_Disgnostics_associes.png"
  }
  if (type == "Diagnostics associes DMS")
  {
    file_path <- "Results/DMS_Disgnostics_associe_entre_2018_2023.xlsx"
    output_path_graph <- "Graph/Disgnostics_associes/Evolution_Disgnostics_associes_DMS"
    output_path_graphs <- "Graph/Disgnostics_associes/Evolution_Disgnostics_associes_DMS.png"
  }
  return(list(file_path, output_path_graph, output_path_graphs))
}

print_graph_evolution_actes <- function(type){
  
  info <- find_type(type)
  file_path <- info[[1]]
  output_path_graph <- info[[2]]
  output_path_graphs <- info[[3]]
  
  plots <- list()
  all_codes <- c()
  first = TRUE
  for (sheet_name in names(sheets_dates)) {
    data <- read_excel(file_path, sheet = sheet_name)
    all_codes <- unique(c(all_codes, data$Code))
 
    if (FALSE){
      #Supp les valeurs trop grandes
      mean <- sum(data$Moyenne, na.rm = TRUE) / nrow(data)
      data <- data %>% 
          filter(Moyenne <=  2 * mean)
    }
    
    data <- data %>% 
      select(-Moyenne, -Description, -Total_effectif) %>% 
      head(5)
    
    palette_pastel <- c(
      '#F6C5AF',
      '#B8E2A8', 
      '#ECD5E3', 
      '#FFDEB4',
      '#D6E5FA', 
      '#E5E8B6',
      '#CAB9E5',
      '#A0E7E5',
      '#FABFE5',
      '#CBE4F9',
      '#FFD1BA',
      '#FAE1DF',
      '#B4F8C8'
    )
    if (length(all_codes) > length(palette_pastel)) {
      palette_pastel <- rep(palette_pastel, length.out = length(all_codes))
    }
    graph_colors <- setNames(palette_pastel[1:length(all_codes)], all_codes)
    
    
  
    data_long <- pivot_longer(data, cols = -Code, names_to = "Année", values_to = "Effectif")
    
    p <- ggplot(data_long, aes(x = Année, y = Effectif, fill = Code)) +
      geom_bar(stat = "identity", position = "dodge") +
      geom_text(aes(label = Effectif), position = position_dodge(width = 0.9), vjust = +2, size = 4) +
      scale_fill_manual(values = graph_colors) +
      labs(title = sheet_name, x = "", y = "Effectifs", fill = type) +
      theme_minimal() +
      theme(text = element_text(size = 50),
            panel.background = element_rect(fill = "white", colour = "white"),
            plot.background = element_rect(fill = "white", colour = "white"),
            axis.title = element_text(size = 30),
            axis.text = element_text(size = 30),
            legend.title = element_text(size = 15),
            legend.text = element_text(size = 30))
    
    graph_name <- paste0(output_path_graph, sheet_name, ".png")
    ggsave(graph_name, p, width = 20, height = 20)
    plots[[sheet_name]] <- p
  }
  
  grid_plot <- grid.arrange(grobs = plots, nrow = 2, ncol = 2)
  ggsave(output_path_graphs, grid_plot, width = 20, height = 20)
}


find_type_DMS<- function(type){
  if (type == "Actes non classant DMS")
  {
    file_path <- "Results/DMS_Actes_non_classant_associe_entre_2018_2023.xlsx"
    output_path_graph <- "Graph/Actes_non_classants/Evolution_Actes_non_classants_DMS"
    output_path_graphs <- "Graph/Actes_non_classants/Evolution_Actes_non_classants_DMS.png"
  }
  if (type == "Diagnostics associes DMS")
  {
    file_path <- "Results/DMS_Disgnostics_associe_entre_2018_2023.xlsx"
    output_path_graph <- "Graph/Disgnostics_associes/Evolution_Disgnostics_associes_DMS_"
    output_path_graphs <- "Graph/Disgnostics_associes/Evolution_Disgnostics_associes_DMS.png"
  }
  return(list(file_path, output_path_graph, output_path_graphs))
}

print_graph_evolution_DMS <- function(type){
  
  info <- find_type_DMS(type)
  file_path <- info[[1]]
  output_path_graph <- info[[2]]
  output_path_graphs <- info[[3]]
  
  plots <- list()
  all_codes <- c()
  first = TRUE
  for (sheet_name in names(sheets_dates)) {
    data <- read_excel(file_path, sheet = sheet_name)
    all_codes <- unique(c(all_codes, data$Code))
    
    if (FALSE){
      #Supp les valeurs trop grandes
      mean <- sum(data$Moyenne, na.rm = TRUE) / nrow(data)
      data <- data %>% 
        filter(Moyenne <=  2 * mean)
    }
    
    data <- data %>% 
      select(-Moyenne, -Description) %>% 
      head(5)

    
    
    palette_pastel <- c(
      '#F6C5AF',
      '#B8E2A8',
      '#ECD5E3', 
      '#FFDEB4', 
      '#D6E5FA', 
      '#E5E8B6',
      '#CAB9E5', 
      '#FECDDC', 
      '#CBE4F9',
      '#FFD1BA', 
      '#FAE1DF',
      '#A0E7E5', 
      '#B4F8C8'   
    )
    
    if (length(all_codes) > length(palette_pastel)) {
      palette_pastel <- rep(palette_pastel, length.out = length(all_codes))
    }
    graph_colors <- setNames(palette_pastel[1:length(all_codes)], all_codes)
    
    
    
    data_long <- pivot_longer(data, cols = -Code, names_to = "Année", values_to = "Effectif")
    
    p <- ggplot(data_long, aes(x = Année, y = Effectif, fill = Code)) +
      geom_bar(stat = "identity", position = "dodge") +
      geom_text(aes(label = Effectif), position = position_dodge(width = 0.9), vjust = +2, size = 4) +
      scale_fill_manual(values = graph_colors) +
      labs(title = sheet_name, x = "", y = "Effectifs", fill = type) +
      theme_minimal() +
      theme(text = element_text(size = 50),
            panel.background = element_rect(fill = "white", colour = "white"),
            plot.background = element_rect(fill = "white", colour = "white"),
            axis.title = element_text(size = 30),
            axis.text = element_text(size = 30),
            legend.title = element_text(size = 15),
            legend.text = element_text(size = 30))
    
    graph_name <- paste0(output_path_graph, sheet_name, ".png")
    ggsave(graph_name, p, width = 20, height = 20)
    
    plots[[sheet_name]] <- p
  }
  
  grid_plot <- grid.arrange(grobs = plots, nrow = 2, ncol = 2)
  ggsave(output_path_graphs, grid_plot, width = 20, height = 20)
}



# Afficher les graphhiques d'évolution:
#   des effectifs
#   des niveaux de sévérité
#   des actes classants
#   des actes non classants
#   des diagnostics principaux
#   des diagnostics assosicés
#   des diagnostics principaux en fonction des DMS les plus longues

print_graph_evolution_cas()
print_graph_evolution_niveaux()
print_graph_evolution_actes("Actes classant")
print_graph_evolution_actes("Actes non classant")
print_graph_evolution_actes("Diagnostics principaux")
print_graph_evolution_actes("Diagnostics associes")

print_graph_evolution_DMS("Diagnostics associes DMS")


extraction()




