library(readxl)
library(tidyverse)
library(writexl)
library(openxlsx)
library(purrr)

#CIM-10

#install.packages("remotes")
#library("remotes")
#remotes::install_github("denisGustin/refpmsi")
library("refpmsi")

#Traitement pour MCO par GHM ou racine
# Pour le GHM : 04M19 
# On a : 

names <- c("Toute la période", "Période avant Covid", 
           "Période pendant Covid", "Période après Covid")
sheets_dates <- list(
  "Toute la période" = c("2018", "2019", "2020", "2021", "2022"),
  "Période avant Covid" = c("2018", "2019"),
  "Période pendant Covid" = c("2020", "2021"),
  "Période après Covid" = c("2022", "2023")
)

all_files <- list(
  "Toute la période" = c(
    "MCO_par_GHM_ou_racine/04M19_2018.xls",
    "MCO_par_GHM_ou_racine/04M19_2019.xls",
    "MCO_par_GHM_ou_racine/04M19_2020.xls",
    "MCO_par_GHM_ou_racine/04M19_2021.xls",
    "MCO_par_GHM_ou_racine/04M19_2022.xls",
    "MCO_par_GHM_ou_racine/04M19_2023.xls"
  ),
  "Période avant Covid" = c(
    "MCO_par_GHM_ou_racine/04M19_2018.xls",
    "MCO_par_GHM_ou_racine/04M19_2019.xls"
  ),
  "Période pendant Covid" = c(
    "MCO_par_GHM_ou_racine/04M19_2020.xls",
    "MCO_par_GHM_ou_racine/04M19_2021.xls"
  ),
  "Période après Covid" = c(
    "MCO_par_GHM_ou_racine/04M19_2022.xls",
    "MCO_par_GHM_ou_racine/04M19_2023.xls"
  )
)

ccam_list = "liste_code_ccam.xlsx"

sheets <- list("Infos", 
               "Actes_non_classants", 
               "Actes_classants", 
               "Diag_associe", 
               "Diag_principal", 
               "Niveaux_severite", 
               "Stats")


# Fonctions Communes à toutes 
get_period_name <- function(name) {
  if (name == "files_GHM") {
    return("Toute la période")
  } else if (name == "files_GHM_avant_covid") {
    return("Période avant Covid")
  } else if (name == "files_GHM_pendant_covid") {
    return("Période pendant Covid")
  } else if (name == "files_GHM_apres_covid") {
    return("Période après Covid")
  }
  return(name)
}

add_code_ccam <- function(data){
  ccam_data <- read_excel("liste_code_ccam.xlsx", 
                          sheet = "CCAM pour usage PMSI",
                          range = cell_cols(c("A", "C")))
  
  ccam_data <- ccam_data %>%
    rename(Code = `Subdivision / Code`) %>%
    rename(Description = `Texte`) %>%
    select(Code, Description)
  
  filtered_ccam_data <- ccam_data %>%
    filter(Code %in% data$Code)
  
  merged_data <- merge(data, filtered_ccam_data, by = "Code") 
  
  
  return(merged_data)
}

add_code_cim10 <- function(data, period){
  date = 2022
  if (period == "Période avant Covid"){
    date = 2019
  }
  if (period == "Période pendant Covid"){
    date = 2020
  }
  cim_2018_2022 <- refpmsi::refpmsi("cim",date)
  cim_2018_2022 <- cim_2018_2022 %>% 
    select(`cim_code`, `cim_lib`)
  
  cim_2018_2022 <- cim_2018_2022 %>%
    rename(Code = `cim_code`) %>%
    rename(Description = `cim_lib`) %>%
    select(Code, Description)
  
  filtered_ccam_data <- cim_2018_2022 %>%
    filter(Code %in% data$Code)
  
  merged_data <- merge(data, filtered_ccam_data, by = "Code") 

  return(merged_data)
}

extract_effectif <- function(data_files_list, sheet_name, output_file_path) {
  
  load_data <- function(file_path) {
    data <- read_excel(file_path, sheet = sheet_name)
    data$Effectif <- gsub(" ", "", data$Effectif)
    data$Effectif <- as.numeric(as.character(data$Effectif))
    return(data)
  }
  
  summarize_data <- function(files_list) {
    # Charger et préparer les données
    loaded_data <- map(files_list, ~load_data(.x))
    
    # Créer une liste de data frames avec une colonne pour 
    # l'effectif de chaque fichier
    effectifs_par_fichier <- map(loaded_data, ~.x %>% 
                                   group_by(Code) %>% 
                                   summarise(Effectif = sum(Effectif, na.rm = TRUE)))
      
    
    
    # Créer un data frame final en fusionnant toutes les données
    data_final <- reduce(effectifs_par_fichier, full_join, by = "Code")
    
    
    # Ajouter une colonne pour l'effectif total
    data_final <- data_final %>% 
      mutate(Total_Effectif = rowSums(
        select(., starts_with("Effectif")), na.rm = TRUE),
        # Calculer la moyenne des effectifs pour chaque code
        Moyenne_Effectif = Total_Effectif / length(files_list)) %>%
        arrange(desc(Total_Effectif))
    
    return(data_final)
  }
  
  rename_col <- function(data, period){
    if (period == "Toute la période") {
      col_names <- c("Code", "2018", "2019", "2020", 
                     "2021", "2022", "2023", "Total_effectif", "Moyenne")
    } else if (period == "Période avant Covid") {
      col_names <- c("Code", "2018", "2019", "Total_effectif", "Moyenne")
    } else if (period == "Période pendant Covid") {
      col_names <- c("Code", "2020", "2021", "Total_effectif", "Moyenne")
    } else if (period == "Période après Covid") {
      col_names <- c("Code", "2022", "2023", "Total_effectif", "Moyenne")
    } else {
      stop("Période inconnue")
    }
    
    colnames(data) <- col_names

    return(data)
  }
  
  
  
  excel_file <- createWorkbook()
  
  
  
  for (i in 1:4) {
    data <- summarize_data(data_files_list[[i]])
    data <- rename_col(data, names[i])
    
    if (grepl("Actes", sheet_name)) {
      data <- add_code_ccam(data)
    }
    if (grepl("Diag", sheet_name)) {
      data <- add_code_cim10(data, names[i])
    }
    
    data[is.na(data)] <- 0
    #View(data)
    
    data <- data %>%
      arrange(desc(Total_effectif))
    
    addWorksheet(excel_file, names[i])
    writeData(excel_file, names[i], data)
  }
  
  saveWorkbook(excel_file, output_file_path, overwrite = TRUE)
  
  return(invisible(NULL))
}

extract_severity_level <- function(files_list, output_file) {
  
  read_and_process_file <- function(file_path) {
    data <- read_excel(file_path, sheet = "Niveaux_severite")
    data <- data %>% mutate(across(everything(), ~ readr::parse_number(as.character(.)) * 100))
    return(data)
  }
  
  excel_file <- createWorkbook()
  
  for (name in names(files_list)) {
    data <- lapply(files_list[[name]], read_and_process_file) %>%
      bind_rows(.id = "Year") %>%
      mutate(Year = case_when(
        Year == "1" ~ ifelse(name == "Toute la période" | name == "Période avant Covid", "2018", 
                             ifelse(name == "Période après Covid", "2022", "2020")),
        Year == "2" ~ ifelse(name == "Toute la période" | name == "Période avant Covid", "2019",
                             ifelse(name == "Période après Covid", "2023", "2020")),
        Year == "3" ~ "2020",
        Year == "4" ~ "2021",
        Year == "5" ~ "2022",
        Year == "6" ~ "2023",
        TRUE ~ Year
      )) %>%
      
      pivot_longer(cols = -Year, names_to = "Niveaux", values_to = "Percentages") %>%
      pivot_wider(names_from = Year, values_from = Percentages, values_fn = mean ,names_sort = TRUE) %>%
      mutate(
        Moyennes = rowMeans(select(., where(is.numeric), -Niveaux), na.rm = TRUE)
      ) %>% 
      arrange(desc(Moyennes))
    
    sheet_name <- get_period_name(name)
    addWorksheet(excel_file, sheet_name)
    writeData(excel_file, sheet_name, data)
  }
  saveWorkbook(excel_file, output_file, overwrite = TRUE)
}


extract_stats <- function(files_list, output_file) {

  read_and_process_stats <- function(file_path) {
    data <- read_excel(file_path, sheet = "Stats", col_names = FALSE)
    colnames(data) <- c("Statistique", "Valeurs")
    data$Valeurs <- gsub(" ", "", data$Valeurs)
    data$Valeurs <- as.numeric(as.character(data$Valeurs))
    return(data)
  }
  
  excel_file <- createWorkbook()
  
  for (name in names(files_list)) {
    stats_data <- lapply(files_list[[name]], read_and_process_stats) %>%
      bind_rows(.id = "Year") %>%
      mutate(Year = case_when(
        Year == "1" ~ ifelse(name == "Toute la période" | name == "Période avant Covid", "2018", 
                             ifelse(name == "Période après Covid", "2022", "2020")),
        Year == "2" ~ ifelse(name == "Toute la période" | name == "Période avant Covid", "2019",
                             ifelse(name == "Période après Covid", "2023", "2020")),
        Year == "3" ~ "2020",
        Year == "4" ~ "2021",
        Year == "5" ~ "2022",
        Year == "6" ~ "2023",
        TRUE ~ Year
      )) %>%
      pivot_wider(names_from = Year, values_from = Valeurs, values_fn = mean) %>%
      mutate(
        Moyenne = rowMeans(select(., where(is.numeric), -Statistique), na.rm = TRUE)
      )
      #mutate(Moyenne = rowMeans(select(., -Statistique), na.rm = TRUE))
    
    sheet_name <- get_period_name(name)
    addWorksheet(excel_file, sheet_name)
    writeData(excel_file, sheet_name, stats_data)
  }
  
  saveWorkbook(excel_file, output_file, overwrite = TRUE)
}

extract_DMS <- function(data_files_list, sheet_name, output_file_path) {
  
  load_data <- function(file_path) {
    data <- read_excel(file_path, sheet = sheet_name)
    data$DMS <- gsub(" ", "", data$DMS)
    data$DMS <- as.numeric(as.character(data$DMS))
    return(data)
  }
  
  summarize_data <- function(files_list) {
    loaded_data <- map(files_list, ~load_data(.x))
    
    DMS_par_fichier <- map(loaded_data, ~.x %>% 
                             group_by(Code) %>% 
                             summarise(DMS = sum(DMS, na.rm = TRUE)))
    
    data_final <- reduce(DMS_par_fichier, full_join, by = "Code")
    
    data_final <- data_final %>% 
      mutate(Total_DMS = rowSums(
        select(., starts_with("DMS")), na.rm = TRUE),
        Moyenne_DMS = Total_DMS / length(files_list)) %>%
      select(., -Total_DMS) %>% 
      arrange(desc(Moyenne_DMS))
    
    return(data_final)
  }
  
  rename_col <- function(data, period){
    if (period == "Toute la période") {
      col_names <- c("Code", "2018", "2019", "2020", 
                     "2021", "2022", "2023", "Moyenne")
    } else if (period == "Période avant Covid") {
      col_names <- c("Code", "2018", "2019", "Moyenne")
    } else if (period == "Période pendant Covid") {
      col_names <- c("Code", "2020", "2021", "Moyenne")
    } else if (period == "Période après Covid") {
      col_names <- c("Code", "2022", "2023", "Moyenne")
    } else {
      stop("Période inconnue")
    }
    
    
    
    colnames(data) <- col_names
    
    return(data)
  }
  
  excel_file <- createWorkbook()
  
  #names <- c("Toute la période", "Période avant Covid", "Période pendant Covid", "Période après Covid")
  for (i in 1:4) {
    data <- summarize_data(data_files_list[[i]])
    data <- rename_col(data, names[i])
    if (grepl("Actes", sheet_name)) {
      data <- add_code_ccam(data)
    }
    if (grepl("Diag", sheet_name)) {
      data <- add_code_cim10(data, names[i])
    }
    
    data <- data %>%
      arrange(desc(Moyenne))
    
    addWorksheet(excel_file, names[i])
    writeData(excel_file, names[i], data)
  }
  
  saveWorkbook(excel_file, output_file_path, overwrite = TRUE)
  
  return(invisible(NULL))
}

extract_diagnostic_associe_covid <- function() {
  data <- read_excel("Results/DMS_Disgnostics_associe_entre_2018_2023.xlsx")
  data <- data %>%
    filter(grepl("COVID", Description, ignore.case = TRUE)) %>% 
    select(., -"2018", -"2019")
  write.xlsx(data, "Results/Disgnostics_associes_avec_lien_Covid_entre_2018_2023.xlsx")
}





