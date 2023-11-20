library(terra)
library(geodata)
library(sf)
library(sp)
library(rnaturalearth)
library(devtools)
library(rnaturalearthhires)
library(readxl)
library(tidyverse)
library(writexl)
library(openxlsx)
library(purrr)

library(viridisLite)
library(viridis)
library(RColorBrewer)
library(ggplot2)
library(magick)



cartographie <- function(input_file, output_file, excel_file ) {
  year <- substr(output_file, nchar(output_file)-3, nchar(output_file))
  
  population <- read_excel("Cartographie/population_2018_2023.xlsx", sheet = "Infos")
  dates <- c("2018", "2019", "2020", "2021", "2022", "2023")
  for (date in dates) {
    population[[date]] <- as.numeric(gsub(" ", "", as.character(population[[date]])))
  }
  
  population <- population %>% 
    select("Régions", !!year)
  
  population <- population %>% 
    rename(pop = !!year)
  
  population$pop <- as.numeric(gsub(" ", "", as.character(population$pop)))

  data <- read_excel(input_file, sheet = "Infos")
  data <- as.data.frame(t(data))
  colnames(data) <- data[1, ]
  data <- data[-1, ]
  data$Effectif <- as.numeric(gsub(" ", "", as.character(data$Effectif)))
  
  data <- rownames_to_column(data, var = "Régions")
  data <- data %>% select(Régions, Effectif)
  
  
  data <- merge(data, population[, c("Régions", "pop")], by = "Régions", all.x = TRUE)
  
  
  FranceRegions <- ne_states(country = "France", returnclass = "sp")
  
  idx <- match(FranceRegions$region, data$Régions)
  population_idx <- match(FranceRegions$region, population$Régions)
  
  FranceRegions$Effectif <- data$Effectif[idx]
  FranceRegions$Population <- population$pop[population_idx]

  
  
  
  FranceRegions$Proportion <- ifelse(FranceRegions$Population == 0, 
                                     0, 
                                     (FranceRegions$Effectif / FranceRegions$Population) * 100000)
  
  FranceRegions$Proportion <- ifelse(is.na(FranceRegions$Proportion), 0, FranceRegions$Proportion)
  
  max_val <- max(FranceRegions$Proportion, na.rm = TRUE) + 1
  min_val <- min(0)
  
  
  colorkey_at <- round(seq(min_val, max_val, length.out = 100))
  colorkey_labels <- colorkey_at
  colorkey_labels <- as.character(colorkey_labels)
  
  breaks <- c(0, quantile(FranceRegions$Proportion, probs = seq(0, 1, 0.25), na.rm = TRUE), max_val)
  breaks <- round(breaks, 1)
  
  
  couleurs <- brewer.pal(name = "YlOrRd", n = 9)
  
  colors <- sapply(FranceRegions$Proportion, function(x) {
    if (!is.na(x) && x > 0) {
      return(couleurs[ceiling(x / max(FranceRegions$Proportion, na.rm = TRUE) * length(couleurs))])
    } else if (x == 0  || is.na(x)) {
      return("gray")
    } else {
      return("white")
    }
  })
  
  selection <- subset(FranceRegions, (region %in% c("Île-de-France", "Mayotte", "Guyane française")))

  selection <- selection[, c("region", "Proportion")]
  selection <- selection[!duplicated(selection$region), ]
  
  
  
  addWorksheet(excel_file, year)
  writeData(excel_file, year , selection)
  
  # Pour Mayotte et la Réunion
  #FranceRegions <- subset(FranceRegions, (region %in% c("Réunion", "Mayotte")))
  
  # Pour La guyane, Guadeloupe et Martinique
  #FranceRegions <- subset(FranceRegions, (region %in% c("Guyane française", "Guadeloupe", "Martinique")))
  
  # Pour la France
  FranceRegions <- subset(FranceRegions, !(region %in% c("Guyane française", "Guadeloupe", "Martinique", "Réunion", "Mayotte")))

  
  
  output_filename <- paste0(output_file, ".png")
  
  png(output_filename, width = 800, height = 600)

  
  plot <- spplot(FranceRegions, "Proportion", col.regions = couleurs, 
         at = breaks,
         main = list(label = paste("Proportion de cas de Tuberculose pumlonaire pour 100,000 habitants en France en", year), cex = .8),
         colorkey = list(space = "right", at = breaks, labels = as.character(breaks), col = couleurs))
  
  print(plot)
  on.exit(dev.off())  
  }

combine_all_images_in_folder <- function(folder_path, output_path, horizontal = TRUE) {
  
  image_files <- list.files(path = folder_path, pattern = "\\.png$", full.names = TRUE)
  
  if (length(image_files) == 0) {
    stop("Aucun fichier image trouvé dans le dossier spécifié.")
  }
  
  images <- lapply(image_files, image_read)
  
  combined_image <- if(horizontal) {
    image_append(image_join(images))
  } else {
    image_montage(image_join(images), tile = "1x")
  }
  
  image_write(combined_image, output_path)
  

}

effectuer_cartographies <- function(base_path_xlsx, output_dir, annees = 2018:2023) {
  
  excel_file <- createWorkbook()
  
  for (annee in annees) {
    
    
    
    fichier_xlsx <- sprintf("%s/04M19_%d_regions.xlsx", base_path_xlsx, annee)
    fichier_sortie <- sprintf("%s/carte_tuberculose_regions_%d", output_dir, annee)
  
    cartographie(fichier_xlsx, fichier_sortie, excel_file)
    
  }
  combine_all_images_in_folder("Cartographie/com/", paste("Cartographie/com/","combined.png"))
  combine_all_images_in_folder("Cartographie/com2/", paste("Cartographie/com2/","combined.png"))
  
  saveWorkbook(excel_file , "Cartographie/chiffres_tuberculose_region_2018_2023.xlsx", overwrite = TRUE)
  
  
}



# Cartographie des cartes par Régions en France avec une combinaison d'image

effectuer_cartographies("Cartographie/", "Cartographie/France/")

