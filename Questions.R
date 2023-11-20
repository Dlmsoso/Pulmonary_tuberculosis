source("Traitement_de_donnees.R")
# On cherche à étudier les principaux diagnostiques des groupes homogénes 

# Extraire les comparatifs pour :
# - Diag_principal 
# - Actes_classants
# - Actes_non_classants 
# - Diag_associe
# - Niveaux_severite
# - Stats

# - Diag_principal 
extract_effectif(all_files, "Diag_principal", 
        "Results_GHM/Disgnostics_principaux_entre_2018_2023.xlsx")

# - Actes_classants
extract_effectif(all_files, "Actes_classants", 
        "Results_GHM/Actes_classants_entre_2018_2023.xlsx")

# - Actes_non_classants
extract_effectif(all_files, "Actes_non_classants", 
        "Results_GHM/Actes_non_classants_entre_2018_2023.xlsx")

# - Diag_associe
extract_effectif(all_files, "Diag_associe", 
        "Results_GHM/Disgnostics_associes_entre_2018_2023.xlsx")

# - Niveaux_severite
extract_severity_level(all_files, 
                       "Results_GHM/Niveaux_severite_entre_2018_2023.xlsx")

# - Stats
extract_stats(all_files, "Results_GHM/Stats_entre_2018_2023.xlsx")

#Pour un groupe de maladies et d'actes associés
#1 - Représenter les hiérarchies correspondantes dans les principales terminologies
#2 - Identifier les diagnostics principaux les plus fréquents du GHM
#3 - Identifier les actes les plus fréquents du GHM
#4 - Identifier les actes et diagnostics associés aux DMS les plus longues
#5 - Etudier révolution dans le temps (nombres de cas, d'actes)
#6 - Identifier des différences entre l'activité publique et privée, par régions,


# 1- Représenter les hiérarchies correspondantes dans les principales terminologies

# 2 - Identifier les diagnostics principaux les plus fréquents du GHM
extract_effectif(all_files, "Diag_principal", 
        "Results/Disgnostics_principaux_entre_2018_2023.xlsx")

# 3 - Identifier les actes les plus fréquents du GHM
# Actes classants et non classants les plus fréquents
extract_effectif(all_files, "Actes_non_classants", 
        "Results/Actes_non_classants_entre_2018_2023.xlsx")
extract_effectif(all_files, "Actes_classants", 
        "Results/Actes_classants_entre_2018_2023.xlsx")

# 4 - Identifier les actes et diagnostics associés aux DMS les plus longues

extract_DMS(all_files, "Diag_associe", 
            "Results/DMS_Disgnostics_associe_entre_2018_2023.xlsx")


extract_DMS(all_files, "Actes_non_classants", 
            "Results/DMS_Actes_non_classant_associe_entre_2018_2023.xlsx")

# 5 - Étudier l'évolution dans le temps (nombres de cas, d'actes)

# 6 - diagnostics associés liés aux Covid 
extract_diagnostic_associe_covid()
# rajotuer pour chaque sheet l'effectif total et une colonne pourcentage !
# Faire la meme chose pour les actes classants !


