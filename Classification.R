rm(list=ls())
setwd("C:/Users/kaffa/OneDrive/Bureau/R")

###############################################
########## Importation des packages ###########
###############################################

source('Cours ADD & Classification R - Packages.R')


dictionnaire_mkt <- read_excel("C:/Users/kaffa/OneDrive/Bureau/R/Dictionnaire de donn??es.xlsx", range = 'B2:C72', col_names = TRUE)
print(dictionnaire_mkt, n = 27)
df_mkt = read_excel("C:/Users/kaffa/OneDrive/Bureau/R/BDD_exam.xlsx", sheet=1)

## Premier apergu de la base
head(df_mkt)
introduce(df_mkt) # Nombre de lignes, colonnes, types de variables...
plot_intro(df_mkt) # Reprisentation graphique du descriptif de la base

## Focus valeurs manquantes
plot_missing(df_mkt) # Valeurs manquantes sur la variable income, 1.07%, soit 24 individus

## Analyse univariie : Reprisentation graphique des variables 1 ` 1
plot_bar(df_mkt) # qualitatives (3 variables)
plot_histogram(df_mkt) # quantitatives (26 variables)

## Criation d'un reporting pour un apergu rapide du comportement de nos donnies
create_report(
  data = df_mkt,
  output_format = html_document(toc = TRUE, toc_depth = 6, theme = "flatly"),
  output_file = "report",
  output_dir = paste0(getwd(), '/Sorties'),
  config = configure_report(
    add_plot_prcomp = TRUE,
    plot_bar_args = list("by_position" = "dodge"),
    plot_correlation_args = list("type" = "continuous", "cor_args" = list("use" = "all.obs"), "theme_config" = list("axis.text.x" = element_text("angle" = 90))),
    plot_prcomp_args = list("nrow" = 1, "ncol" = 3),
    plot_str_args = list(type = "diagonal", fontSize = 35, width = 1000, margin =
                           list(left = 350, right = 250)),
    global_ggtheme = quote(theme_linedraw(base_size = 16L)),
  )
)

##########################################################################
# Rapport apres suppression des variables avec des valeurs manquantes >5%#
##########################################################################

# Liste des variables ?? supprimer
vars_to_remove <- c(
  "emp_title", "total_rev_hi_lim", "tot_cur_bal", "tot_coll_amt", 
  "mths_since_last_delinq", "desc", "mths_since_last_major_derog", 
  "mths_since_last_record", "inq_last_12m", "total_cu_tl", "inq_fi", 
  "all_util", "max_bal_bc", "open_rv_24m", "open_rv_12m", "il_util", 
  "total_bal_il", "mths_since_rcnt_il", "open_il_24m", "open_il_12m", 
  "open_il_6m", "open_acc_6m", "verification_status_joint", "dti_joint", 
  "annual_inc_joint", "next_pymnt_d")

# Suppression des variables
df_mkt <- df_mkt %>% select(-all_of(vars_to_remove))


## Premier apergu de la base
head(df_mkt)
introduce(df_mkt) # Nombre de lignes, colonnes, types de variables...
plot_intro(df_mkt) # Reprisentation graphique du descriptif de la base

## Focus valeurs manquantes
plot_missing(df_mkt) # Valeurs manquantes sur la variable income, 1.07%, soit 24 individus

## Analyse univariie : Reprisentation graphique des variables 1 ` 1
plot_bar(df_mkt) # qualitatives (3 variables)
plot_histogram(df_mkt) # quantitatives (26 variables)

## Criation d'un reporting pour un apergu rapide du comportement de nos donnies
create_report(
  data = df_mkt,
  output_format = html_document(toc = TRUE, toc_depth = 6, theme = "flatly"),
  output_file = "report2",
  output_dir = paste0(getwd(), '/Sorties'),
  config = configure_report(
    add_plot_prcomp = TRUE,
    plot_bar_args = list("by_position" = "dodge"),
    plot_correlation_args = list("type" = "continuous", "cor_args" = list("use" = "all.obs"), "theme_config" = list("axis.text.x" = element_text("angle" = 90))),
    plot_prcomp_args = list("nrow" = 1, "ncol" = 3),
    plot_str_args = list(type = "diagonal", fontSize = 35, width = 1000, margin =
                           list(left = 350, right = 250)),
    global_ggtheme = quote(theme_linedraw(base_size = 16L)),
  )
)
##########################################################################################
########################## Suprression des valeurs manquantes ############################
##########################################################################################

library(tidyr)

# Variables avec valeurs manquantes
vars_with_na <- c("title", "last_credit_pull_d", "collections_12_mths_ex_med", "revol_util", "emp_length")

# Supprimer les lignes manquantes des variables sp??cifi??es
df_mkt <- df_mkt %>% drop_na(all_of(vars_with_na))

# Afficher le nombre de lignes avant et apr??s le nettoyage
print(paste("Nombre de lignes avant nettoyage :", nrow(df_mkt)))


df <- df_mkt
# Ajouter la variable `y` qui compte les lignes
df <- df %>% mutate(y = row_number())

##########################################################################################
## Exclusion des valeurs extrjmes/aberrantes par prudence et pour soucis de gain de temps#
##########################################################################################

quant_vars <- c("revol_util", "revol_bal", "annual_inc")
outliers <- list(
  revol_util = c(160956),
  revol_bal = c(153986, 90696, 89292, 148703, 87795, 86438),
  annual_inc = c(30813, 22620, 138050,87795,23008,80765,85650)
)

# Fonction pour remplacer les outliers par le quantile 99% en excluant les outliers
replace_outliers_with_quantile <- function(df, var, outliers) {
  # Exclure les outliers
  non_outlier_data <- df[-outliers[[var]], var]
  
  # Calculer le quantile 99% sur les donn??es non-outliers
  quantile_99 <- quantile(non_outlier_data, 0.99, na.rm = TRUE)
  
  # Remplacer les valeurs des outliers par le quantile 99%
  df[outliers[[var]], var] <- quantile_99
  
  return(df)
}
# Remplacer les outliers pour chaque variable sp??cifi??e
for (var in quant_vars) {
  df <- replace_outliers_with_quantile(df, var, outliers)
}


# Num??ros de ligne des valeurs ?? remplacer
outlier_rows <- c(80765, 23008, 87795, 138050, 22620, 30813)

# Remplacer les valeurs des lignes sp??cifiques par 2,000,000
df$annual_inc[outlier_rows] <- 2000000


################################################################################
## Criation d'un reporting pour un apergu rapide du comportement de nos donnies
################################################################################

# Remplacer les valeurs de `revol_util` sup??rieures ?? 100 par 100
df$revol_util[df$revol_util > 100] <- 100


##########################################################################################
########################## ## Traitement les variables ###################################
##########################################################################################

#suppression des variables ?? une seul modalit??s :

vars_to_remove <- c(
  "Recovery_process", "application_type", "recoveries", 
  "collection_recovery_fee", "policy_code","pymnt_plan","url","X","id","member_id"
)
df <- df %>% select(-all_of(vars_to_remove))





## Premier apercu de la base
head(df)
introduce(df) # Nombre de lignes, colonnes, types de variables...
plot_intro(df) # Reprisentation graphique du descriptif de la base

## Focus valeurs manquantes
plot_missing(df) # Valeurs manquantes sur la variable income, 1.07%, soit 24 individus

## Analyse univariie : Reprisentation graphique des variables 1 ` 1
plot_bar(df) # qualitatives (3 variables)
plot_histogram(df) # quantitatives (26 variables)

#reporting
create_report(
  data = df,
  output_format = html_document(toc = TRUE, toc_depth = 6, theme = "flatly"),
  output_file = "report999999",
  output_dir = paste0(getwd(), '/Sorties'),
  config = configure_report(
    add_plot_prcomp = TRUE,
    plot_bar_args = list("by_position" = "dodge"),
    plot_correlation_args = list("type" = "continuous", "cor_args" = list("use" = "all.obs"), "theme_config" = list("axis.text.x" = element_text("angle" = 90))),
    plot_prcomp_args = list("nrow" = 1, "ncol" = 3),
    plot_str_args = list(type = "diagonal", fontSize = 35, width = 1000, margin =
                           list(left = 350, right = 250)),
    global_ggtheme = quote(theme_linedraw(base_size = 16L)),
  )
)


#### 
#### 
#### Regroupement des variables #####
#### 
#### 


### Variable Purpose
df <- df %>%
  mutate(purpose_grouped = case_when(
    purpose %in% c("educational", "vacation", "wedding", "medical", "moving") ~ "consumption_credit",
    purpose %in% c("major_purchase", "car", "house") ~ "goods_purchase",
    purpose %in% c("home_improvement", "renewable_energy") ~ "improvement_maintenance",
    purpose %in% c("debt_consolidation", "credit_card") ~ "financial_management",
    purpose %in% c("small_business", "other") ~ "others",
    TRUE ~ purpose  # Pour g??rer les cas non sp??cifi??s
  ))


## Variable emp_length

df <- df %>%
  mutate(emp_length_grouped = case_when(
    emp_length %in% c("< 1 year", "1 year", "2 years") ~ "less_than_3_years",
    emp_length %in% c("3 years", "4 years", "5 years") ~ "3_to_5_years",
    emp_length %in% c("6 years", "7 years", "8 years", "9 years", "10+ years") ~ "6_to_10_years",
    TRUE ~ emp_length  # Pour g??rer les cas non sp??cifi??s
  ))

### Variable addr_state

df <- df %>%
  mutate(region = case_when(
    addr_state %in% c("AZ", "CA", "CO", "HI", "ID", "MT", "NV", "NM", "OR", "UT", "WA", "WY") ~ "West",
    addr_state %in% c("IA", "IL", "IN", "KS", "MI", "MN", "MO", "NE", "ND", "OH", "SD", "WI") ~ "Midwest",
    addr_state %in% c("AL", "AR", "DC", "DE", "FL", "GA", "KY", "LA", "MD", "MS", "NC", "OK", "SC", "TN", "TX", "VA", "WV") ~ "South",
    addr_state %in% c("CT", "MA", "ME", "NH", "NJ", "NY", "PA", "RI", "VT") ~ "Northeast",
    TRUE ~ "Other"  # Pour g??rer les cas non sp??cifi??s
  ))

## Variable Home owner ship
df <- df %>%
  mutate(home_ownership_grouped = case_when(
    home_ownership %in% c("NONE", "ANY", "OTHER") ~ "OTHER",
    TRUE ~ home_ownership  # Conserve les autres cat??gories telles quelles
  ))

## Variable verification_status 
df <- df %>%
  mutate(verification_status_grouped = case_when(
    verification_status %in% c("Verified", "Source Verified") ~ "Verified",
    verification_status == "Not Verified" ~ "Not Verified",
    TRUE ~ verification_status  # Pour g??rer les cas non sp??cifi??s
  ))


###############################################################################
################# Fusionner les variables entre elles ########################
#############################################################################
df <- df %>%
  mutate(
    # Calculer la proportion de principal rembours??
    principal_repaid_proportion = total_rec_prncp / total_pymnt,
    
    # Calculer le total des incidents n??gatifs
    total_negative_incidents = collections_12_mths_ex_med + pub_rec,
    
    # Calculer la proportion de comptes ouverts
    open_acc_proportion = open_acc / total_acc,
    
    # Calculer le cr??dit utilis??
    credit_used = revol_bal * revol_util / 100,
    
  )

# regroupement des deux variables : montant re??us et inter??ts re??us :
df$total_rec <- df$total_rec_int + df$total_rec_prncp

####

vars_to_remove <- c("total_rec_prncp","collections_12_mths_ex_med"
                    ,"pub_rec", "open_acc", "total_acc", "revol_bal", "revol_util",
                    "total_rec_int", "total_rec_prncp",
                    "purpose", "emp_length", "addr_state","home_ownership", "verification_status",
                    "sub_grade", "issue_d", "title", "zip_code", "earliest_cr_line", "last_pymnt_d",
                    "last_credit_pull_d","y","initial_list_status")


df <- df %>% select(-all_of(vars_to_remove))




# Exporter le dataframe dans un fichier Excel
file_path <- "C:/Users/kaffa/OneDrive/Bureau/R/Sorties/df_export.xlsx"
write.xlsx(df, file = file_path)

#########################################################
############# Encodage des variables ####################
#########################################################

# Liste des variables ?? encoder
vars_to_encode <- c("grade", "purpose_grouped", "emp_length_grouped", "region", "home_ownership_grouped", "verification_status_grouped")

# Encoder les variables et cr??er le mapping
mappings <- list()
df <- df %>%
  mutate(across(all_of(vars_to_encode), ~ {
    factor_var <- factor(.)
    levels_encoded <- as.numeric(factor_var)
    mappings[[cur_column()]] <<- data.frame(
      Variable = cur_column(),
      Original = levels(factor_var),
      Encoded = 1:length(levels(factor_var))
    )
    levels_encoded
  }))

# Convertir la liste de mappings en un seul dataframe
mapping_df <- do.call(rbind, mappings)

# R??organiser les colonnes du dataframe mapping
mapping_df <- mapping_df %>%
  select(Variable, Encoded, Original)

## Variable Term
df <- df %>%
  mutate(term = str_replace(term, "36 months", "36"),
         term = str_replace(term, "60 months", "60"))
df <- df %>%
  mutate(term = as.numeric(term))

mapping_df
###################################################################
###################################################################
###################################################################
###################################################################

head(df)
introduce(df) 
plot_intro(df)
plot_missing(df)
plot_bar(df)
plot_histogram(df)
## Criation d'un reporting pour un apergu rapide du comportement de nos donnies
create_report(
  data = df,
  output_format = html_document(toc = TRUE, toc_depth = 6, theme = "flatly"),
  output_file = "post_traitement",
  output_dir = paste0(getwd(), '/Sorties'),
  config = configure_report(
    add_plot_prcomp = TRUE,
    plot_bar_args = list("by_position" = "dodge"),
    plot_correlation_args = list("type" = "continuous", "cor_args" = list("use" = "all.obs"), "theme_config" = list("axis.text.x" = element_text("angle" = 90))),
    plot_prcomp_args = list("nrow" = 1, "ncol" = 3),
    plot_str_args = list(type = "diagonal", fontSize = 35, width = 1000, margin =
                           list(left = 350, right = 250)),
    global_ggtheme = quote(theme_linedraw(base_size = 16L)),
  )
)



###################################################################
################# LGD #############################################
df$LGD <- (df$installment * df$term - (df$total_pymnt)) / (df$installment * df$term)



# Exporter le dataframe dans un fichier Excel
file_path <- "C:/Users/kaffa/OneDrive/Bureau/R/Sorties/BDD2.xlsx"
write.xlsx(df, file = file_path)



##############################
##############################
##############################

df = df_abs
#df_abs = read_excel("C:/Users/kaffa/OneDrive/Bureau/R/BDD2.xlsx", sheet=1)

## Focus ACP
pca_noscale = prcomp(df_abs, scale = FALSE)
summary(pca_noscale) # La PC1 permet de capter 70,7% de la variance 
pca_noscale$rotation[, 1:5] # Sans standardisation des donn?es, Transportation.expense capte toute la variance 
# c'est pourquoi elle explique en grande partie PC1, qui elle m?me explique la variance totale
# les resultats sont donc biaises par l'echelle trop importante de la variable


pca = prcomp(df_abs, scale = TRUE)
summary(pca) # Apr?s standardisation, la CP1 capte 20,3% et la totalit? de CP1 et CP2 est 56,44%
pca$rotation[, 1:5]

plot(pca) # Apr?s le traitement, 7 composantes permettent d'expliquer environ 80% de la variance avec 13 CP


# Analyse contribution des variables dans chaque CP

#factoextra
pca_contrib = get_pca_var(pca)$contrib[, 1:10]
pca_contrib = data.frame("Feature" = rownames(pca_contrib),
                         pca_contrib)
#forcats
for (i in c('Dim.1', 'Dim.2')){
  a = pca_contrib %>% mutate(Feature = fct_reorder(Feature, desc(pca_contrib[[i]])))
  b = ggplot(data = a, aes(x=Feature, y=a[[i]])) +
    geom_bar(stat="identity", fill="#076fa2", alpha=.6, width=.4) +
    coord_flip() +
    xlab("Variables") +
    ylab(i) +
    ggtitle('Contribution des variables') +
    theme_bw()
  print(b)
}

# Analyse des corr?lations des variables avec chaque CP (les deux premi?res)
pca_cor = get_pca_var(pca)$cor[, 1:10] %>% data.frame('Feature' = rownames(get_pca_var(pca)$cor[, 1:10]))
for (i in c('Dim.1', 'Dim.2')){
  a = pca_cor %>% mutate(Feature = fct_reorder(Feature, desc(pca_cor[[i]])))
  b = ggplot(data = a, aes(x=Feature, y=a[[i]])) +
    geom_bar(stat="identity", fill="#076fa2", alpha=.6, width=.4) +
    coord_flip() +
    xlab("Variables") +
    ylab(i) +
    ggtitle('Corr?lation des variables')
  theme_bw()
  print(b)
}

## Visualisation de la variance expliqu?e par les CP par la m?thode du coude
fviz_eig(pca, 
         addlabels = TRUE, 
         ylim = c(0, 30),
         main="Elbow method plot",
         ncp = 24) 

## Choix du nombre de CP selon la r?gle de Kaiser (Valeur propre > 1) 
fviz_eig(pca, 
         addlabels = TRUE, 
         choice="eigenvalue",
         main="Kaiser rule method",
         ncp = 24) +
  geom_hline(yintercept=1, 
             linetype="dashed", 
             color = "red") 


paran(df_abs, all = TRUE, graph = TRUE) 
##
##
##
# De notre c?t?, on retiendra 6 CP, permettant d'expliquer 80% de la variance
pca_df = data.frame(pca$x[, 1:6])

# Repr?sentation graphique des 2 premi?res CP (56,44% de la variance expliqu?e)
plot(x = pca_df$PC1, y = pca_df$PC2) 

###
# CAH ne marche pas, probleme d'allocation de memoire, donc on passe direct au K-means.
###


########################
#########K-means########
########################

df_kmean = df_abs
##
df_kmean_pca = pca_df 
##
#
# L'algo pour selectionner le nombre de cluster optimal ne marche pas, probleme d'allocation de m??moire
# on specifie directement le nombre de cluster dans notre code
# On a decide de faire 3 clusters :
# Faible risque (Bon credit)
# Risque moyen (Credit moyen)
# Risque eleve (Credit risqu??)

kmeans = kmeans(x = pca_df, centers = 3, nstart = 1000, iter.max = 1000)
warnings()
# On specifie 3 clusters 

# Nombre d'individus dans chaque cluster :
# 1    # 2    # 3 
# 30924# 44809# 102445 
# Nombre d'observations total : 178 178
##


# Assignation des classes 
df_kmean$classekmean = kmeans$cluster
table(df_kmean$classekmean) 

# Representation graphiquement
autoplot(pca, df_kmean, colour = 'classekmean')
##

file_path <- "C:/Users/kaffa/OneDrive/Bureau/R/Sorties/df_kmean.xlsx"
# Exporter le dataframe au format Excel
write.xlsx(df_kmean, file = file_path)


## Determiner le k
elb_wss=rep(0, times=10)
for (k in 1:10){
  clus=kmeans(pca_df, centers=k)
  elb_wss[k]=clus$tot.withinss
}
# Plot
pdf("rplot.pdf")
# Plot
plot(1:10, elb_wss, type="b", xlab="NB of clusters", ylab="WSS")
dev.off()


###############################################################################################
########## K-MEANS - ANALYSE DES R?SULTATS PAR FORET ALEATOIRE ET ARBRE DE DECISION ###########
###############################################################################################

## Random Forest pour comprendre l'importance des variables
set.seed(123)
sample <- sample.split(df_kmean$classekmean, SplitRatio = 0.7)
train  <- subset(df_kmean, sample == TRUE)
test   <- subset(df_kmean, sample == FALSE)

rf = randomForest(x = train[, !names(train) %in% 'classekmean'],
                  y = train$classekmean,
                  importance = TRUE,
                  ntree = 500,
                  xtest = test[, !names(test) %in% 'classekmean'],
                  ytest = test$classekmean
)
rf2 = randomForest(x = df_kmean[, !names(df_kmean) %in% 'classekmean'],
                   y = df_kmean$classekmean,
                   importance = TRUE,
                   ntree = 500
)
varImpPlot(rf)
varImpPlot(rf2)
## Arbre de d?cision
dt = rpart(formula = classekmean ~.,
           data = train,
           method = 'class',
           control = rpart.control(cp = 0, 
                                   maxdepth = 8,
                                   minsplit = 150,
                                   minbucket = 70,
                                   xval = 0))
rpart.plot(dt, extra = 104)
for (i in list(train,test)){
  print(confusionMatrix(data = predict(dt, i, type = 'class'), 
                        reference = i[['classekmean']], 
                        dnn = c("Pr??dictions", "Vraies valeurs")))
} 

#####################################################################
########## K-MEANS - TESTS STATISTIQUES SUR LES VARIABLES ###########
#####################################################################

########################################################################

df_kmean$classekmean <- as.factor(df_kmean$classekmean)

# V??rifiez les noms des variables s??lectionn??es
num_vars <- names(subset(df_kmean, select = !names(df_kmean) %in% c('classekmean', 'loan_amnt', 'int_rate', 'installment', 'annual_inc','LGD','grade')))

# Ex??cutez la fonction test_stats apr??s avoir fait la conversion
test_stats(df_kmean,
           numvar = num_vars,
           var_classe = 'classekmean',
           qualvar = c('grade'),
           name_file = 'C:/Users/kaffa/OneDrive/Bureau/R/Projet R/reporting k_means/resultats_tests_classe_kmeans.xlsx')


df_kmean = read_xlsx("C:/Users/kaffa/OneDrive/Bureau/R/Projet R/df_kmean.xlsx")
##############################################################################################################
########## K-MEANS - ANALYSE DES RESULTATS PAR VISUALISATION GRAPHIQUE & STATISTIQUES DESCRIPTIVES ###########
##############################################################################################################

# Identification des variables cat??gorielles dans df_kmean
categorical_vars <- c('classekmean', 'grade','term')  # Liste des variables cat??gorielles

# Conversion des variables en facteur
df_kmean[categorical_vars] <- lapply(df_kmean[categorical_vars], as.factor)

## Statistiques descriptives par groupe
# Reporting avec la classekmean
create_report(
  data = df_kmean,
  output_format = html_document(toc = TRUE, toc_depth = 6, theme = "flatly"),
  output_file = "Absence_repor9",
  output_dir = paste0(getwd(), '/Sorties'),
  y = 'classekmean',
  config = configure_report(
    add_plot_prcomp = TRUE,
    plot_bar_args = list("by_position" = "fill", by = 'classekmean'),
    plot_correlation_args = list("type" = "continuous", "cor_args" = list("use" = "all.obs"), "theme_config" = list("axis.text.x" = element_text("angle" = 90))),
    plot_prcomp_args = list("nrow" = 1, "ncol" = 3),
    plot_str_args = list(type = "diagonal", fontSize = 35, width = 1000, margin = list(left = 350, right = 250)),
    global_ggtheme = quote(theme_linedraw(base_size = 16L))
  )
)
##################################################################################


# Inclure la variable 'classekmean' dans le dataframe 'cat_var'

cat_var = df_kmean[, c('term', 'classekmean')]
# Convertir 'term' et 'classekmean' en caract??res si n??cessaire

cat_var = cat_var %>% mutate_at(c('term', 'classekmean'), as.character)
# Convertir 'classekmean' en facteur pour l'utiliser comme variable discr??te

cat_var$classekmean = as.factor(cat_var$classekmean)
# Ordonner les colonnes par leur nom

cat_var = cat_var[, order(names(cat_var))]
# Cr??er le graphique ?? barres

plot_bar(cat_var, by = 'classekmean', by_position = 'fill', nrow = 4, ncol = 3)

#################################################################################
# Variables cat?gorielles
Freqtable = ExpCTable(df_kmean, Target = 'classekmean', per = TRUE, round = 0) 
Freqtable = subset(Freqtable, subset = !(Freqtable$CATEGORY == 'TOTAL' & Freqtable$Number == '%'))
Freqtable[order(Freqtable$Number), ] %>% flextable()
write_xlsx(Freqtable, paste0(getwd(), 'C:/Users/kaffa/OneDrive/Bureau/R/Sorties/at_kmeans.xlsx'))
cat_var = df_kmean[, c('term')]
cat_var = cat_var %>% mutate_at(names(cat_var), as.character)
cat_var = cat_var[, order(names(cat_var))]


plot_bar(cat_var, by = 'classekmean', by_position = 'fill', nrow = 4, ncol = 3)

# Variables num?riques 
ExpNumStat(df_kmean, by="GA", gp="classekmean", Outlier=TRUE, Qnt = c(.25, .75), round = 2, Nlim = 4) %>% subset(select = c('Vname', 'Group', 'sum', 'min', 'max', 'mean', 'median', 'SD', '25%', '75%')) %>% flextable()
num_var = df_kmean[, !names(df_kmean) %in% c('loan_amnt', 'int_rate', 'installment', 'annual_inc','LGD')]

num_var = num_var[, order(names(num_var))]
plot_boxplot(num_var, by ='classekmean', geom_boxplot_args = list(aes("fill" = "classekmean")), nrow = 3, ncol = 3)
write_xlsx(ExpNumStat(num_var, by="GA", gp="classekmean", Outlier=TRUE, Qnt = c(.25, .75), round = 2, Nlim = 4) 
           %>% subset(select = c('Vname', 'Group', 'sum', 'min', 'max', 'mean', 'median', 'SD', '25%', '75%')), 
           paste0(getwd(), 'C:/Users/kaffa/OneDrive/Bureau/R/Sorties/at_kmeans.xlsx'))






