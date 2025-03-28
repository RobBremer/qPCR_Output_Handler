
getwd()

library(purrr)
library(readr)
library(tidyverse)
library(readxl)
library(janitor)
library(outliers)
library(writexl)
library(dplyr)
?list.files

excel_files <- list.files(path = "C:/Users/Robert.Bremer/Documents/NOAA Projects/qPCR Outputs Old/mwkBB",
    pattern = "*.xls",
    full.names = TRUE)
    
excel_files

searcher <- function(df) {
  df <- read_excel(df, sheet = "Results", col_names = F)
  block <- df[1,2]
  chemistry <- df[2,2]
  filename <- df[3,2]
  runtime <- df[4,2]
  instrument <- df[5,2]
  reference <- df[6,2]
  print(chemistry)
  print(reference)
  df_edited <- df %>%
    row_to_names(8) %>%
    clean_names() %>%
    mutate("block" = block) %>%
    mutate("chemistry" = chemistry) %>%
    mutate("filename" = filename) %>%
    mutate("runtime" = runtime) %>%
    mutate("instrument" = instrument) %>%
    mutate("reference" = reference)
  return(df_edited)
}


all_df <- map_df(excel_files, searcher)

all_df_edited <- all_df %>%
  mutate(sample_name = as.numeric(sample_name)) %>%
  mutate(across(c(ct, ct_mean, ct_sd, quantity, quantity_sd, ct_threshold), na_if, "Undetermined")) %>%
  mutate(ct = as.numeric(ct)) %>%
  mutate(ct_mean = as.numeric(ct_mean)) %>%
  mutate(ct_sd = as.numeric(ct_sd)) %>%
  mutate(quantity = as.numeric(quanitity)) %>%
  mutate(quantity_sd = as.numeric(quantity_sd)) %>%
  mutate(ct_threshold = as.numeric(ct_threshold))

all_df_edited$block <- unlist(all_df_edited$block)
all_df_edited$chemistry <- unlist(all_df_edited$chemistry)
all_df_edited$filename <- unlist(all_df_edited$filename)
all_df_edited$runtime <- unlist(all_df_edited$runtime)
all_df_edited$instrument <- unlist(all_df_edited$instrument)
all_df_edited$reference <- unlist(all_df_edited$reference)

write_xlsx(all_df_edited, "C:/Users/Robert.Bremer/Documents/mwkBB_qPCR_20250325.xlsx")

df.list <- lapply(excel_files, read_excel(path = excel_files, sheet="Results"))

df.list1 <- map_df(df.list, searcher)

?extract

?map_df
# A lot of the stuff directly below this is experimental, trying to get it to work perfect every time, with the old and new qpcr machines
excel_files1 <- purrr::map_df(excel_files, ~.x, read_excel(sheet = "Results"), .id = "id")

which(excel_files == "Chemistry", arr.ind = TRUE)

excel_files <- purrr::map_df(excel_files, ~.x) %>%
  read_excel(sheet="Results")
  
which(excel_files == "Chemistry", arr.ind = TRUE)

  #add_column(.,"highsd" = NA) %>%
  mutate(experiment = .[2,2]) %>% #Adjusted depending on sheet [2,2] or [3,2]
  mutate(date = .[3,2]) %>% # these are adjusted depending on sheet [3,2]
  mutate(chemistry = .[1,2]) %>% # these are adjusted depending on sheet [1,2]
  mutate(instrumentType = .[4,2]) %>%
  mutate(passiveReference = .[5,2])) %>%
  drop_na(3) %>%
  row_to_names(row_number = 1) %>%
  clean_names() %>%
  dplyr::rename("Experiment" = 24,"date" = 25,"instrument" = 26,"passive_reference" = 27) %>% #Change slightly
  filter(task == "UNKNOWN") %>%
  filter(sample_name != "NEC") %>%
  mutate(assumed_quantity = ifelse(is.na(quantity), 0, quantity)) %>%
  group_by(sample_name, target_name) %>%
  mutate(quantity_mean = mean(as.numeric(assumed_quantity)))
  #mutate(occurrence = row_number())
  
  #mutate(copiesPer100ml = as.numeric(quantity_mean)*6.6667) %>%
  #relocate(copiesPer100ml, .after = quantity_sd)
  
excel_files$Experiment = as.character(unlist(excel_files$Experiment))
excel_files$date = as.character(unlist(excel_files$date))

#Adding this in to perform multiple imputation
duplicated_qpcr_data <- excel_files %>%
  group_by_all() |>
  filter(n() > 1) |>
  ungroup()

qpcr_data <- excel_files %>%
  ungroup() %>%
  distinct() %>%
  filter(sample_name != "bad std") %>%
  filter(sample_name != "outlier 10 copy STD") %>%
  select(sample_name, target_name, ct_mean, quantity_mean) %>%
  #mutate(row = row_number()) %>%
  pivot_wider(names_from = target_name, values_from = c(ct_mean, quantity_mean)) %>%
  #select(-row) %>%
  clean_names()
  
qpcr_data %>%
  write_xlsx(path = ("C:/Users/Robert.Bremer/Documents/NOAA Projects/MiamiBeach_2025_02_12.xlsx"))

excel_files <- excel_files %>%
  select(c(sample_name,target_name,ct_mean,quantity_mean,Experiment,date))%>%
  distinct()
  
qpcr_data %>%
  write_xlsx(path = ("C:/Users/Robert.Bremer/Documents/NOAA Projects/MiamiBeach_2025_01_25.xlsx"))

??qpcrImpute

#If all 3 values exist ,check if any of the values are outliers

#Merge with all excel sheets (and do this protocol to every sheet in a folder)

# Calculate Interquartile range
# <1st Quartile-1.5 IQ = outlier, >3rd Quartile+1.5 IQ = outlier
# Add notes to sheet saying something is an outlier

