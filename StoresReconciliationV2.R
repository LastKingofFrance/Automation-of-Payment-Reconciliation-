#Clearing the environment and setting work directory to where the files are
rm(list = ls())
setwd("C:/Users/anyariki/Documents/Working with R/StoresReconciliation")
# Load required libraries
library(readxl)
library(dplyr)
library(tidyr)
library(writexl)

# Dynamic file selection
file_name <- file.choose()
base_name <- tools::file_path_sans_ext(basename(file_name))

# Load data
lab_stock_data <- read_excel(file_name, sheet = 1) %>% select(`Document No.`, Amount)
lab_cost_data <- read_excel(file_name, sheet = 2) %>% select(`Document No.`, Amount)

# Summarize data with group_by inside summarise
summed_stock_data <- lab_stock_data %>%
  summarise(Total_Amount_Stock = sum(Amount, na.rm = TRUE), .by = `Document No.`)

summed_cost_data <- lab_cost_data %>%
  summarise(Total_Amount_Cost = sum(Amount, na.rm = TRUE), .by = `Document No.`)

# Join and calculate variance
final_data <- full_join(summed_stock_data, summed_cost_data, by = "Document No.") %>%
  replace_na(list(Total_Amount_Stock = 0, Total_Amount_Cost = 0)) %>%
  mutate(Variance = Total_Amount_Stock + Total_Amount_Cost)

# Filter specific categories
categories <- list(
  Purchases = "GRN|PINV",
  Purchase_Returns = "PRS",
  Purchase_Credit_Memo = "PCRE",
  # Change for this category to use summed_cost_data
  Purchased_Directly_No_Stock_Account = "PINV"
)

# Extract important values and ensure Category is the first column
important_values <- bind_rows(
  lapply(names(categories), function(cat) {
    if (cat == "Purchased_Directly_No_Stock_Account") {
      # Use summed_cost_data for Purchased_Directly_No_Stock_Account
      summed_cost_data %>%
        filter(grepl(categories[[cat]], `Document No.`)) %>%
        summarise(Value = sum(Total_Amount_Cost, na.rm = TRUE)) %>%
        mutate(Category = cat) %>%
        select(Category, Value)
    } else {
      # For other categories, use final_data
      final_data %>%
        filter(grepl(categories[[cat]], `Document No.`)) %>%
        summarise(Value = sum(Variance, na.rm = TRUE)) %>%
        mutate(Category = cat) %>%
        select(Category, Value)
    }
  })
)

# Export results
write_xlsx(important_values, path = paste0(base_name, "_important_values.xlsx"))
