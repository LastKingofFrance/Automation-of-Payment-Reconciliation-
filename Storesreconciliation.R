rm(list = ls())
setwd("C:/Users/anyariki/Documents/Working with R/StoresReconciliation")
library(readxl)

# Define the file name
file_name <- "Surgi cons.xlsx"  # Change this as needed for other files

# Extract the name of the file without the extension
base_name <- tools::file_path_sans_ext(file_name)

# Import data from the Excel file
lab_stock_data <- read_excel(file_name, sheet = 1)
lab_cost_data <- read_excel(file_name, sheet = 2)
# Note above I inputted the sheets separately as R only reads the first sheet by default
#Isolating the columns we need 
library(dplyr)
isolated_stock_data <- lab_stock_data %>%
  select(`Document No.`, Amount)
isolated_cost_data <- lab_cost_data %>%
  select(`Document No.`, Amount)
#Now we need to do sum up all the amounts with the same document no
summed_stock_data<-isolated_stock_data %>%
  group_by(`Document No.`) %>%
  summarise(Total_Amount = sum(Amount, na.rm = TRUE)) 

summed_cost_data <- isolated_cost_data %>%
  group_by(`Document No.`) %>%
  summarise(Total_Amount = sum(Amount, na.rm = TRUE))
#Now we need to right join based on document numbers for both stock and cost data
#First we rename columns to avoid confusion 
summed_stock_data <- summed_stock_data %>%
  rename(Total_Amount_Stock = Total_Amount)

summed_cost_data <- summed_cost_data %>%
  rename(Total_Amount_Cost = Total_Amount)
#We use leftjoins to replace the VLOOKUP part of reconciliation 
#that is we join values with the same document number to both dataframes 
summed_stock_data <- summed_stock_data %>%
  left_join(summed_cost_data %>%
              select(`Document No.`, Total_Amount_Cost), 
            by = "Document No.")
summed_cost_data <- summed_cost_data %>%
  left_join(summed_stock_data %>%
              select(`Document No.`, Total_Amount_Cost),
            by = "Document No.")
summed_cost_data <- summed_cost_data %>%
  rename(Total_Amount_Stock.y = Total_Amount_Cost.y)
#Now we add the variance column 
summed_stock_data <- summed_stock_data %>%
  mutate(Variance = Total_Amount_Stock + Total_Amount_Cost)
summed_cost_data <- summed_cost_data %>%
  mutate(Variance = Total_Amount_Cost.x + Total_Amount_Stock.y)
#At this stage the NA values arise, we turn NA to zero in the next line
# Step 2: Replace NA values with 0 after the join and then recalculate our variance
library(tidyr)
summed_stock_data <- summed_stock_data %>%
  replace_na(list(Total_Amount_Cost = 0))
# Now calculate the Variance column
summed_stock_data <- summed_stock_data %>%
  mutate(Variance = Total_Amount_Stock + Total_Amount_Cost)

summed_cost_data <- summed_cost_data %>%
  replace_na(list(Total_Amount_Stock.y = 0))
# Now calculate the Variance column
summed_cost_data <- summed_cost_data %>%
  mutate(Variance = Total_Amount_Cost.x + Total_Amount_Stock.y)
#Now we add the items pertaining to purchase 
#which is document numbers containing PINV and GRN in summed_stock_data
#filtering data with "GRN" and "PINV" in the Document No. column. 
filtered_data_purchases <- summed_stock_data %>%
  filter(grepl("GRN", `Document No.`) | grepl("PINV", `Document No.`)) %>%
  summarise(Purchases = sum(Variance, na.rm = TRUE))
#Showcasing Purchases
print(filtered_data_purchases$Purchases)
#Now we repeat to get Purchase returns total
#filter out PRS Document Nos
filtered_data_PRS <- summed_stock_data %>%
  filter(grepl("PRS", `Document No.`)) %>%
  summarise(Purchase_Returns = sum(Variance, na.rm = TRUE))
#Showing Purchase Returns
print(filtered_data_PRS$Purchase_Returns)

#Getting Purchase Credit Memo Values
# Filter summed_stock_data for rows with "PCRE" in the Document No.
filtered_data_pcre <- summed_stock_data %>%
  filter(grepl("PCRE", `Document No.`)) %>%
  summarise(Purchase_Credit_Memo = sum(Variance, na.rm = TRUE))

# View the result
print(filtered_data_pcre$Purchase_Credit_Memo)

#searching for "PINV" in summed_cost_data
#this is to find the total of purchased directly not through stock account 
# Filter summed_cost_data for rows with "PINV" in the Document No.
filtered_data_purchaseddirectly <- summed_cost_data %>%
  filter(grepl("PINV", `Document No.`)) %>%
  summarise(Purchased_Directly_No_Stock_Account = sum(Variance, na.rm = TRUE))

# View the result
print(filtered_data_purchaseddirectly$Purchased_Directly_No_Stock_Account)
#Finally we combined the important information in one sheet and export it. 

# Combine the results into a single data frame
important_values <- data.frame(
  Category = c("Purchased_Directly_No_Stock_Account", "Purchase_Credit_Memo", "Purchase_Returns", "Purchases"),
  Value = c(filtered_data_purchaseddirectly$Purchased_Directly_No_Stock_Account,
            filtered_data_pcre$Purchase_Credit_Memo,
            filtered_data_PRS$Purchase_Returns,
            filtered_data_purchases$Purchases)
)

# Print the important values
print(important_values)


# install.packages("writexl")  # Uncomment if not installed
library(writexl)

# Dynamically name the variable to store important values
assign(paste0(base_name, "_important_values"), important_values)

# Print the dynamically named data frame
print(get(paste0(base_name, "_important_values")))

# Write the important values to an Excel file
write_xlsx(important_values, path = paste0(base_name, "_important_values.xlsx"))