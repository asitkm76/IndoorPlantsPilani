# Load necessary libraries
library(openxlsx)
library(psych)
library(foreign)
library(tidyverse)
library(effsize)
library(readxl)
library(car)
library(rcompanion)
library(moments)
library(coin)

# Set working directory
setwd("F:/Google drive data/Plants and IEQ Project/Main experiments/Data analysis")

# Read the data from the Excel file
Data <- read_excel("All data.xlsx", sheet = "Sheet1")

# Convert Group and Survey variables to factors
Data$Group <- factor(Data$Group)
Data$Survey <- factor(Data$Survey)

# Define the column names you want to analyze
column_names <- c('DryItchyWateryEyes', 'NoseThroatIrritationDryness', 'DifficultyBreathing', 'LethargyTiredness', 
                  'Headache','SkinIrritationDryItching','Dizziness','DifficultyConcentrating')


###################
#SBS_WoP Survey II Vs WoP Survey IV 
# Initialize an empty data frame for results
results1 <- data.frame()
# Filter data for WoP group only
WoP_Data <- Data %>% 
  filter(Group == "WoP" & (Survey == "II" | Survey == "IV")) %>% 
  select(Group, Survey, all_of(column_names))

# Iterate over each of the specified columns
for (column_name in column_names) {
  
  # Calculate mean, median, and standard deviation for each Survey within WoP group
  summary_stats <- WoP_Data %>%
    group_by(Survey) %>%
    summarize(Mean = mean(get(column_name), na.rm = TRUE),
              Median = median(get(column_name), na.rm = TRUE),
              SD = sd(get(column_name), na.rm = TRUE))
  
  # Get the means for Surveys II and IV
  M1 <- summary_stats$Mean[summary_stats$Survey == "II"]
  M2 <- summary_stats$Mean[summary_stats$Survey == "IV"]
  
  # Add mean and residuals to the data frame
  WoP_Data <- WoP_Data %>%
    mutate(mean = ifelse(Survey == "II", M1, M2),
           Residual = get(column_name) - mean)
  
  # Filter out NA values for the current column
  valid_data <- WoP_Data %>%
    filter(!is.na(get(column_name)))
  
  # Extract data for Wilcoxon test
  x <- valid_data[[column_name]][valid_data$Survey == "II"]
  y <- valid_data[[column_name]][valid_data$Survey == "IV"]
  
  # Check lengths of the two groups
  cat("Length of Survey II:", length(x), "\n")
  cat("Length of Survey IV:", length(y), "\n")
  
  # Perform paired Wilcoxon test only if both groups have data
  if (length(x) > 0 && length(y) > 0 && length(x) == length(y)) {
    wilcox_test <- wilcox.test(x, y, distribution = "asymptotic",alternative = "less", paired = TRUE)
    
    # Calculate effect size
    effect_size <- wilcoxonPairedR(x = c(x, y), g = rep(c("II", "IV"), times = c(length(x), length(y))), ci = TRUE)
    
    # Create a data frame with the required statistics
    column_results <- data.frame(
      Column = column_name,
      Survey = c("II", "IV"),
      Median = summary_stats$Median,
      Mean = c(mean(x, na.rm = TRUE), mean(y, na.rm = TRUE)),
      SD = summary_stats$SD,
      p = as.numeric(wilcox_test$p.value),
      Z = as.numeric(wilcox_test$statistic),
      r = effect_size$r
    )
  } else {
    # Handle cases where one of the surveys has no data or lengths do not match
    column_results <- data.frame(
      Column = column_name,
      Survey = c("II", "IV"),
      Median = summary_stats$Median,
      Mean = c(NA, NA),
      SD = summary_stats$SD,
      p = NA,
      Z = NA,
      r = NA
    )
  }
  
  # Append the column results to the overall results data frame
  results1 <- bind_rows(results1, column_results)
}

# Save the results to a CSV file
write.csv(results1, file = "SBS_WoP Survey II Vs IV.csv", row.names = FALSE)
warnings()

###################
#SBS_WP Survey II Vs IV
# Initialize an empty data frame for results
results2 <- data.frame()
# Filter data for WP group only
WP_Data <- Data %>% 
  filter(Group == "WP" & (Survey == "II" | Survey == "IV")) %>% 
  select(Group, Survey, all_of(column_names))

# Iterate over each of the specified columns
for (column_name in column_names) {
  
  # Calculate mean, median, and standard deviation for each Survey within WP group
  summary_stats <- WP_Data %>%
    group_by(Survey) %>%
    summarize(Mean = mean(get(column_name), na.rm = TRUE),
              Median = median(get(column_name), na.rm = TRUE),
              SD = sd(get(column_name), na.rm = TRUE))
  
  # Get the means for Surveys II and IV
  M1 <- summary_stats$Mean[summary_stats$Survey == "II"]
  M2 <- summary_stats$Mean[summary_stats$Survey == "IV"]
  
  # Add mean and residuals to the data frame
  WP_Data <- WP_Data %>%
    mutate(mean = ifelse(Survey == "II", M1, M2),
           Residual = get(column_name) - mean)
  
  # Filter out NA values for the current column
  valid_data <- WP_Data %>%
    filter(!is.na(get(column_name)))
  
  # Extract data for Wilcoxon test
  x <- valid_data[[column_name]][valid_data$Survey == "II"]
  y <- valid_data[[column_name]][valid_data$Survey == "IV"]
  
  # Check lengths of the two groups
  cat("Length of Survey II:", length(x), "\n")
  cat("Length of Survey IV:", length(y), "\n")
  
  # Perform paired Wilcoxon test only if both groups have data
  if (length(x) > 0 && length(y) > 0 && length(x) == length(y)) {
    wilcox_test <- wilcox.test(x, y, distribution = "asymptotic",alternative = "greater", paired = TRUE)
    
    # Calculate effect size
    effect_size <- wilcoxonPairedR(x = c(x, y), g = rep(c("II", "IV"), times = c(length(x), length(y))), ci = TRUE)
    
    # Create a data frame with the required statistics
    column_results <- data.frame(
      Column = column_name,
      Survey = c("II", "IV"),
      Median = summary_stats$Median,
      Mean = c(mean(x, na.rm = TRUE), mean(y, na.rm = TRUE)),
      SD = summary_stats$SD,
      p = as.numeric(wilcox_test$p.value),
      Z = as.numeric(wilcox_test$statistic),
      r = effect_size$r
    )
  } else {
    # Handle cases where one of the surveys has no data or lengths do not match
    column_results <- data.frame(
      Column = column_name,
      Survey = c("II", "IV"),
      Median = summary_stats$Median,
      Mean = c(NA, NA),
      SD = summary_stats$SD,
      p = NA,
      Z = NA,
      r = NA
    )
  }
  
  # Append the column results to the overall results data frame
  results2 <- bind_rows(results2, column_results)
}

# Save the results to a CSV file
write.csv(results2, file = "SBS_WP Survey II Vs IV.csv", row.names = FALSE)
warnings()

###################
#SBS_WoP Survey IV Vs WP Survey IV
# Initialize an empty data frame for results
results3 <- data.frame()
# Filter data for SurveyIV only
SurveyIV_Data <- Data %>% 
  filter(Survey == "IV" & (Group == "WoP" | Group == "WP")) %>% 
  select(Group, Survey, all_of(column_names))

# Iterate over each of the specified columns
for (column_name in column_names) {
  
  # Calculate mean, median, and standard deviation for each Group within Survey IV
  summary_stats <- SurveyIV_Data %>%
    group_by(Group) %>%
    summarize(Mean = mean(get(column_name), na.rm = TRUE),
              Median = median(get(column_name), na.rm = TRUE),
              SD = sd(get(column_name), na.rm = TRUE))
  
  # Get the means for WoP and WP for Survey IV
  M1 <- summary_stats$Mean[summary_stats$Group == "WoP"]
  M2 <- summary_stats$Mean[summary_stats$Group == "WP"]
  
  # Add mean and residuals to the data frame
  SurveyIV_Data <- SurveyIV_Data %>%
    mutate(mean = ifelse(Group == "WoP", M1, M2),
           Residual = get(column_name) - mean)
  
  # Filter out NA values for the current column
  valid_data <- SurveyIV_Data %>%
    filter(!is.na(get(column_name)))
  
  # Extract data for Wilcoxon test
  x <- valid_data[[column_name]][valid_data$Group == "WoP"]
  y <- valid_data[[column_name]][valid_data$Group == "WP"]
  
  # Check lengths of the two groups
  cat("Length of Survey II:", length(x), "\n")
  cat("Length of Survey IV:", length(y), "\n")
  
  # Perform unpaired Wilcoxon test only if both groups have data
  if (length(x) > 0 && length(y) > 0 && length(x) == length(y)) {
    wilcox_test <- wilcox.test(x, y, distribution = "asymptotic",alternative = "greater", paired = FALSE)
    
    # Calculate effect size
    effect_size <- wilcoxonR(x = c(x, y), g = rep(c("WoP", "WP"), times = c(length(x), length(y))), ci = TRUE)
    
    # Create a data frame with the required statistics
    column_results <- data.frame(
      Column = column_name,
      Group = c("WoP", "WP"),
      Median = summary_stats$Median,
      Mean = c(mean(x, na.rm = TRUE), mean(y, na.rm = TRUE)),
      SD = summary_stats$SD,
      p = as.numeric(wilcox_test$p.value),
      Z = as.numeric(wilcox_test$statistic),
      r = effect_size$r
    )
  } else {
    # Handle cases where one of the surveys has no data or lengths do not match
    column_results <- data.frame(
      Column = column_name,
      Group = c("WoP", "WP"),
      Median = summary_stats$Median,
      Mean = c(NA, NA),
      SD = summary_stats$SD,
      p = NA,
      Z = NA,
      r = NA
    )
  }
  
  # Append the column results to the overall results data frame
  results3 <- bind_rows(results3, column_results)
}

# Save the results to a CSV file
write.csv(results3, file = "SBS_WoP Survey IV Vs WP Survey IV.csv", row.names = FALSE)
