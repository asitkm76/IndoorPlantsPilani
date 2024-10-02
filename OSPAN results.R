# Load necessary libraries
library(openxlsx)
library(foreign)
library(tidyverse)
library(car)
library(moments)
library(readxl)
library(lsr)

# Set working directory
setwd("F:/Google drive data/Plants and IEQ Project/Main experiments/Data analysis")

# Read the data from the Excel file
Data <- read_excel("All data.xlsx", sheet = "Sheet1")

# Convert Group and Survey variables to factors
Data$Group <- factor(Data$Group)
Data$Survey <- factor(Data$Survey)

# Define the column names you want to analyze
column_names <- c('PercentageWMC', 'PercentageMaths', 'AverageResponseTimeSec')

# Initialize an empty data frame for results
results <- data.frame()

# Filter data for Survey III only
SurveyIII_Data <- Data %>% 
  filter(Survey == "III" & (Group == "WoP" | Group == "WP")) %>% 
  select(Group, Survey, all_of(column_names))

# Iterate over each of the specified columns
for (column_name in column_names) {
  
  # Calculate mean, median, and standard deviation for each Group within Survey III
  summary_stats <- SurveyIII_Data %>%
    group_by(Group) %>%
    summarize(Mean = mean(get(column_name), na.rm = TRUE),
              Median = median(get(column_name), na.rm = TRUE),
              SD = sd(get(column_name), na.rm = TRUE))
  
  # Get the means for WoP and WP for Survey III
  M1 <- summary_stats$Mean[summary_stats$Group == "WoP"]
  M2 <- summary_stats$Mean[summary_stats$Group == "WP"]
  
  # Add mean and residuals to the data frame
  SurveyIII_Data <- SurveyIII_Data %>%
    mutate(mean = ifelse(Group == "WoP", M1, M2),
           Residual = get(column_name) - mean)
  
  # Filter out NA values for the current column
  valid_data <- SurveyIII_Data %>%
    filter(!is.na(get(column_name)))
  
  # Extract data for t-test
  x <- valid_data[[column_name]][valid_data$Group == "WoP"]
  y <- valid_data[[column_name]][valid_data$Group == "WP"]
  
  # Check lengths of the two groups
  cat("Length of Group WoP:", length(x), "\n")
  cat("Length of Group WP:", length(y), "\n")
  
  # Perform t-test only if both groups have data
  if (length(x) > 0 && length(y) > 0) {
    t_test <- t.test(x, y, var.equal = FALSE)
    
    # Calculate effect size (Cohen's d)
    effect_size <- cohensD(x, y)
    
    # Create a data frame with the required statistics
    column_results <- data.frame(
      Column = column_name,
      Group = c("WoP", "WP"),
      Median = summary_stats$Median,
      Mean = c(mean(x, na.rm = TRUE), mean(y, na.rm = TRUE)),
      SD = summary_stats$SD,
      p = as.numeric(t_test$p.value),
      t = as.numeric(t_test$statistic),
      CohenD = effect_size
    )
  } else {
    # Handle cases where one of the groups has no data
    column_results <- data.frame(
      Column = column_name,
      Group = c("WoP", "WP"),
      Median = summary_stats$Median,
      Mean = c(NA, NA),
      SD = summary_stats$SD,
      p = NA,
      t = NA,
      CohenD = NA
    )
  }
  
  # Append the column results to the overall results data frame
  results <- bind_rows(results, column_results)
}

# Save the results to a CSV file
write.csv(results, file = "OSPAN results.csv", row.names = FALSE)
warnings()
