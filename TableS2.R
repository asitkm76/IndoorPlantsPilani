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
# Set the working directory
setwd("F:/Google drive data/Plants and IEQ Project/Main experiments/Data analysis")

# Read the data from the Excel file
Data <- read_excel("All data.xlsx", sheet = "Sheet1")

# Convert Group and Survey variables to factors
Data$Group <- factor(Data$Group)
Data$Survey <- factor(Data$Survey)

# List of specific column_names to test
column_names <- c("Lighting", "GlareAndReflection", "SpaceAvailable", "Furnishing", 
               "RoomDecoration", "RoomCleanliness", "Noise", 
               "OverallVisualComfort", "OverallRoomEnvironment")

# Initialize an empty data frame for results
results <- data.frame()
# Filter data for SurveyIV only
SurveyI_Data <- Data %>% 
  filter(Survey == "I" & (Group == "WoP" | Group == "WP")) %>% 
  select(Group, Survey, all_of(column_names))

# Iterate over each of the specified columns
for (column_name in column_names) {
  
  # Calculate mean, median, and standard deviation for each Group within Survey IV
  summary_stats <- SurveyI_Data %>%
    group_by(Group) %>%
    summarize(Mean = mean(get(column_name), na.rm = TRUE),
              Median = median(get(column_name), na.rm = TRUE),
              SD = sd(get(column_name), na.rm = TRUE))
  
  # Get the means for WoP and WP for Survey I
  M1 <- summary_stats$Mean[summary_stats$Group == "WoP"]
  M2 <- summary_stats$Mean[summary_stats$Group == "WP"]
  
  # Add mean and residuals to the data frame
  SurveyI_Data <- SurveyI_Data %>%
    mutate(mean = ifelse(Group == "WoP", M1, M2),
           Residual = get(column_name) - mean)
  
  # Filter out NA values for the current column
  valid_data <- SurveyI_Data %>%
    filter(!is.na(get(column_name)))
  
  # Extract data for Wilcoxon test
  x <- valid_data[[column_name]][valid_data$Group == "WoP"]
  y <- valid_data[[column_name]][valid_data$Group == "WP"]
  
  # Check lengths of the two groups
  cat("Length of Group WoP:", length(x), "\n")
  cat("Length of Group WP:", length(y), "\n")
  
  # Perform unpaired Wilcoxon test only if both groups have data
  if (length(x) > 0 && length(y) > 0 && length(x) == length(y)) {
    wilcox_test <- wilcox.test(x, y, distribution = "asymptotic",alternative = "less", paired = FALSE)
    
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
  result <- bind_rows(result, column_results)
}

# Save the results to a CSV file
write.csv(result, file = "Working environment_WoP Survey IV Vs WP Survey IV.csv", row.names = FALSE)
