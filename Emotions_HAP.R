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

# Filter data for WoP group and 'HAP' column only for Surveys II and IV
WoP_Data <- Data %>% 
  filter(Group == "WoP" & (Survey == "II" | Survey == "IV")) %>% 
  select(Group, Survey, HAP)

# Calculate mean, median, and standard deviation for 'HAP' within WoP group
summary_stats <- WoP_Data %>%
  group_by(Survey) %>%
  summarize(Mean = mean(HAP, na.rm = TRUE),
            Median = median(HAP, na.rm = TRUE),
            SD = sd(HAP, na.rm = TRUE))

# Get the means for Surveys II and IV
M1 <- summary_stats$Mean[summary_stats$Survey == "II"]
M2 <- summary_stats$Mean[summary_stats$Survey == "IV"]

# Add mean and residuals to the data frame
WoP_Data <- WoP_Data %>%
  mutate(mean = ifelse(Survey == "II", M1, M2),
         Residual = HAP - mean)

# Filter out NA values for 'HAP'
valid_data <- WoP_Data %>%
  filter(!is.na(HAP))

# Extract data for Wilcoxon test
x <- valid_data$HAP[valid_data$Survey == "II"]
y <- valid_data$HAP[valid_data$Survey == "IV"]

# Perform paired Wilcoxon test if both groups have data
if (length(x) > 0 && length(y) > 0 && length(x) == length(y)) {
  wilcox_test <- wilcox.test(x, y, distribution = "asymptotic", alternative = "greater", paired = TRUE)
  
  # Calculate effect size
  effect_size <- wilcoxonPairedR(x = c(x, y), g = rep(c("II", "IV"), times = c(length(x), length(y))), ci = TRUE)
  
  # Create a results data frame
  results1 <- data.frame(
    Column = 'HAP',
    Survey = c("II", "IV"),
    Median = summary_stats$Median,
    Mean = c(mean(x, na.rm = TRUE), mean(y, na.rm = TRUE)),
    SD = summary_stats$SD,
    p = as.numeric(wilcox_test$p.value),
    Z = as.numeric(wilcox_test$statistic),
    r = effect_size$r
  )
} else {
  results1 <- data.frame(
    Column = 'HAP',
    Survey = c("II", "IV"),
    Median = summary_stats$Median,
    Mean = c(NA, NA),
    SD = summary_stats$SD,
    p = NA,
    Z = NA,
    r = NA
  )
}

# Save the results to a CSV file
write.csv(results1, file = "HAP_WoP_Survey_II_vs_IV.csv", row.names = FALSE)
warnings()

######
##HAP_WP_Survey_II_vs_IV
# Filter data for WP group and 'HAP' column only for Surveys II and IV
WP_Data <- Data %>% 
  filter(Group == "WP" & (Survey == "II" | Survey == "IV")) %>% 
  select(Group, Survey, HAP)

# Calculate mean, median, and standard deviation for 'HAP' within WoP group
summary_stats <- WP_Data %>%
  group_by(Survey) %>%
  summarize(Mean = mean(HAP, na.rm = TRUE),
            Median = median(HAP, na.rm = TRUE),
            SD = sd(HAP, na.rm = TRUE))

# Get the means for Surveys II and IV
M1 <- summary_stats$Mean[summary_stats$Survey == "II"]
M2 <- summary_stats$Mean[summary_stats$Survey == "IV"]

# Add mean and residuals to the data frame
WoP_Data <- WP_Data %>%
  mutate(mean = ifelse(Survey == "II", M1, M2),
         Residual = HAP - mean)

# Filter out NA values for 'HAP'
valid_data <- WP_Data %>%
  filter(!is.na(HAP))

# Extract data for Wilcoxon test
x <- valid_data$HAP[valid_data$Survey == "II"]
y <- valid_data$HAP[valid_data$Survey == "IV"]

# Perform paired Wilcoxon test if both groups have data
if (length(x) > 0 && length(y) > 0 && length(x) == length(y)) {
  wilcox_test <- wilcox.test(x, y, distribution = "asymptotic", alternative = "less", paired = TRUE)
  
  # Calculate effect size
  effect_size <- wilcoxonPairedR(x = c(x, y), g = rep(c("II", "IV"), times = c(length(x), length(y))), ci = TRUE)
  
  # Create a results data frame
  results2 <- data.frame(
    Column = 'HAP',
    Survey = c("II", "IV"),
    Median = summary_stats$Median,
    Mean = c(mean(x, na.rm = TRUE), mean(y, na.rm = TRUE)),
    SD = summary_stats$SD,
    p = as.numeric(wilcox_test$p.value),
    Z = as.numeric(wilcox_test$statistic),
    r = effect_size$r
  )
} else {
  results2 <- data.frame(
    Column = 'HAP',
    Survey = c("II", "IV"),
    Median = summary_stats$Median,
    Mean = c(NA, NA),
    SD = summary_stats$SD,
    p = NA,
    Z = NA,
    r = NA
  )
}

# Save the results to a CSV file
write.csv(results2, file = "HAP_WP_Survey_II_vs_IV.csv", row.names = FALSE)
warnings()


######
##HAP_WoP Survey IV vs WP Survey IV
# Filter data for WP group and 'HAP' column only for Surveys IV
Survey_IV_Data <- Data %>% 
  filter(Survey == "IV" & (Group == "WoP" | Group == "WP")) %>% 
  select(Survey, Group, HAP)

# Calculate mean, median, and standard deviation for 'HAP' within WoP group
summary_stats <- Survey_IV_Data %>%
  group_by(Group) %>%
  summarize(Mean = mean(HAP, na.rm = TRUE),
            Median = median(HAP, na.rm = TRUE),
            SD = sd(HAP, na.rm = TRUE))

# Get the means for Surveys IV WoP and WP
M1 <- summary_stats$Mean[summary_stats$Group == "WoP"]
M2 <- summary_stats$Mean[summary_stats$Group == "WP"]

# Add mean and residuals to the data frame
Survey_IV_Data <- Survey_IV_Data %>%
  mutate(mean = ifelse(Group == "WoP", M1, M2),
         Residual = HAP - mean)

# Filter out NA values for 'HAP'
valid_data <- Survey_IV_Data %>%
  filter(!is.na(HAP))

# Extract data for Wilcoxon test
x <- valid_data$HAP[valid_data$Group == "WoP"]
y <- valid_data$HAP[valid_data$Group == "WP"]

# Perform unpaired Wilcoxon test if both groups have data
if (length(x) > 0 && length(y) > 0 && length(x) == length(y)) {
  wilcox_test <- wilcox.test(x, y, distribution = "asymptotic", alternative = "less", paired = FALSE)
  
  # Calculate effect size
  effect_size <- wilcoxonR(x = c(x, y), g = rep(c("WoP", "WP"), times = c(length(x), length(y))), ci = TRUE)
  
  # Create a results data frame
  results3 <- data.frame(
    Column = 'HAP',
    Group = c("WoP", "WP"),
    Median = summary_stats$Median,
    Mean = c(mean(x, na.rm = TRUE), mean(y, na.rm = TRUE)),
    SD = summary_stats$SD,
    p = as.numeric(wilcox_test$p.value),
    Z = as.numeric(wilcox_test$statistic),
    r = effect_size$r
  )
} else {
  results3 <- data.frame(
    Column = 'HAP',
    Group = c("WoP", "WP"),
    Median = summary_stats$Median,
    Mean = c(NA, NA),
    SD = summary_stats$SD,
    p = NA,
    Z = NA,
    r = NA
  )
}

# Save the results to a CSV file
write.csv(results3, file = "HAP_Survey IV WoP vs WP.csv", row.names = FALSE)
