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

# Filter data for WoP group and 'HA' column only for Surveys II and IV
WoP_Data <- Data %>% 
  filter(Group == "WoP" & (Survey == "II" | Survey == "IV")) %>% 
  select(Group, Survey, HA)

# Calculate mean, median, and standard deviation for 'HA' within WoP group
summary_stats <- WoP_Data %>%
  group_by(Survey) %>%
  summarize(Mean = mean(HA, na.rm = TRUE),
            Median = median(HA, na.rm = TRUE),
            SD = sd(HA, na.rm = TRUE))

# Get the means for Surveys II and IV
M1 <- summary_stats$Mean[summary_stats$Survey == "II"]
M2 <- summary_stats$Mean[summary_stats$Survey == "IV"]

# Add mean and residuals to the data frame
WoP_Data <- WoP_Data %>%
  mutate(mean = ifelse(Survey == "II", M1, M2),
         Residual = HA - mean)

# Filter out NA values for 'HA'
valid_data <- WoP_Data %>%
  filter(!is.na(HA))

# Extract data for Wilcoxon test
x <- valid_data$HA[valid_data$Survey == "II"]
y <- valid_data$HA[valid_data$Survey == "IV"]

# Perform paired Wilcoxon test if both groups have data
if (length(x) > 0 && length(y) > 0 && length(x) == length(y)) {
  wilcox_test <- wilcox.test(x, y, distribution = "asymptotic", alternative = "greater", paired = TRUE)
  
  # Calculate effect size
  effect_size <- wilcoxonPairedR(x = c(x, y), g = rep(c("II", "IV"), times = c(length(x), length(y))), ci = TRUE)
  
  # Create a results data frame
  results1 <- data.frame(
    Column = 'HA',
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
    Column = 'HA',
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
write.csv(results1, file = "HA_WoP_Survey_II_vs_IV.csv", row.names = FALSE)
warnings()

######
##HA_WP_Survey_II_vs_IV
# Filter data for WP group and 'HA' column only for Surveys II and IV
WP_Data <- Data %>% 
  filter(Group == "WP" & (Survey == "II" | Survey == "IV")) %>% 
  select(Group, Survey, HA)

# Calculate mean, median, and standard deviation for 'HA' within WoP group
summary_stats <- WP_Data %>%
  group_by(Survey) %>%
  summarize(Mean = mean(HA, na.rm = TRUE),
            Median = median(HA, na.rm = TRUE),
            SD = sd(HA, na.rm = TRUE))

# Get the means for Surveys II and IV
M1 <- summary_stats$Mean[summary_stats$Survey == "II"]
M2 <- summary_stats$Mean[summary_stats$Survey == "IV"]

# Add mean and residuals to the data frame
WoP_Data <- WP_Data %>%
  mutate(mean = ifelse(Survey == "II", M1, M2),
         Residual = HA - mean)

# Filter out NA values for 'HA'
valid_data <- WP_Data %>%
  filter(!is.na(HA))

# Extract data for Wilcoxon test
x <- valid_data$HA[valid_data$Survey == "II"]
y <- valid_data$HA[valid_data$Survey == "IV"]

# Perform paired Wilcoxon test if both groups have data
if (length(x) > 0 && length(y) > 0 && length(x) == length(y)) {
  wilcox_test <- wilcox.test(x, y, distribution = "asymptotic", alternative = "greater", paired = TRUE)
  
  # Calculate effect size
  effect_size <- wilcoxonPairedR(x = c(x, y), g = rep(c("II", "IV"), times = c(length(x), length(y))), ci = TRUE)
  
  # Create a results data frame
  results2 <- data.frame(
    Column = 'HA',
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
    Column = 'HA',
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
write.csv(results2, file = "HA_WP_Survey_II_vs_IV.csv", row.names = FALSE)
warnings()


######
##HA_WoP Survey IV vs WP Survey IV
# Filter data for WP group and 'HA' column only for Surveys IV
Survey_IV_Data <- Data %>% 
  filter(Survey == "IV" & (Group == "WoP" | Group == "WP")) %>% 
  select(Survey, Group, HA)

# Calculate mean, median, and standard deviation for 'HA' within WoP group
summary_stats <- Survey_IV_Data %>%
  group_by(Group) %>%
  summarize(Mean = mean(HA, na.rm = TRUE),
            Median = median(HA, na.rm = TRUE),
            SD = sd(HA, na.rm = TRUE))

# Get the means for Surveys IV WoP and WP
M1 <- summary_stats$Mean[summary_stats$Group == "WoP"]
M2 <- summary_stats$Mean[summary_stats$Group == "WP"]

# Add mean and residuals to the data frame
Survey_IV_Data <- Survey_IV_Data %>%
  mutate(mean = ifelse(Group == "WoP", M1, M2),
         Residual = HA - mean)

# Filter out NA values for 'HA'
valid_data <- Survey_IV_Data %>%
  filter(!is.na(HA))

# Extract data for Wilcoxon test
x <- valid_data$HA[valid_data$Group == "WoP"]
y <- valid_data$HA[valid_data$Group == "WP"]

# Perform unpaired Wilcoxon test if both groups have data
if (length(x) > 0 && length(y) > 0 && length(x) == length(y)) {
  wilcox_test <- wilcox.test(x, y, distribution = "asymptotic", alternative = "less", paired = FALSE)
  
  # Calculate effect size
  effect_size <- wilcoxonR(x = c(x, y), g = rep(c("WoP", "WP"), times = c(length(x), length(y))), ci = TRUE)
  
  # Create a results data frame
  results3 <- data.frame(
    Column = 'HA',
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
    Column = 'HA',
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
write.csv(results3, file = "HA_Survey IV WoP vs WP.csv", row.names = FALSE)
