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


##EffortLevelVeryLow_WoP Survey III vs WP Survey III
# Filter data for WP group and 'EffortLevelVeryLow' column only for Surveys III
Survey_III_Data <- Data %>% 
  filter(Survey == "III" & (Group == "WoP" | Group == "WP")) %>% 
  select(Survey, Group, EffortLevelVeryLow)

# Calculate mean, median, and standard deviation for 'EffortLevelVeryLow' within WoP group
summary_stats <- Survey_III_Data %>%
  group_by(Group) %>%
  summarize(Mean = mean(EffortLevelVeryLow, na.rm = TRUE),
            Median = median(EffortLevelVeryLow, na.rm = TRUE),
            SD = sd(EffortLevelVeryLow, na.rm = TRUE))

# Get the means for Surveys III WoP and WP
M1 <- summary_stats$Mean[summary_stats$Group == "WoP"]
M2 <- summary_stats$Mean[summary_stats$Group == "WP"]

# Add mean and residuals to the data frame
Survey_III_Data <- Survey_III_Data %>%
  mutate(mean = ifelse(Group == "WoP", M1, M2),
         Residual = EffortLevelVeryLow - mean)

# Filter out NA values for 'EffortLevelVeryLow'
valid_data <- Survey_III_Data %>%
  filter(!is.na(EffortLevelVeryLow))

# Extract data for Wilcoxon test
x <- valid_data$EffortLevelVeryLow[valid_data$Group == "WoP"]
y <- valid_data$EffortLevelVeryLow[valid_data$Group == "WP"]

# Perform unpaired Wilcoxon test if both groups have data
if (length(x) > 0 && length(y) > 0 && length(x) == length(y)) {
  wilcox_test <- wilcox.test(x, y, distribution = "asymptotic", alternative = "less", paired = FALSE)
  
  # Calculate effect size
  effect_size <- wilcoxonR(x = c(x, y), g = rep(c("WoP", "WP"), times = c(length(x), length(y))), ci = TRUE)
  
  # Create a results data frame
  results <- data.frame(
    Column = 'EffortLevelVeryLow',
    Group = c("WoP", "WP"),
    Median = summary_stats$Median,
    Mean = c(mean(x, na.rm = TRUE), mean(y, na.rm = TRUE)),
    SD = summary_stats$SD,
    p = as.numeric(wilcox_test$p.value),
    Z = as.numeric(wilcox_test$statistic),
    r = effect_size$r
  )
} else {
  results <- data.frame(
    Column = 'EffortLevelVeryLow',
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
write.csv(results, file = "EffortLevelVeryLow_Survey III WoP vs WP.csv", row.names = FALSE)
warnings()