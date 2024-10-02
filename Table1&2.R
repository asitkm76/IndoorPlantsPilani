
# Load the necessary libraries
library(readxl)
library(dplyr)
# Set working directory
setwd("F:/Google drive data/Plants and IEQ Project/Main experiments/Data analysis")

# Read the Excel file into R
data <- read_excel("All data.xlsx")

# Calculate median, 25th percentile, and 75th percentile for the original and additional columns
summary_stats <- data %>%
  group_by(Group) %>%
  summarise(
    Height_median = median(Height.meter, na.rm = TRUE),
    Height_25th = quantile(Height.meter, 0.25, na.rm = TRUE),
    Height_75th = quantile(Height.meter, 0.75, na.rm = TRUE),
    
    Weight_median = median(Weight.kg, na.rm = TRUE),
    Weight_25th = quantile(Weight.kg, 0.25, na.rm = TRUE),
    Weight_75th = quantile(Weight.kg, 0.75, na.rm = TRUE),
    
    BMI_median = median(BMI.kg_m2, na.rm = TRUE),
    BMI_25th = quantile(BMI.kg_m2, 0.25, na.rm = TRUE),
    BMI_75th = quantile(BMI.kg_m2, 0.75, na.rm = TRUE),
    
    clo_median = median(clo, na.rm = TRUE),
    clo_25th = quantile(clo, 0.25, na.rm = TRUE),
    clo_75th = quantile(clo, 0.75, na.rm = TRUE),
    
    SleepScore_median = median(SleepScore, na.rm = TRUE),
    SleepScore_25th = quantile(SleepScore, 0.25, na.rm = TRUE),
    SleepScore_75th = quantile(SleepScore, 0.75, na.rm = TRUE),
    
    Age_median = median(Age.years, na.rm = TRUE),
    Age_25th = quantile(Age.years, 0.25, na.rm = TRUE),
    Age_75th = quantile(Age.years, 0.75, na.rm = TRUE),
    
    # New columns
    Rh_median = median(Rh.pcnt, na.rm = TRUE),
    Rh_25th = quantile(Rh.pcnt, 0.25, na.rm = TRUE),
    Rh_75th = quantile(Rh.pcnt, 0.75, na.rm = TRUE),
    
    AirTemp_median = median(AirTemp.C, na.rm = TRUE),
    AirTemp_25th = quantile(AirTemp.C, 0.25, na.rm = TRUE),
    AirTemp_75th = quantile(AirTemp.C, 0.75, na.rm = TRUE),
    
    GlobeTemp_median = median(GlobeTemp.C, na.rm = TRUE),
    GlobeTemp_25th = quantile(GlobeTemp.C, 0.25, na.rm = TRUE),
    GlobeTemp_75th = quantile(GlobeTemp.C, 0.75, na.rm = TRUE),
    
    CO2_median = median(CO2.ppm, na.rm = TRUE),
    CO2_25th = quantile(CO2.ppm, 0.25, na.rm = TRUE),
    CO2_75th = quantile(CO2.ppm, 0.75, na.rm = TRUE),
    
    AirSpeed_median = median(AirSpeed.m_s, na.rm = TRUE),
    AirSpeed_25th = quantile(AirSpeed.m_s, 0.25, na.rm = TRUE),
    AirSpeed_75th = quantile(AirSpeed.m_s, 0.75, na.rm = TRUE),
    
    OutdoorAirTemp_median = median(OutdoorAirTemp.C, na.rm = TRUE),
    OutdoorAirTemp_25th = quantile(OutdoorAirTemp.C, 0.25, na.rm = TRUE),
    OutdoorAirTemp_75th = quantile(OutdoorAirTemp.C, 0.75, na.rm = TRUE),
    
    OutdoorRH_median = median(OutdoorRH.pcnt, na.rm = TRUE),
    OutdoorRH_25th = quantile(OutdoorRH.pcnt, 0.25, na.rm = TRUE),
    OutdoorRH_75th = quantile(OutdoorRH.pcnt, 0.75, na.rm = TRUE),
    
    OutdoorAirSpeed_median = median(OutdoorAirSpeed.m_s, na.rm = TRUE),
    OutdoorAirSpeed_25th = quantile(OutdoorAirSpeed.m_s, 0.25, na.rm = TRUE),
    OutdoorAirSpeed_75th = quantile(OutdoorAirSpeed.m_s, 0.75, na.rm = TRUE),
    
    CalculatedVent_median = median(CalculatedVent.ACH, na.rm = TRUE),
    CalculatedVent_25th = quantile(CalculatedVent.ACH, 0.25, na.rm = TRUE),
    CalculatedVent_75th = quantile(CalculatedVent.ACH, 0.75, na.rm = TRUE)
  )

# View the result
print(summary_stats)

# Save the result to a CSV file
write.csv(summary_stats, "Table1&2 results.csv", row.names = FALSE)
