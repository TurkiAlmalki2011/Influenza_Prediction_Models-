# Install and load required packages
library(readxl)
library(dplyr)
library(lubridate)
library(forecast)
library(openxlsx)  # For writing to Excel
library(ggplot2)   # For ggplot2 visualizations

# Load the Excel file
file_path <- "/Users/turkialmalki/Desktop/Influenza prediction model/Influenza data for all countires.xlsx"

# List the sheet names (each representing a different country)
sheet_names <- excel_sheets(file_path)

# Exclude the sheet for Kuwait
sheet_names <- sheet_names[sheet_names != "Kuwait"]

# Function to read and process data for each sheet
process_country_data <- function(sheet_name) {
  data <- read_excel(file_path, sheet = sheet_name)
  
  # Rename the column "Date" to "Week start date"
  data <- rename(data, `Week start date` = `Date`)
  
  # Convert "Week start date" to Date format and aggregate data by date
  data <- data %>%
    mutate(`Week start date` = as.Date(`Week start date`, format = "%Y-%m-%d")) %>%
    group_by(`Week start date`) %>%
    summarise(
      `Specimen tested` = sum(`Specimen tested`, na.rm = TRUE),
      `Influenza positive` = sum(`Influenza positive`, na.rm = TRUE)
    ) %>%
    arrange(`Week start date`)  # Sort data by date
  
  return(data)
}

# Load all the country data into a list (excluding Kuwait)
country_data_list <- lapply(sheet_names, process_country_data)
names(country_data_list) <- sheet_names  # Name the list elements based on sheet names

# Function to fit SARIMA model with Fourier terms and Holt-Winters models
fit_models <- function(data) {
  n_points <- nrow(data)
  
  # Check for sufficient data points
  if (n_points < 12) {
    stop("Not enough data to model this country")
  }
  
  freq <- if (n_points >= 52) 52 else 12  # Weekly or monthly data
  
  # Dynamically calculate K based on the frequency (K must be <= frequency/2)
  K <- min(2, floor(freq / 2), floor(n_points / 2))
  
  # Convert the "Influenza positive" column to a time series object
  ts_data <- ts(data$`Influenza positive`, start = c(year(min(data$`Week start date`)), 1), frequency = freq)
  
  # Fit Holt-Winters model with error handling
  hw_model <- tryCatch(HoltWinters(ts_data), error = function(e) NULL)
  
  # Create Fourier terms to represent seasonality for the existing data
  fourier_terms <- fourier(ts_data, K = K)
  
  # Fit SARIMA model with Fourier terms and error handling
  sarima_model <- tryCatch(auto.arima(ts_data, xreg = fourier_terms, seasonal = TRUE), error = function(e) NULL)
  
  return(list(Holt_Winters = hw_model, SARIMA = sarima_model, Fourier_Terms = fourier_terms, ts_data = ts_data, K = K))
}

# Initialize the workbook to store both predictions
wb_combined <- createWorkbook()

# Define the output directory for the PNG files
output_path <- "/Users/turkialmalki/Desktop/Influenza prediction model/Influenza prediction/forecasts/"

# Ensure the output directory exists
if (!dir.exists(output_path)) {
  dir.create(output_path, recursive = TRUE)
}

# Function to forecast, save PNGs, and write both predictions to Excel
plot_forecasts_and_save <- function(country_name, models, forecast_periods, data) {
  # Sanitize the country_name to remove invalid characters for the file name
  sanitized_country_name <- gsub("[^[:alnum:]_]", "_", country_name)
  
  # Define file name for PNG
  png_filename <- paste0(output_path, "/", sanitized_country_name, "_forecast.png")
  
  # Create new x-axis that extends the original dates by 52 weeks
  last_week <- max(data$`Week start date`)
  extended_weeks <- seq(last_week + weeks(1), by = "week", length.out = forecast_periods)
  
  # Combine observed and predicted data for export
  observed_data <- data.frame(Date = data$`Week start date`, Observed = data$`Influenza positive`)
  predicted_sarima <- if (!is.null(models$SARIMA)) {
    fourier_forecast_terms <- fourier(models$ts_data, K = models$K, h = forecast_periods)
    sarima_forecast <- forecast(models$SARIMA, h = forecast_periods, xreg = fourier_forecast_terms, level = c(95))
    sarima_forecast$mean[sarima_forecast$mean < 0] <- 0  # Ensure non-negative predictions
    sarima_forecast$mean
  } else {
    rep(NA, forecast_periods)
  }
  
  # Create a dataframe for all values
  combined_data <- data.frame(
    Date = c(as.character(data$`Week start date`), as.character(extended_weeks)),
    Observed = c(data$`Influenza positive`, rep(NA, forecast_periods)),
    SARIMA_Predicted = c(rep(NA, nrow(data)), predicted_sarima)
  )
  
  # Write the combined data to the Excel workbook
  sheet_name <- sanitized_country_name
  addWorksheet(wb_combined, sheet_name)
  writeData(wb_combined, sheet_name, combined_data, startCol = 1, startRow = 1, colNames = TRUE)
  
  # Plot using ggplot2, including all years from 2015 to August 2025
  gg_plot <- ggplot(data = combined_data, aes(x = as.Date(Date))) +
    geom_line(aes(y = Observed, color = "Observed"), size = 0.5, na.rm = TRUE) +  # Do not plot missing values (na.rm = TRUE)
    geom_line(aes(y = SARIMA_Predicted, color = "Forecast"), size = 0.5, linetype = "longdash", na.rm = TRUE) +  # Do not plot missing values (na.rm = TRUE)
    labs(x = "Years", y = "Influenza Positive Cases", color = "Positive Cases") +  # Changed legend title to "Positive Cases"
    scale_color_manual(values = c("Observed" = "black", "Forecast" = "blue")) +
    scale_x_date(date_breaks = "1 year", date_labels = "%Y", limits = as.Date(c("2015-01-01", "2025-08-31"))) +  # Adjust x-axis from 2015 to August 2025
    theme_minimal() +  
    theme(
      panel.background = element_rect(fill = "white", color = "black"),  # Keep the plot borders
      plot.background = element_rect(fill = "white", color = "black"),   # Keep the background borders
      axis.title.y = element_text(margin = margin(r = 15)),  # Move y-axis label to the right
      plot.title = element_text(hjust = 0.5, vjust = 1, size = 16),  # Center title above the plot
      plot.margin = margin(t = 80, r = 20, b = 20, l = 40)  # Add space for the title outside the plot
    ) +
    ggtitle(country_name)  # Add the country name as the title above the plot, outside the borders
  
  # Save the plot
  ggsave(filename = paste0(output_path, "/", country_name, "_forecast_plot.png"), plot = gg_plot, width = 8, height = 6)
  
  # Save the plot to PNG
  ggsave(filename = png_filename, plot = gg_plot, width = 8, height = 6)
}

# Apply the models to each country's data and save as PNG and Excel
for (country in names(country_data_list)) {
  cat("Processing:", country, "\n")  # Log country being processed
  
  data <- country_data_list[[country]]
  
  # Handle errors for individual countries
  tryCatch({
    # Use the last available date to start the forecast
    last_date <- max(data$`Week start date`)
    forecast_periods <- 52  # Forecast for the next year
    
    # Fit models
    models <- fit_models(data)
    
    # Plot and save forecasts to PNG and write to Excel
    plot_forecasts_and_save(country, models, forecast_periods, data)
  }, error = function(e) {
    cat("Error processing", country, ":", e$message, "\n")
  })
}

# Save the combined workbook with predictions for all countries
output_excel_file <- "/Users/turkialmalki/Desktop/Influenza prediction model/Influenza prediction/Influenza_predictions_combined.xlsx"
saveWorkbook(wb_combined, output_excel_file, overwrite = TRUE)

cat("Processing completed. All forecasts saved to Excel.")