# Install and load required packages
library(readxl)
library(dplyr)
library(lubridate)
library(forecast)
library(openxlsx)  # For writing to Excel
library(ggplot2)   # For ggplot2 visualizations
library(magick)    # For image manipulation

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

# Function to plot and save both forecasts, then combine into one image
plot_forecasts_and_save <- function(country_name, models, forecast_periods, data) {
  # Sanitize the country_name to remove invalid characters for the file name
  sanitized_country_name <- gsub("[^[:alnum:]_]", "_", country_name)
  
  # Create new x-axis that extends the original dates by 52 weeks
  last_week <- max(data$`Week start date`)
  extended_weeks <- seq(last_week + weeks(1), by = "week", length.out = forecast_periods)
  
  # Combine observed and predicted data for export
  observed_data <- data.frame(Date = data$`Week start date`, Observed = data$`Influenza positive`)
  
  predicted_hw <- if (!is.null(models$Holt_Winters)) {
    hw_forecast <- forecast(models$Holt_Winters, h = forecast_periods)
    hw_forecast$mean
  } else {
    rep(NA, forecast_periods)
  }
  
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
    HW_Predicted = c(rep(NA, nrow(data)), predicted_hw),
    SARIMA_Predicted = c(rep(NA, nrow(data)), predicted_sarima)
  )
  
  # Write the combined data to the Excel workbook
  sheet_name <- sanitized_country_name
  addWorksheet(wb_combined, sheet_name)
  writeData(wb_combined, sheet_name, combined_data, startCol = 1, startRow = 1, colNames = TRUE)
  
  # Plot Holt-Winters forecast
  hw_plot <- ggplot(data = combined_data, aes(x = as.Date(Date))) +
    geom_line(aes(y = Observed, color = "Actual"), size = 1) +
    geom_line(aes(y = HW_Predicted, color = "Holt-Winters Forecast"), size = 1, linetype = "dashed") +
    labs(title = paste("Holt-Winters Forecast for", country_name),
         x = "Weeks", y = "Influenza Cases") +
    scale_color_manual(values = c("Actual" = "grey", "Holt-Winters Forecast" = "red")) +
    theme_minimal() +
    theme(plot.background = element_rect(fill = "white"))
  
  hw_filename <- paste0(output_path, "/", sanitized_country_name, "_holt_winters.png")
  ggsave(filename = hw_filename, plot = hw_plot, width = 8, height = 6)
  
  # Plot SARIMA forecast
  sarima_plot <- ggplot(data = combined_data, aes(x = as.Date(Date))) +
    geom_line(aes(y = Observed, color = "Actual"), size = 1) +
    geom_line(aes(y = SARIMA_Predicted, color = "SARIMA Forecast"), size = 1, linetype = "dashed") +
    labs(title = paste("SARIMA Forecast for", country_name),
         x = "Weeks", y = "Influenza Cases") +
    scale_color_manual(values = c("Actual" = "grey", "SARIMA Forecast" = "blue")) +
    theme_minimal() +
    theme(plot.background = element_rect(fill = "white"))
  
  sarima_filename <- paste0(output_path, "/", sanitized_country_name, "_sarima.png")
  ggsave(filename = sarima_filename, plot = sarima_plot, width = 8, height = 6)
  
  # Combine the two images into one
  hw_image <- image_read(hw_filename)
  sarima_image <- image_read(sarima_filename)
  
  combined_image <- image_append(c(hw_image, sarima_image), stack = TRUE)  # Combine vertically
  combined_image_filename <- paste0(output_path, "/", sanitized_country_name, "_combined_forecasts.png")
  image_write(combined_image, path = combined_image_filename)
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
    cat("Error processing", country, ":", e$message, "\n")  # Log any errors
  })
}

# Save the combined workbook
saveWorkbook(wb_combined, file = "/Users/turkialmalki/Desktop/Influenza prediction model/Influenza prediction/forecasts/combined_forecasts.xlsx", overwrite = TRUE)