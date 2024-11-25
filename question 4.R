
library(readxl)
library(rstudioapi)

# Set working directory to the location of your Excel file
# Set the working directory to the folder where the file is located
setwd("G:\Math assignment")

# Verify if the file exists
file.exists("United Airlines Aircraft Operating Statistics- Cost Per Block Hour (Unadjusted).xls")
# If the file exists, this should return TRUE

# If the file extension is .xlsx, update the file name accordingly:
data_file <- "United Airlines Aircraft Operating Statistics- Cost Per Block Hour (Unadjusted).xlsx"

# If the file extension is .xls, use:
data_file <- "United Airlines Aircraft Operating Statistics- Cost Per Block Hour (Unadjusted).xls"

# Read the Excel file with the specified range
all_data <- read_excel(data_file, range = "B2:W158")
all_data

maintenance_categories <- c("labor", "materials", "third party", "burden")
years <- 1995:2015

# Function to extract data for a given row number (Maintenance/Load Factor)
get_data_by_row <- function(row_num) {
  return(na.omit(as.numeric(all_data[row_num, -1])))
}
get_maintenace_category <- function(row_num) {
  labor <- get_data_by_row(row_num + 1)
  materials <- get_data_by_row(row_num + 2)
  third_party <- get_data_by_row(row_num + 3)
  burden <- get_data_by_row(row_num + 5)
  return(setNames(
    c(sum(labor), sum(materials), sum(third_party), sum(burden)),
    maintenance_categories
  ))
}

# For ploting Load Factor bar plot
plot_bar <- function(data, title) {
  barplot(data,
          main = title,
          xlab = "Years",
          ylab = "Load Factor (%)",
          col = "lightblue",
          border = "red"
  )
}

# Maintenance and Load Factor row numbers
maintenance_rows <- c(16, 55, 94, 133)
load_factor_rows <- c(34, 73, 112, 151)

fleet_category <- c(
  "small narrowbodies",
  "large narrowbodies",
  "widebodies",
  "total fleet"
)

# Load necessary library
library(RColorBrewer)  # For color palettes

# Pie chart for maintenance
windows(width = 1920 / 100, height = 1080 / 100) # Set window size
par(mfrow = c(2, 2), oma = c(0, 0, 3, 0))

# Define a color palette
colors <- brewer.pal(4, "Set3")  # Using RColorBrewer for a set of 4 distinct colors

# Create pie charts for each maintenance category with enhancements
lapply(1:4, function(i) {
  data <- get_maintenace_category(maintenance_rows[i])
  
  # Calculate percentages
  percentages <- round(100 * data / sum(data), 1)
  
  # Create labels with category names and percentages
  labels <- paste0(names(data), ": ", percentages, "%")
  
  # Create pie chart
  pie(data, 
      labels = labels,        # Use labels with percentages
      main = paste("Maintenance Costs for", fleet_category[i]),  # Descriptive title
      col = colors,          # Set colors for slices
      border = "white")      # Add border to slices
})

# Add an outer title for all pie charts
mtext("Maintenance Cost Distribution", outer = TRUE, cex = 1.5)

# Reset plotting parameters to default
par(mfrow = c(1, 1))


# bar chart for load factor
windows(width = 1920 / 100, height = 1080 / 100) # Set window size
par(mfrow = c(2, 2), oma = c(0, 0, 3, 0))
lapply(1:4, function(i) {
  data <- setNames(get_data_by_row(load_factor_rows[i]), years)
  plot_bar(data, fleet_category[i])
})
mtext("Load Factor", outer = TRUE, cex = 1.5)
par(mfrow = c(1, 1))

