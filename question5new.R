library(readxl)
library(rstudioapi)

# Set working directory
setwd("G:/Math assignment")

# Verify if the file exists
file_name <- "United Airlines Aircraft Operating Statistics- Cost Per Block Hour (Unadjusted)"
file_xlsx <- paste0(file_name, ".xlsx")
file_xls <- paste0(file_name, ".xls")

if (file.exists(file_xlsx)) {
  data_file <- file_xlsx
} else if (file.exists(file_xls)) {
  data_file <- file_xls
} else {
  stop("File not found.")
}

# Read the Excel file with the specified range
all_data <- read_excel(data_file, range = "B2:W158")

# Define categories
daily_utilization_categories <- c("Block hours", "Airborne hours", "Departures")
ownership_categories <- c("Rental", "Depreciation and Amortization")
purchased_goods_categories <- c("Fuel/Oil", "Insurance", "Other (inc. Tax)")
fleet_category <- c(
  "small narrowbodies",
  "large narrowbodies",
  "widebodies",
  "total fleet"
)

# Define row numbers
purchased_goods_rows <- c(11, 50, 89, 128)  # Adjusted to match original `- 5` logic
ownership_rows <- purchased_goods_rows + 12
daily_utilization_rows <- ownership_rows + 13

# Function to extract row data
get_data_by_row <- function(row_num) {
  if (row_num > nrow(all_data)) {
    stop("Row number exceeds data range.")
  }
  return(na.omit(as.numeric(all_data[row_num, -1])))
}

# Function to create category data
get_category_data <- function(row_num, categories) {
  rows_data <- lapply(
    seq_along(categories),
    function(i) get_data_by_row(row_num + i)
  )
  costs <- unlist(rows_data)
  category <- factor(rep(categories, sapply(rows_data, length)))
  return(data.frame(costs = costs, category = category))
}

# Function to create a box plot
box_plot <- function(data, title, ylab) {
  boxplot(costs ~ category,
          data = data,
          main = title,
          col = "orange",
          ylab = ylab,
          border = "red")
}

# Function to plot category data
plot_category <- function(rows, categories, title, ylab) {
  windows(width = 1920 / 100, height = 1080 / 100)  # Set window size
  par(mfrow = c(2, 2), oma = c(0, 0, 3, 0))        # Adjust layout
  lapply(
    seq_along(rows),
    function(i) {
      box_plot(
        get_category_data(
          rows[i], categories
        ), fleet_category[i], ylab
      )
    }
  )
  mtext(title, outer = TRUE, cex = 1.5)            # Add title
  par(mfrow = c(1, 1))                             # Reset layout
}

# Plot categories
plot_category(
  purchased_goods_rows,
  purchased_goods_categories,
  "Purchased Goods",
  "Hours"
)

plot_category(
  ownership_rows,
  ownership_categories,
  "Aircraft Ownership",
  "Cost ($)"
)

plot_category(
  daily_utilization_rows,
  daily_utilization_categories,
  "Daily Utilization",
  "Cost ($)"
)
