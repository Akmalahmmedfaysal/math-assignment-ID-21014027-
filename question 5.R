
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

# define categories
daily_utilization_categories <- c("Block hours", "Airborne hours", "Departures")
ownership_categories <- c("Rental", "Depreciation and Amortization")
purchased_goods_categories <- c("Fuel/Oil", "Insurance", "Other (inc. Tax)")
fleet_category <- c(
  "small narrowbodies",
  "large narrowbodies",
  "widebodies",
  "total fleet"
)

# row numbers
purchased_goods_rows <- c(16, 55, 94, 133) - 5
ownership_rows <- purchased_goods_rows + 12
daily_utilization_rows <- ownership_rows + 13

get_data_by_row <- function(row_num) {
  if (row_num > nrow(all_data)) {
    stop("Row number exceeds data range.")
  }
  return(na.omit(as.numeric(all_data[row_num, -1])))
}

get_category_data <- function(row_num, categories) {
  rows_data <- lapply(
    seq_along(categories),
    function(i) get_data_by_row(row_num + i)
  )
  costs <- unlist(rows_data)
  category <- factor(rep(categories, sapply(rows_data, length)))
  return(data.frame(costs = costs, category = category))
}

box_plot <- function(data, title, ylab) {
  boxplot(costs ~ category,
          data = data,
          main = title,
          col = "orange",
          ylab = ylab,
          border = "red"
  )
}

plot_category <- function(rows, categories, title, ylab) {
  windows(width = 1920 / 100, height = 1080 / 100) # Set window size
  par(mfrow = c(2, 2), oma = c(0, 0, 3, 0))
  # Create the box plot using the formula interface
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
  mtext(title, outer = TRUE, cex = 1.5)
  par(mfrow = c(1, 1))
}

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
