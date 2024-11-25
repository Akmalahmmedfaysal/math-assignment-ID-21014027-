
library(readxl)
library(rstudioapi)

# Set working directory to the location of your Excel file
# Set the working directory to the folder where the file is located
setwd("G:\Math assignment")

# Verify if the file exists
file.exists("C:/Users/saifu/OneDrive/Desktop/Math Assignment.xls")
# If the file exists, this should return TRUE

# If the file extension is .xlsx, update the file name accordingly:
data_file <- "United Airlines Aircraft Operating Statistics- Cost Per Block Hour (Unadjusted).xlsx"

# If the file extension is .xls, use:
data_file <- "United Airlines Aircraft Operating Statistics- Cost Per Block Hour (Unadjusted).xls"

# Read the Excel file with the specified range
all_data <- read_excel(data_file, range = "B2:W158")
all_data

# Helper function to extract salary data by row
get_salary_wages <- function(row_num, data = all_data) {
  return(na.omit(as.numeric(data[row_num, -1])))
}
get_salary_wages
# Ensure that salary_wages data has 23 points
# Extract salary data from the dataset using get_salary_wages()
salary_wages_snbodies <- get_salary_wages(6)    # For small narrowbodies
salary_wages_lnbodies <- get_salary_wages(45)   # For large narrowbodies
salary_wages_wbodies <- get_salary_wages(84)    # For widebodies
salary_wages_tfleet <- get_salary_wages(123)    # For total fleet

# Now, combine the extracted salary data into one sample
# Assuming you want a combined sample of all these salary data sets
salary_wages_sample <- c(salary_wages_snbodies, salary_wages_lnbodies, salary_wages_wbodies, salary_wages_tfleet)
salary_wages_sample
# Check the number of observations in the combined sample
length(salary_wages_sample)

# If you need exactly 27 observations, you can either take the first 23, sample 23 randomly, or apply some selection method.
# Randomly select 27 observations from the combined data
set.seed(123)  # For reproducibility
salary_wages_sample_27 <- sample(salary_wages_sample, 27, replace = FALSE)
salary_wages_sample_27
# View the sample of 28 observations
print(salary_wages_sample_27)

## Assuming these functions and salary data extraction methods have already been defined:
get_modes <- function(data) {
  freq_table <- table(data)
  max_freq <- max(freq_table)
  modes <- as.numeric(names(freq_table[freq_table == max_freq]))
  if (length(modes) == length(data)) {
    return(NULL)
  }
  return(modes)
}

get_frequency_distribution <- function(wage_data) {
  # Number of observations
  n <- length(wage_data)
  
  # Calculate k directly as log2(n), and round it up
  k <- ceiling(log2(n))
  
  # Calculate class interval (interval >= (max - min)/k)
  min_salary <- min(wage_data)
  max_salary <- max(wage_data)
  class_interval <- (max_salary - min_salary) / k
  class_interval <- ceiling(class_interval)  # Ensure class interval is a whole number
  
  # Create breakpoints
  break_points <- seq(
    min_salary - (class_interval / 2),  # Start the first break point slightly before the min value
    max_salary + (class_interval / 2),  # End the last break point slightly after the max value
    by = class_interval
  )
  
  # Create frequency distribution
  salary_bins <- cut(wage_data, breaks = break_points, right = TRUE)
  frequency_distribution <- table(salary_bins)
  
  return(frequency_distribution)
}

# Example: Extract salary data from each category
salary_wages_snbodies <- get_salary_wages(6)    # For small narrowbodies
salary_wages_lnbodies <- get_salary_wages(45)   # For large narrowbodies
salary_wages_wbodies <- get_salary_wages(84)    # For widebodies
salary_wages_tfleet <- get_salary_wages(123)    # For total fleet

# Combine the extracted salary data
combined_salary_wages <- c(salary_wages_snbodies, salary_wages_lnbodies, salary_wages_wbodies, salary_wages_tfleet)

# If you want to take exactly 28 samples, you can sample from combined data
set.seed(123)  # For reproducibility
salary_wages_sample_27 <- sample(combined_salary_wages, 27, replace = FALSE)

# Get frequency distribution for the sample data
frequency_distribution_sample <- get_frequency_distribution(salary_wages_sample_27)

# Print the frequency distribution for the sample
cat("Frequency Distribution for Sample of 27 Observations:\n")
print(frequency_distribution_sample)

# Perform analysis on the sample
print_analysis <- function(wage_data, title) {
  mean <- mean(wage_data)
  median <- median(wage_data)
  modes <- get_modes(wage_data)
  sample_sd <- sd(wage_data)  # sample
  sample_var <- var(wage_data)  # sample
  quartiles <- quantile(wage_data, probs = c(0.25, 0.5, 0.75))
  tenth_percentile <- quantile(wage_data, probs = 0.10)
  ninth_decile <- quantile(wage_data, probs = 0.90)
  range <- max(wage_data) - min(wage_data)
  
  # Print results
  cat("Analysis of ", title, "::\n")
  cat("Mean:", mean, "\n")
  cat("Median:", median, "\n")
  if (is.null(modes) || length(modes) == 0) {
    cat("Modes: None\n")
  } else {
    cat("Modes:", paste(modes, collapse = ", "), "\n")
  }
  
  cat("Sample Standard Deviation:", sample_sd, "\n")
  cat("Sample Variance:", sample_var, "\n")
  cat("Quartiles (Q1, Q2, Q3):", quartiles, "\n")
  cat("10th Percentile:", tenth_percentile, "\n")
  cat("9th Decile:", ninth_decile, "\n")
  cat("Range:", range, "\n")
  cat("\n\n")
}

# Perform analysis on the 27 sample
print_analysis(salary_wages_sample_27, "Salary Wages Sample of 27 Observations")

# Plot the histogram for the sample data
plot_histogram <- function(frequency_distribution, window_title) {
  barplot(frequency_distribution,
          xlab = "Salary Ranges",
          ylab = "Frequency",
          col = "lightblue",
          border = "black",
          space = 0,  # No space between bars
          width = 1,  # Adjust width to fill the space better
          main = window_title
  )
}

# Plot histogram for the sample frequency distribution
plot_histogram(frequency_distribution_sample, "Histogram of Salary Wages Sample of 27 Observations")
