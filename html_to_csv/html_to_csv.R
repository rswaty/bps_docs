


# Load necessary libraries
library(rvest)
library(xml2)
library(dplyr)
library(readr)

# Define the directory where your files are located
doc_directory <- "test_docs/docs_as_html/"

# Get a list of HTML files in the directory
files <- list.files(path = doc_directory, pattern = "*.htm$")

# Initialize an empty list to store the sections
list_df <- list()

# Loop over each file
for (file in files) {
  # Read the HTML file
  html <- read_html(paste0(doc_directory, file))
  
  # Extract the sections denoted by bold text
  sections <- html %>% html_nodes('b , strong') %>% html_text()
  
  # Extract the content of each section
  content <- html %>% html_nodes('p') %>% html_text()
  
  # Create a data frame with sections as column names and content as values
  df <- data.frame(matrix(ncol = length(sections), nrow = 1))
  colnames(df) <- sections
  df[1, ] <- content
  
  # Add the data frame to the list
  list_df[[file]] <- df
}

# Combine all data frames in the list into one data frame
final_df <- bind_rows(list_df)

# Write the data frame to a CSV file
write_csv(final_df, "output.csv")
