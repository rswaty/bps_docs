

## try parsing with predefined column headings (from the docs)

## code altered from Copilot

## using short docs with fewer sections at first

## NOTES -----
# need to add code to list unparsed documents


# Install and load necessary packages

library(readtext)
library(tidyverse)

## first try ----

# Define the column names
column_names <- c("Vegetation Type", "Map Zones", "Model Splits or Lumps", "Geographic Range", 
                  "Biophysical Site Description", "Vegetation Description", "Class A", "Class B")

# Read the Word documents
docs <- readtext("test_docs/short_docs/*.docx")

# Initialize an empty dataframe with the desired column names
df <- data.frame(matrix(ncol = length(column_names), nrow = 0))
colnames(df) <- column_names

# Loop through each document
for(i in 1:nrow(docs)) {
  # Get the text of the document
  text <- docs$text[i]
  
  # Split the text into sections based on bold text (assuming it's surrounded by '**')
  sections <- strsplit(text, "\\*\\*.*?\\*\\*")[[1]]
  
  # If there are not enough sections, pad the vector with NA
  if(length(sections) < length(column_names)) {
    sections <- c(sections, rep(NA, length(column_names) - length(sections)))
  }
  
  # Add the sections as a new row in the dataframe
  df <- rbind(df, sections[1:length(column_names)])
}

# View the dataframe
print(df)


## trying to ignore all content before first pre-defined section

# Load necessary packages
library(readtext)
library(tidyverse)

# Define the column names
column_names <- c("Vegetation Type", "Map Zones", "Model Splits or Lumps", "Geographic Range", 
                  "Biophysical Site Description", "Vegetation Description", "Class A", "Class B")

# Read the Word documents
docs <- readtext("test_docs/short_docs/*.docx")

# Initialize an empty dataframe with the desired column names
df <- data.frame(matrix(ncol = length(column_names), nrow = 0))
colnames(df) <- column_names

# Loop through each document
for(i in 1:nrow(docs)) {
  # Get the text of the document
  text <- docs$text[i]
  
  # Ignore all text before the first pre-defined section
  text <- strsplit(text, "Vegetation Type", fixed = TRUE)[[1]][2]
  
  # Split the text into sections based on bold text (assuming it's surrounded by '**')
  sections <- strsplit(text, "\\*\\*.*?\\*\\*")[[1]]
  
  # If there are not enough sections, pad the vector with NA
  if(length(sections) < length(column_names)) {
    sections <- c(sections, rep(NA, length(column_names) - length(sections)))
  }
  
  # Add the sections as a new row in the dataframe
  df <- rbind(df, sections[1:length(column_names)])
}

# View the dataframe
print(df)


## Trying again. Uploaded document to co-pilot.  
# Load necessary packages
library(officer)
library(tidyverse)
library(stringr)

# Define the column names
column_names <- c("Map Zones", "Model Splits or Lumps", "Geographic Range", 
                  "Biophysical Site Description", "Vegetation Description", "Class A", "Class B")

# Get the list of Word documents
docs <- list.files(path = "test_docs/short_docs/", pattern = "\\.docx$", full.names = TRUE)

# Initialize an empty dataframe with the desired column names
df <- data.frame(matrix(ncol = length(column_names), nrow = 0))
colnames(df) <- column_names

# Loop through each document
for(doc in docs) {
  # Read the Word document
  doc_read <- read_docx(doc)
  
  # Get the text of the document
  text <- lapply(doc_read$content, function(x) { if(inherits(x, "block")) x$text })
  text <- unlist(text)
  
  # Initialize a list to store the sections
  sections <- list()
  
  # Loop through each column name
  for(j in 1:length(column_names)) {
    # Define the pattern for the section
    if(j < length(column_names)) {
      pattern <- paste0("(?<=", column_names[j], ").*?(?=", column_names[j + 1], ")")
    } else {
      pattern <- paste0("(?<=", column_names[j], ").*")
    }
    
    # Extract the section
    section <- str_extract(text, pattern)
    
    # If the section is not found or is of length zero, assign NA
    if(length(section) == 0 || is.na(section)) {
      section <- NA
    }
    
    # Add the section to the list
    sections[[j]] <- section
  }
  
  # Add the sections as a new row in the dataframe
  df <- rbind(df, setNames(sections, column_names))
}

# View the dataframe
print(df)
