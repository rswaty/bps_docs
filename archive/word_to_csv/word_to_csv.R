
# does not work as of June 27, 2024

# Load necessary libraries
library(officer)
library(dplyr)
library(readr)

# Directory containing Word documents
doc_directory <- "test_docs/two_docs/" # Replace with the actual directory path

# Function to read a Word document and extract headers and content
extract_content <- function(doc_path) {
  # Read the Word document
  doc <- read_docx(doc_path)
  
  # Extract all text from the document
  all_text <- docx_summary(doc)
  
  # Initialize an empty list to store headers and content
  content_list <- list()
  current_header <- NULL
  
  # Check if 'style_name' column exists in all_text
  if ("style_name" %in% names(all_text)) {
    # Loop through each paragraph in the document
    for (i in seq_len(nrow(all_text))) {
      # Check if the paragraph is a header based on style or formatting
      if (!is.na(all_text$style_name[i]) && grepl("heading", tolower(all_text$style_name[i]))) {
        # The current paragraph is a header
        current_header <- all_text$text[i]
        content_list[[current_header]] <- ""
      } else if (!is.null(current_header)) {
        # The current paragraph is content, append it to the current header
        content_list[[current_header]] <- paste(content_list[[current_header]], all_text$text[i], sep = " ")
      }
    }
  }
  
  # Return the content as a named vector, ensuring that the vector is not empty
  if (length(content_list) > 0) {
    return(unlist(content_list))
  } else {
    return(character(0))  # Return an empty character vector if no content was found
  }
}

# Automatically list all .docx files in the directory
doc_paths <- list.files(path = doc_directory, pattern = "\\.docx$", full.names = TRUE)

# Initialize a list to store all headers from all documents
all_headers <- list()

# Loop through each document to collect all unique headers
for (path in doc_paths) {
  content_vector <- extract_content(path)
  all_headers <- union(all_headers, names(content_vector))
}
