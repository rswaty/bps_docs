

## chat tries

install.packages("officer")
install.packages("dplyr")

library(officer)
library(dplyr)
library(stringr)


## trial one -----
extract_sections <- function(doc_path, sections) {
  doc <- read_docx(doc_path)
  paragraphs <- docx_summary(doc)
  
  content_list <- lapply(sections, function(section) {
    section_index <- which(paragraphs$text == section)
    if (length(section_index) == 0) return(NA)
    content <- ""
    for (i in (section_index + 1):nrow(paragraphs)) {
      if (paragraphs$text[i] %in% sections) break
      content <- paste(content, paragraphs$text[i], sep = "\n")
    }
    return(trimws(content))  # Trim leading/trailing whitespace
  })
  
  names(content_list) <- sections
  return(as.data.frame(content_list))
}

directory_path <- "test_docs/short_docs/"
sections <- c("Vegetation Type", "Map Zones", "Model Splits or Lumps")

# Get all docx files in the directory
doc_paths <- list.files(directory_path, pattern = "\\.docx$", full.names = TRUE)

data_list <- lapply(doc_paths, function(doc_path) {
  extract_sections(doc_path, sections)
})

df <- bind_rows(data_list, .id = "document")
print(df)

# worked-now would like document names as first column

## try to add document names as first column ----

library(officer)
library(dplyr)
library(stringr)

# Define the directory path and sections
directory_path <- "test_docs/short_docs/"
sections <- c("Vegetation Type", 
              "Map Zones", 
              "Geographic Range",
              "Model Splits or Lumps")

# Function to extract sections from a Word document
extract_sections <- function(doc_path, sections) {
  doc <- read_docx(doc_path)
  paragraphs <- docx_summary(doc)
  
  content_list <- lapply(sections, function(section) {
    section_index <- which(paragraphs$text == section)
    if (length(section_index) == 0) return(NA)
    content <- ""
    for (i in (section_index + 1):nrow(paragraphs)) {
      if (paragraphs$text[i] %in% sections) break
      content <- paste(content, paragraphs$text[i], sep = "\n")
    }
    return(trimws(content))  # Trim leading/trailing whitespace
  })
  
  names(content_list) <- sections
  content_list <- as.data.frame(content_list)
  content_list$document <- basename(doc_path)
  return(content_list)
}

# Get all docx files in the directory
doc_paths <- list.files(directory_path, pattern = "\\.docx$", full.names = TRUE)

# Extract data from each document
data_list <- lapply(doc_paths, function(doc_path) {
  extract_sections(doc_path, sections)
})

# Combine the data into a single dataframe
df <- bind_rows(data_list)

# Move 'document' column to the first position
df <- df %>% select(document, everything())

# Print the dataframe
print(df)


## get bold?  -----

# Define the path to your document
doc_path <- "test_docs/short_docs/13023_51-short.docx"

# Read the document and inspect the structure
doc <- read_docx(doc_path)
content <- docx_summary(doc)

# Print the structure of the content
str(content)

# Want the 'Info_Para' phrases


# Function to extract paragraphs with the style "Info_Para"
extract_info_para <- function(doc_path) {
  doc <- read_docx(doc_path)
  content <- docx_summary(doc)
  
  # Filter paragraphs to get only the ones with the style "Info_Para"
  info_para_texts <- content %>%
    filter(style_name == "Info_Para") %>%
    select(text) %>%
    distinct()
  
  return(info_para_texts$text)
}

# Extract "Info_Para" text
info_para_phrases <- extract_info_para(doc_path)

# Print "Info_Para" phrases
print(info_para_phrases)


## try to extract the "Info_Para" sections ------


library(officer)
library(dplyr)
library(tidyr)

# Define the directory path
directory_path <- "test_docs/short_docs/"

# Function to extract "Info_Para" paragraphs and their content
extract_info_para <- function(doc_path) {
  doc <- read_docx(doc_path)
  content <- docx_summary(doc)
  
  # Extract paragraphs with the style "Info_Para" and their content
  info_para_indices <- which(content$style_name == "Info_Para")
  if (length(info_para_indices) == 0) return(NULL) # No "Info_Para" in this document
  
  info_para_texts <- content$text[info_para_indices]
  extracted_content <- sapply(seq_along(info_para_indices), function(i) {
    start <- info_para_indices[i] + 1
    end <- if (i < length(info_para_indices)) info_para_indices[i + 1] - 1 else nrow(content)
    paste(content$text[start:end], collapse = " ")
  })
  
  names(extracted_content) <- info_para_texts
  return(as.data.frame(t(extracted_content), stringsAsFactors = FALSE))
}

# Get all docx files in the directory
doc_paths <- list.files(directory_path, pattern = "\\.docx$", full.names = TRUE)

# Extract data from each document and combine into a single dataframe
data_list <- lapply(doc_paths, function(doc_path) {
  doc_content <- extract_info_para(doc_path)
  if (is.null(doc_content)) return(NULL) # Skip documents with no "Info_Para"
  doc_content$document <- basename(doc_path)
  return(doc_content)
})

# Combine all documents into a single dataframe
combined_df <- bind_rows(data_list)

# Reorder columns to move 'document' to the first position
combined_df <- combined_df %>% select(document, everything())

# Print the dataframe
print(combined_df)


## try to add in sclass description info -----
