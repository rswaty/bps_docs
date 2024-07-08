

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

## make a list of bold phrases

doc_for_bold <- "test_docs/short_docs/13023_51-short.docx"

extract_




