

## NOTES -----------
# Want to get indicator species per BpS doc
# Several docs in one directory
# BpS Dominant and Indicator Species is the specific table, which is table 2

## PACKAGES ------------

# Load the dplyr package
library(dplyr)
library(docxtractr)

## SET DIRECTORY AND MAKE LIST OF DOCS ------------

# Specify the directory containing the docx files
docx_directory <- "extract_tables/input_docs/"

# List all docx files in the directory
docx_files <- list.files(path = docx_directory, pattern = "\\.docx$", full.names = TRUE)


## CREATE EMPTY LIST AND LOOP THROUGH ALL DOCUMENTS TO POPULATE LIST  ------------

# Initialize an empty list to store tables from each document
all_tables <- list()

# Iterate through each docx file
for (docx_file_path in docx_files) {
  # Read in the docx file
  docx_file <- read_docx(docx_file_path)
  
  docx_describe_tbls(docx_file)
  
  # Extract table from docx
  table_from_docx <- docx_extract_tbl(docx_file, tbl_number = 2)
  
  # Get the document name dynamically
  doc_name <- tools::file_path_sans_ext(basename(docx_file_path))
  
  # Add a new column for the document name
  table_from_docx$BPS_MODEL <- doc_name
  
  # Append the table to the list
  all_tables[[doc_name]] <- table_from_docx
}

## MAKE AND WRITE THE TABLE

# Combine all tables into a single dataframe
indicator_species <- bind_rows(all_tables)

write.csv(indicator_species, file = 'extract_tables/output_tables/indicator_species.csv', row.names = FALSE)
