

# Define the directory containing the Word documents and the directory to save modified documents
input_dir <- "test_docs/two_docs/"
output_dir <- "no_tables/"

library(officer)
library(purrr)
library(xml2)

# Define the function to remove tables from a Word document
remove_tables_from_doc <- function(input_path, output_path) {
  # Read the Word document
  doc <- read_docx(input_path)
  
  # Extract the document content as XML
  doc_xml <- docx_body_xml(doc)
  
  # Parse the XML content
  doc_xml <- read_xml(as.character(doc_xml))
  
  # Remove all table nodes
  tables <- xml_find_all(doc_xml, ".//w:tbl")
  xml_remove(tables)
  
  # Create a new Word document
  new_doc <- read_docx()
  
  # Add the modified content back to the new document
  new_doc <- body_add_xml(new_doc, as.character(doc_xml))
  
  # Save the new document
  print(new_doc, target = output_path)
}

# Get a list of all Word documents in the input directory
doc_files <- list.files(input_dir, pattern = "\\.docx$", full.names = TRUE)

# Define output file paths
output_files <- file.path(output_dir, basename(doc_files))

# Apply the function to all documents
purrr::map2(doc_files, output_files, remove_tables_from_doc)
