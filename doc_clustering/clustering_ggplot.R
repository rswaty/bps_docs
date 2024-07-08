

## Background -----

# Goal is to test clustering of BpS docs by their corpus
# Searched https://landfirereview.org/search.php then downloaded all docs in MZs 50 and 51 (most of MI and WI) 
# Code by Randy Swaty and chat GPT
# March 7, 2024



# Install and load required packages

library(tm)
library(proxy)
library(ggplot2)
library(plotly)
library(igraph)
library(ggraph)
library(ggdendro)
library(dendextend)
library(ape)

# Read in multiple Word documents
doc_dir <- "doc_clustering/bps_docs_mzs_50_51/"
docs <- Corpus(DirSource(doc_dir, pattern = ".docx"))

# Preprocess the documents
docs <- tm_map(docs, content_transformer(tolower))
docs <- tm_map(docs, removePunctuation)
docs <- tm_map(docs, removeNumbers)
docs <- tm_map(docs, removeWords, stopwords("english"))
docs <- tm_map(docs, stemDocument)

# Create a document-term matrix (DTM)
dtm <- DocumentTermMatrix(docs)

# Calculate the cosine similarity matrix using proxy
similarity_matrix <- proxy::simil(as.matrix(dtm), method = "cosine")



# Create a hierarchical clustering tree
hc <- hclust(as.dist(1 - similarity_matrix), method = "complete")


ggdendrogram(hc)
# Rotate the plot and remove default theme
ggdendrogram(hc, rotate = TRUE, theme_dendro = FALSE)

# Build dendrogram object from hclust results
dend <- as.dendrogram(hc)
# Extract the data (for rectangular lines)
# Type can be "rectangle" or "triangle"
dend_data <- dendro_data(dend, type = "rectangle")
# What contains dend_data
names(dend_data)


ggplot() + 
  geom_segment(data=segment(dend_data), aes(x=x, y=y, xend=xend, yend=yend)) + 
  geom_text(data=label(dend_data), aes(x=x, y=y, label=label, hjust=0), size=3) +
  coord_flip() + scale_y_reverse(expand=c(0.2, 0.5)) + 
  theme_bw(base_size = 16)
