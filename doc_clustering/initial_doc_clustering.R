

## use document clustering to find grouping options

## based on https://rstudio-pubs-static.s3.amazonaws.com/266040_d2920f956b9d4bd296e6464a5ccc92a1.html#:~:text=Here's%20a%20simplified%20description%20of,each%20of%20the%20old%20clusters.

## got bps docs by serching 'pine-oak' in MZs 41, 50, 51 

## packages
library(pacman)

p_load(tm, proxy, RTextTools, fpc, wordcloud, cluster, stringi, readtext)


## read in docs and create corpus
# Set your directory path
docx_folder <- "doc_clustering/pine_oak_41_50_51/"

# List all .docx files in the directory
docx_files <- list.files(path = docx_folder, pattern = ".docx", full.names = TRUE)

# Read the text from the .docx files
text_data <- readtext(docx_files)

# Create a Corpus using VCorpus
corpus <- Corpus(VectorSource(text_data$text))
summary(corpus)

# create Document Term Matrix

ndocs <- length(corpus)
# ignore extremely rare words i.e. terms that appear in less then 1% of the documents
minTermFreq <- ndocs * 0.01
# ignore overly common words i.e. terms that appear in more than 50% of the documents
maxTermFreq <- ndocs * .5
dtm = DocumentTermMatrix(corpus,
                         control = list(
                           stopwords = TRUE, 
                           wordLengths=c(4, 15),
                           removePunctuation = T,
                           removeNumbers = T,
                           #stemming = T,
                           bounds = list(global = c(minTermFreq, maxTermFreq))
                         ))

write.csv((as.matrix(dtm)), "test.csv")
dtm.matrix = as.matrix(dtm)

inspect(dtm)


m  <- as.matrix(dtm)
# # # m <- m[1:2, 1:3]
distMatrix <- dist(m, method="euclidean")
#print(distMatrix)
#distMatrix <- dist(m, method="cosine")
#print(distMatrix)

groups <- hclust(distMatrix,method="ward.D")
plot(groups, cex=0.9, hang=-1)
rect.hclust(groups, k=5)