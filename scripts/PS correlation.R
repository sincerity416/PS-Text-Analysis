install.packages("pdftools")
install.packages("tidyverse")
install.packages("tidytext")
install.packages("readtext")
install.packages("widyr")
install.packages("igraph")
install.packages("ggraph")

library(pdftools)
library(tidyverse)
library(tidytext)
library(readtext)
library(widyr)
library(igraph)
library(ggraph)

setwd("C:/Users/shinh/Dropbox/00 USM/Research/R Studio/Data/PS_PDFfiles")

files <- list.files(pattern = "*.pdf")
files

text_data <- readtext("*.pdf")

head(text_data)

words <- text_data %>%
  unnest_tokens(word, text)

head(words)


data("stop_words")

clean_words <- words %>%
  anti_join(stop_words)

word_frequency <- clean_words %>%
  count(word, sort = TRUE)

head(word_frequency, 20)


bigrams <- text_data %>%
  unnest_tokens(bigram, text, token = "ngrams", n = 2)


bigrams_separated <- bigrams %>%
  separate(bigram, c("word1", "word2"), sep = " ")


bigrams_filtered <- bigrams_separated %>%
  filter(!word1 %in% stop_words$word) %>%
  filter(!word2 %in% stop_words$word)


bigrams_clean <- bigrams_filtered %>%
  unite(bigram, word1, word2, sep = " ")


bigram_counts <- bigrams_clean %>%
  count(bigram, sort = TRUE)

head(bigram_counts, 20)

head(bigrams)


word_correlation <- clean_words %>%
  pairwise_cor(word, doc_id, sort = TRUE)

head(word_correlation)


word_correlation %>%
  filter(correlation > 0.3) %>%
  graph_from_data_frame() %>%
  ggraph(layout = "fr") +
  geom_edge_link() +
  geom_node_point() +
  geom_node_text(aes(label = name), repel = TRUE)

