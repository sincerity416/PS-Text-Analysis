# =============================================================================
# Psychological Safety Literature Review - Tidy Text Analysis
# Bigrams, TF-IDF, Network Visualization, Yearly Trends
#
# 프로젝트 구조 (모든 경로는 BASE_DIR 기준):
#   PS Text Analysis/
#   ├── data/
#   │   ├── raw_text/   <- PDF에서 추출한 .txt 캐시 파일
#   │   └── results/    <- 분석 결과 CSV
#   ├── figures/        <- 시각화 PNG
#   └── scripts/        <- 이 스크립트
# =============================================================================

# ── 0. 패키지 설치 및 로드 ──────────────────────────────────────────────────
packages <- c("pdftools", "tidyverse", "tidytext", "widyr", "igraph", "ggraph", "ggrepel")

new_packages <- packages[!packages %in% installed.packages()[,"Package"]]
if (length(new_packages)) install.packages(new_packages)

library(pdftools)
library(tidyverse)
library(tidytext)
library(widyr)
library(igraph)
library(ggraph)
library(ggrepel)

# ── 1. 경로 설정 ──────────────────────────────────────────────────────────────
BASE_DIR   <- "~/Library/CloudStorage/Dropbox/Claude Code/PS Text Analysis"
PDF_DIR    <- "~/Library/CloudStorage/Dropbox/00 USM/Research/R Studio/Data/PS_PDFfiles"
CACHE_DIR  <- file.path(BASE_DIR, "data", "raw_text")
OUTPUT_DIR <- file.path(BASE_DIR, "data", "results")
FIG_DIR    <- file.path(BASE_DIR, "figures")

dir.create(CACHE_DIR,  showWarnings = FALSE, recursive = TRUE)
dir.create(OUTPUT_DIR, showWarnings = FALSE, recursive = TRUE)
dir.create(FIG_DIR,    showWarnings = FALSE, recursive = TRUE)

cat("=== 경로 설정 ===\n")
cat("  PDF 원본:    ", PDF_DIR,    "\n")
cat("  텍스트 캐시: ", CACHE_DIR,  "\n")
cat("  결과 CSV:    ", OUTPUT_DIR, "\n")
cat("  그래프:      ", FIG_DIR,    "\n\n")

# ── 2. PDF -> 텍스트 캐싱 (최초 1회만 파싱, 이후 .txt 재사용) ──────────────
pdf_files <- list.files(PDF_DIR, pattern = "\\.pdf$", full.names = TRUE)

cat("PDF 총", length(pdf_files), "개 발견\n")

walk(pdf_files, function(pdf_path) {
  txt_path <- file.path(
    CACHE_DIR,
    paste0(tools::file_path_sans_ext(basename(pdf_path)), ".txt")
  )
  if (!file.exists(txt_path)) {
    cat("  파싱 중:", basename(pdf_path), "\n")
    tryCatch({
      text <- paste(pdf_text(pdf_path), collapse = "\n")
      writeLines(text, txt_path)
    }, error = function(e) {
      warning("파싱 실패: ", basename(pdf_path), " - ", e$message)
    })
  }
})

cat("캐싱 완료 (", length(list.files(CACHE_DIR, "*.txt")), "개 txt 파일)\n\n")

# ── 3. 캐시된 텍스트 로드 + 메타데이터 추출 ──────────────────────────────────
txt_files <- list.files(CACHE_DIR, pattern = "\\.txt$", full.names = TRUE)

raw_texts <- map_dfr(txt_files, function(txt_path) {
  fname <- tools::file_path_sans_ext(basename(txt_path))
  year  <- as.integer(str_extract(fname, "\\d{4}"))
  tibble(
    doc_id = fname,
    year   = year,
    text   = read_file(txt_path)
  )
}) %>%
  mutate(
    text = str_to_lower(text),
    text = str_replace_all(text, "[^a-z\\s]", " "),
    text = str_squish(text)
  )

cat("로드된 논문 수:", nrow(raw_texts), "\n")
cat("연도 범위:", min(raw_texts$year, na.rm = TRUE),
    "~", max(raw_texts$year, na.rm = TRUE), "\n\n")

# ── 4. BIGRAM 분석 ────────────────────────────────────────────────────────────
data("stop_words")

custom_stopwords <- tibble(word = c(
  "al", "et", "pp", "vol", "doi", "http", "www", "org",
  "journal", "research", "study", "paper", "results",
  "table", "figure", "appendix", "na", "de", "en",
  # URL / 출판사 노이즈 (참고문헌 섹션)
  "https", "onlinelibrary", "wiley", "online", "library",
  "springer", "elsevier", "sage", "taylor", "francis",
  "cambridge", "oxford", "com", "edu", "pdf", "retrieved",
  # 단일 알파벳 노이즈
  "p", "r", "b", "e", "s", "t", "c", "d", "f", "g",
  # 일반 학술 불용어
  "article", "review", "findings", "data", "analysis",
  "measure", "sample", "scale", "item", "items", "scores",
  "model", "models", "hypothesis", "hypotheses", "coefficient"
))

all_stopwords <- bind_rows(stop_words, custom_stopwords)

bigrams_all <- raw_texts %>%
  unnest_tokens(bigram, text, token = "ngrams", n = 2) %>%
  separate(bigram, c("word1", "word2"), sep = " ") %>%
  filter(
    !word1 %in% all_stopwords$word,
    !word2 %in% all_stopwords$word,
    str_length(word1) > 1,
    str_length(word2) > 1
  ) %>%
  unite(bigram, word1, word2, sep = " ")

bigram_counts <- bigrams_all %>%
  count(bigram, sort = TRUE)

cat("=== TOP 30 BIGRAMS (전체 논문 합산) ===\n")
print(head(bigram_counts, 30))

# ── 5. TOP 20 BIGRAM 시각화 ───────────────────────────────────────────────────
p_bigram <- bigram_counts %>%
  head(20) %>%
  mutate(bigram = reorder(bigram, n)) %>%
  ggplot(aes(x = bigram, y = n, fill = n)) +
  geom_col(show.legend = FALSE) +
  coord_flip() +
  scale_fill_viridis_c(option = "plasma") +
  labs(
    title    = "Psychological Safety Literature Review",
    subtitle = "Top 20 Most Frequent Bigrams",
    x        = NULL,
    y        = "Frequency",
    caption  = paste0("Based on ", nrow(raw_texts), " papers")
  ) +
  theme_minimal(base_size = 13)

print(p_bigram)
ggsave(file.path(FIG_DIR, "PS_bigram_top20.png"), p_bigram, width = 10, height = 7, dpi = 150)

# ── 6. 논문별 BIGRAM 비교 (상위 5개) ─────────────────────────────────────────
bigrams_by_paper <- bigrams_all %>%
  count(doc_id, bigram, sort = TRUE) %>%
  group_by(doc_id) %>%
  slice_max(n, n = 5) %>%
  ungroup()

top_papers <- bigrams_all %>%
  count(doc_id) %>%
  slice_max(n, n = 8) %>%
  pull(doc_id)

p_by_paper <- bigrams_by_paper %>%
  filter(doc_id %in% top_papers) %>%
  mutate(
    doc_short = str_trunc(doc_id, 30),
    bigram    = reorder_within(bigram, n, doc_id)
  ) %>%
  ggplot(aes(x = bigram, y = n, fill = doc_short)) +
  geom_col(show.legend = FALSE) +
  facet_wrap(~doc_short, scales = "free_y", ncol = 2) +
  coord_flip() +
  scale_x_reordered() +
  labs(
    title    = "Top 5 Bigrams per Paper",
    subtitle = "(Top 8 papers by volume shown)",
    x = NULL, y = "Frequency"
  ) +
  theme_minimal(base_size = 10) +
  theme(strip.text = element_text(size = 7))

print(p_by_paper)
ggsave(file.path(FIG_DIR, "PS_bigram_by_paper.png"), p_by_paper, width = 14, height = 10, dpi = 150)

# ── 7. TF-IDF 분석 (논문마다 독특한 bigram 추출) ──────────────────────────────
tfidf_bigrams <- bigrams_all %>%
  count(doc_id, bigram, sort = TRUE) %>%
  bind_tf_idf(bigram, doc_id, n) %>%
  arrange(desc(tf_idf))

cat("\n=== TOP 20 TF-IDF BIGRAMS (논문별 특징어) ===\n")
print(head(tfidf_bigrams, 20))

p_tfidf <- tfidf_bigrams %>%
  filter(doc_id %in% top_papers) %>%
  group_by(doc_id) %>%
  slice_max(tf_idf, n = 5) %>%
  ungroup() %>%
  mutate(
    doc_short = str_trunc(doc_id, 30),
    bigram    = reorder_within(bigram, tf_idf, doc_id)
  ) %>%
  ggplot(aes(x = bigram, y = tf_idf, fill = doc_short)) +
  geom_col(show.legend = FALSE) +
  facet_wrap(~doc_short, scales = "free_y", ncol = 2) +
  coord_flip() +
  scale_x_reordered() +
  labs(
    title    = "TF-IDF: Paper-Specific Characteristic Bigrams",
    subtitle = "Higher TF-IDF = More unique to that paper",
    x = NULL, y = "TF-IDF Score"
  ) +
  theme_minimal(base_size = 10) +
  theme(strip.text = element_text(size = 7))

print(p_tfidf)
ggsave(file.path(FIG_DIR, "PS_tfidf_by_paper.png"), p_tfidf, width = 14, height = 10, dpi = 150)

# ── 8. 네트워크 시각화 (Bigram Network) ──────────────────────────────────────
set.seed(42)
TOP_N_BIGRAMS <- 60

bigram_graph <- bigram_counts %>%
  head(TOP_N_BIGRAMS) %>%
  separate(bigram, c("word1", "word2"), sep = " ") %>%
  graph_from_data_frame()

p_network <- ggraph(bigram_graph, layout = "fr") +
  geom_edge_link(
    aes(edge_alpha = n, edge_width = n),
    edge_colour = "#4a9eff",
    show.legend = FALSE
  ) +
  geom_node_point(color = "#ff6b6b", size = 3) +
  geom_node_text(
    aes(label = name),
    repel        = TRUE,
    size         = 3.5,
    color        = "black",
    fontface     = "bold",
    max.overlaps = 20
  ) +
  scale_edge_width(range = c(0.3, 2)) +
  labs(
    title    = "Psychological Safety - Bigram Network",
    subtitle = paste0("Top ", TOP_N_BIGRAMS, " bigrams across all papers"),
    caption  = "Node = word, Edge = bigram co-occurrence (thicker = more frequent)"
  ) +
  theme_graph(base_family = "sans", base_size = 12)

print(p_network)
ggsave(file.path(FIG_DIR, "PS_bigram_network.png"), p_network, width = 14, height = 10, dpi = 150)

# ── 9. 연도별 트렌드 ───────────────────────────────────────────────────────────
key_bigrams <- c(
  "psychological safety", "team learning", "learning behavior",
  "interpersonal risk", "team performance", "organizational learning",
  "leader behavior", "creative performance", "voice behavior",
  "work environment"
)

yearly_trend <- bigrams_all %>%
  filter(bigram %in% key_bigrams) %>%
  count(year, bigram) %>%
  filter(!is.na(year)) %>%
  complete(year, bigram, fill = list(n = 0))

n_bigrams    <- length(unique(yearly_trend$bigram))
color_values <- setNames(scales::hue_pal()(n_bigrams), unique(yearly_trend$bigram))

p_trend <- yearly_trend %>%
  ggplot(aes(x = year, y = n, color = bigram, group = bigram)) +
  geom_line(linewidth = 1) +
  geom_point(size = 2) +
  scale_color_manual(values = color_values) +
  labs(
    title    = "Key Bigram Frequency Over Time",
    subtitle = "How core PS concepts appear across publication years",
    x        = "Year",
    y        = "Frequency",
    color    = "Bigram",
    caption  = "Note: Each dot = one paper published that year"
  ) +
  theme_minimal(base_size = 12) +
  theme(
    legend.position = "bottom",
    legend.text     = element_text(size = 9)
  )

print(p_trend)
ggsave(file.path(FIG_DIR, "PS_yearly_trend.png"), p_trend, width = 12, height = 7, dpi = 150)

# ── 10. 결과 CSV 저장 ──────────────────────────────────────────────────────────
write_csv(bigram_counts,    file.path(OUTPUT_DIR, "PS_bigram_counts.csv"))
write_csv(tfidf_bigrams,    file.path(OUTPUT_DIR, "PS_tfidf_bigrams.csv"))
write_csv(bigrams_by_paper, file.path(OUTPUT_DIR, "PS_bigrams_by_paper.csv"))
write_csv(yearly_trend,     file.path(OUTPUT_DIR, "PS_yearly_trend.csv"))

cat("\n=== 분석 완료 ===\n")
cat("저장 위치: ~/Dropbox/Claude Code/PS Text Analysis/\n")
cat("  data/raw_text/ :", length(list.files(CACHE_DIR)), "개 txt\n")
cat("  data/results/  : PS_bigram_counts.csv, PS_tfidf_bigrams.csv\n")
cat("                   PS_bigrams_by_paper.csv, PS_yearly_trend.csv\n")
cat("  figures/       : PS_bigram_top20.png, PS_bigram_by_paper.png\n")
cat("                   PS_tfidf_by_paper.png, PS_bigram_network.png\n")
cat("                   PS_yearly_trend.png\n")
