# 2026-03-12 | PS Tidy Text 분석 설계 및 환경 구축

## 개요

심리적 안전감(Psychological Safety) 관련 논문 20편을 대상으로
R Tidy Text 분석을 설계하고, 분석 환경 및 GitHub 저장소를 구축한 작업 세션.

---

## 1. 문제 정의

### 배경
- 분석 대상: PDF 논문 20편 (`PS_PDFfiles/`)
- 목표: Bigram 빈도 분석을 통한 심리적 안전감 문헌 리뷰
- 우려 사항: PDF 20개 → 텍스트 용량이 많을 경우 컨텍스트/메모리 초과 가능성

### 규모 추정
| 항목 | 추정값 |
|------|--------|
| 논문 수 | 20개 |
| 논문당 평균 단어 수 | ~10,000단어 |
| 총 단어 수 | ~200,000단어 |
| 연도 범위 | 1999 ~ 2024 |

---

## 2. 의사결정 과정

### 2-1. 컨텍스트 초과 문제 해결 방안 검토

총 5가지 옵션을 비교 검토하였음.

| 옵션 | 방법 | 장점 | 단점 |
|------|------|------|------|
| Option 1 | 단순 일괄 처리 (Standard Tidytext) | 코드 단순, 즉시 실행 | 재실행 시 PDF 재파싱 |
| Option 2 | 전처리 캐싱 파이프라인 | 2회차부터 ~10배 빠름 | 디스크 사용 |
| Option 3 | 배치 처리 + 중간 저장 | 메모리 효율, 확장성 | 코드 복잡 |
| Option 4 | DuckDB 기반 SQL 쿼리 | 반복 쿼리 효율적 | 패키지 추가 필요 |
| Option 5 | RAG (의미 기반 검색) | 의미론적 쿼리 가능 | 과도한 복잡도 |

### 2-2. 최종 의사결정: Option 1 + Option 2 조합

**선택 근거:**
- 20개 논문은 R 메모리 관점에서 충분히 처리 가능한 규모
- 캐싱(Option 2)으로 재실행 효율성 확보
- 코드 단순성 유지 → 이후 분석 수정이 용이
- RAG는 bigram 빈도 분석 목적에 비해 과도한 복잡도

**구체적 구현:**
```r
# PDF → .txt 캐싱 (최초 1회만 파싱)
if (!file.exists(txt_path)) {
  text <- paste(pdf_text(pdf_path), collapse = "\n")
  writeLines(text, txt_path)
}
# 이후 실행은 .txt에서 직접 로드
```

### 2-3. 추가 분석 항목 결정

단순 bigram 빈도 외에 아래 4가지 분석을 추가하기로 결정:

| 분석 | 목적 |
|------|------|
| 논문별 Bigram 비교 | 각 논문의 핵심 개념 파악 |
| TF-IDF 분석 | 논문마다 독특한 특징어 추출 |
| 네트워크 시각화 | 단어 간 연결 관계 파악 |
| 연도별 트렌드 | 시간에 따른 개념 변화 추적 |

---

## 3. 분석 실행 결과

### Top 10 Bigrams (20개 논문 합산)

| 순위 | Bigram | 빈도 |
|------|--------|------|
| 1 | psychological safety | 3,080 |
| 2 | employee engagement | 631 |
| 3 | team learning | 200 |
| 4 | organizational learning | 192 |
| 5 | psychologically safe | 189 |
| 6 | team psychological | 187 |
| 7 | learning behavior | 154 |
| 8 | human resource | 143 |
| 9 | organizational behavior | 130 |
| 10 | team performance | 125 |

### 노이즈 제거 이슈 및 대응

1차 실행 결과 URL·출판사명이 bigram으로 추출되는 문제 발견:
- `wiley online`, `onlinelibrary wiley`, `https onlinelibrary` 등
- **대응:** custom_stopwords에 출판사·URL 관련 단어 추가 후 재실행

---

## 4. 작업 디렉토리 의사결정

### 경로 변경 이력

| 시점 | 경로 | 변경 이유 |
|------|------|-----------|
| 초기 | `R Studio/Data/Claude Code/` | 임시 생성 |
| 중간 | `R Studio/Data/PS Text Analysis/` | 학술 구조 적용 시도 |
| **최종** | `Dropbox/Claude Code/PS Text Analysis/` | 사용자 지정 최종 경로 |

### 최종 폴더 구조 (학술적 분류)

```
PS Text Analysis/
├── data/
│   ├── raw_text/     ← PDF에서 추출한 .txt 캐시 (20개)
│   └── results/      ← 분석 결과 CSV (4개)
├── figures/          ← 시각화 PNG (5개)
├── scripts/          ← R 분석 스크립트
├── discussion/       ← 작업 로그 및 의사결정 기록
└── CLAUDE.md         ← 세션 간 작업 기억 파일
```

---

## 5. GitHub 연동 의사결정

### 인증 방식 선택

| 방식 | 검토 결과 |
|------|-----------|
| HTTPS + PAT | 매번 토큰 관리 필요 |
| **SSH 키 (선택)** | 1회 설정 후 영구 사용 가능 |

**SSH 키 생성 및 등록 과정:**
1. `ssh-keygen -t ed25519` 로 키 생성
2. `cat ~/.ssh/id_ed25519.pub | pbcopy` 로 클립보드 복사 (공백 오류 방지)
3. GitHub Settings → SSH Keys 등록
4. `ssh-keyscan github.com >> ~/.ssh/known_hosts` 로 호스트 신뢰 등록
5. remote URL을 SSH 형식으로 변경 후 push 성공

> **주의:** GitHub 웹 UI에 키를 붙여넣을 때 채팅 복사본은 줄바꿈 오류 발생.
> 반드시 `pbcopy`로 클립보드에 직접 복사 후 붙여넣을 것.

### 저장소

- **GitHub:** https://github.com/sincerity416/PS-Text-Analysis
- **로컬:** `~/Library/CloudStorage/Dropbox/Claude Code/PS Text Analysis/`

---

## 6. 생성된 산출물 목록

| 구분 | 파일 | 설명 |
|------|------|------|
| 스크립트 | `scripts/PS_TidyText_Analysis.R` | 메인 분석 스크립트 |
| 스크립트 | `scripts/PS correlation.R` | 기존 상관분석 스크립트 |
| 데이터 | `data/raw_text/*.txt` | 논문 텍스트 캐시 20개 |
| 결과 | `data/results/PS_bigram_counts.csv` | 전체 bigram 빈도 |
| 결과 | `data/results/PS_tfidf_bigrams.csv` | TF-IDF 점수 |
| 결과 | `data/results/PS_bigrams_by_paper.csv` | 논문별 bigram |
| 결과 | `data/results/PS_yearly_trend.csv` | 연도별 트렌드 |
| 그래프 | `figures/PS_bigram_top20.png` | Top 20 bigram 막대그래프 |
| 그래프 | `figures/PS_bigram_by_paper.png` | 논문별 bigram 비교 |
| 그래프 | `figures/PS_tfidf_by_paper.png` | TF-IDF 특징어 |
| 그래프 | `figures/PS_bigram_network.png` | 단어 네트워크 |
| 그래프 | `figures/PS_yearly_trend.png` | 연도별 트렌드 |
