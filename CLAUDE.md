# CLAUDE.md — PS Text Analysis 프로젝트 작업 기억 파일

> 이 파일은 Claude Code가 세션마다 읽고 업데이트하는 프로젝트 컨텍스트 파일입니다.
> 새 세션 시작 시 이 파일을 먼저 참조하여 이전 작업을 이어받습니다.

---

## 프로젝트 기본 정보

| 항목 | 내용 |
|------|------|
| 프로젝트명 | Psychological Safety 문헌 Tidy Text 분석 |
| 작업 디렉토리 | `~/Library/CloudStorage/Dropbox/Claude Code/PS Text Analysis/` |
| PDF 원본 경로 | `~/Library/CloudStorage/Dropbox/00 USM/Research/R Studio/Data/PS_PDFfiles/` |
| GitHub | https://github.com/sincerity416/PS-Text-Analysis |
| 분석 대상 | 심리적 안전감(Psychological Safety) 논문 20편 (1999~2024) |
| 분석 언어 | R (tidytext, pdftools, ggraph, widyr) |

---

## 폴더 구조

```
PS Text Analysis/
├── CLAUDE.md              ← 이 파일 (세션 간 작업 기억)
├── .gitignore
├── data/
│   ├── raw_text/          ← PDF → .txt 캐시 (20개, 재파싱 불필요)
│   └── results/           ← 분석 결과 CSV
├── figures/               ← 시각화 PNG
├── scripts/
│   ├── PS_TidyText_Analysis.R   ← 메인 분석 스크립트
│   ├── PS correlation.R         ← 상관분석 스크립트
│   └── PS_PDFfiles_R code.txt   ← 초기 코드 참고용
└── discussion/            ← 세션별 의사결정 기록 md 파일
```

---

## 핵심 설계 원칙

1. **PDF 캐싱:** `data/raw_text/`에 .txt 파일이 있으면 재파싱하지 않음 (성능)
2. **커스텀 불용어:** URL·출판사명(`wiley`, `onlinelibrary` 등) 제거 필수
3. **경로 변수:** 스크립트 내 `BASE_DIR` 하나만 수정하면 전체 경로 자동 변경
4. **Git push:** SSH 키 인증 (`~/.ssh/id_ed25519`), remote = `git@github.com:sincerity416/PS-Text-Analysis.git`

---

## 세션 작업 로그

### 2026-03-12 | PS Tidy Text 분석 설계 및 환경 구축

**완료한 작업:**
- 컨텍스트 초과 방지 전략 5가지 검토 → Option 1+2(단순 tidytext + 캐싱) 채택
- R 분석 스크립트 작성 및 실행 완료
  - Bigram 빈도 분석
  - 논문별 Bigram 비교 (TF-IDF)
  - 네트워크 시각화 (ggraph)
  - 연도별 트렌드 분석
- 분산된 파일들을 단일 작업 디렉토리로 통합
- GitHub SSH 키 생성 및 저장소 연동 완료

**주요 분석 결과 (Top 5 Bigrams):**
1. psychological safety (3,080)
2. employee engagement (631)
3. team learning (200)
4. organizational learning (192)
5. psychologically safe (189)

**해결한 이슈:**
- 1차 실행 시 `wiley online`, `onlinelibrary` 등 출판사 URL이 bigram에 포함됨
  → custom_stopwords 확장으로 해결
- SSH 키 GitHub 등록 시 채팅 복사 → 공백 오류
  → `pbcopy`로 클립보드 직접 복사로 해결

**다음 세션 참고 사항:**
- 분석 재실행 시 `data/raw_text/`에 캐시가 있으므로 즉시 시작 가능
- `key_bigrams` 목록은 결과 보고 필요 시 수정 가능
- 논문 추가 시 PDF를 `PS_PDFfiles/`에 넣으면 자동으로 캐싱됨

### 2026-03-12 | APA 7 보고서 생성 및 표/그림 삽입

**완료한 작업:**
- python-docx로 APA 7 스타일 Word 보고서 생성 (`scripts/generate_report.py`)
- 30개 참고문헌 WebSearch로 hallucination check 완료 (28개 DOI 확인, 2개 수동 검증 권고)
- 보고서 본문 내 표 3개 + 그림 5개 맥락에 맞게 삽입
  - Table 1 (서론 후): 논문 20편 개요 (저자/연도/저널/초점)
  - Table 2 (bigram 결과절): Top 20 bigrams + 한글 의미
  - Figure 1: PS_bigram_top20.png
  - Table 3 (TF-IDF 결과절): 논문별 특징 bigram
  - Figure 2: PS_tfidf_by_paper.png, Figure 3: PS_bigram_by_paper.png
  - Figure 4: PS_bigram_network.png, Figure 5: PS_yearly_trend.png
- GitHub push 완료

**산출물:**
- `2026-03-12_PS-TidyText-Analysis-Report.docx` (표 3개, 그림 5개 포함)
- `scripts/generate_report.py` (APA 7 문서 생성 스크립트)

**해결한 이슈:**
- pandoc 미설치 → python-docx로 전환 (`pip3 install python-docx`)
- TF-IDF 표에 인용 노이즈(`asq june`, `lance frazier`) 포함 → Python에서 노이즈 패턴 필터링
- APA 7 표 세로선 제거: OxmlElement로 left/right/insideV 경계선 'none' 설정

**다음 세션 참고 사항:**
- 보고서 재생성: `python3 scripts/generate_report.py` 실행
- 표/그림 내용 수정 시 `generate_report.py` 내 해당 섹션만 수정

---

<!-- 새 세션 작업 시 아래 형식으로 추가 -->
<!--
### YYYY-MM-DD | 세션 제목

**완료한 작업:**
-

**주요 결정:**
-

**해결한 이슈:**
-

**다음 세션 참고 사항:**
-
-->
