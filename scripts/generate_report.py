#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
APA 7 Word Report Generator — with Tables & Figures
Psychological Safety Literature Review: Tidy Text Analysis
2026-03-12
"""

from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL

BASE = "/Users/shinheepark/Library/CloudStorage/Dropbox/Claude Code/PS Text Analysis"
OUTPUT_PATH = f"{BASE}/2026-03-12_PS-TidyText-Analysis-Report.docx"
FIG_DIR     = f"{BASE}/figures"

# ── 헬퍼: 텍스트 ──────────────────────────────────────────────────────────────

def set_font(run, size=12, bold=False, italic=False):
    run.font.name = "Times New Roman"
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic

def ds(para):
    """double-spacing 적용"""
    para.paragraph_format.line_spacing = Pt(24)
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after  = Pt(0)
    return para

def add_body(doc, text, first_line=0.5, left=0, italic=False):
    p = doc.add_paragraph()
    ds(p)
    p.paragraph_format.first_line_indent = Inches(first_line)
    p.paragraph_format.left_indent = Inches(left)
    r = p.add_run(text)
    set_font(r, italic=italic)
    return p

def add_blank(doc):
    p = doc.add_paragraph()
    ds(p)
    p.paragraph_format.first_line_indent = Inches(0)

def add_heading(doc, text, level=1):
    p = doc.add_paragraph()
    ds(p)
    p.paragraph_format.first_line_indent = Inches(0)
    if level == 1:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(text); set_font(r, bold=True)
    elif level == 2:
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r = p.add_run(text); set_font(r, bold=True)
    elif level == 3:
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r = p.add_run(text); set_font(r, bold=True, italic=True)
    elif level == 4:
        p.paragraph_format.first_line_indent = Inches(0.5)
        r = p.add_run(text + "  "); set_font(r, bold=True)
    return p

def add_ref(doc, text):
    p = doc.add_paragraph()
    ds(p)
    p.paragraph_format.first_line_indent = Inches(-0.5)
    p.paragraph_format.left_indent = Inches(0.5)
    r = p.add_run(text); set_font(r)
    return p

def page_break(doc):
    doc.add_page_break()

# ── 헬퍼: 테이블 ─────────────────────────────────────────────────────────────

def make_apa_table(doc, title_num, title_text, note_text=None):
    """APA 7 테이블 제목 추가 (테이블 위에)"""
    p = doc.add_paragraph()
    ds(p); p.paragraph_format.first_line_indent = Inches(0)
    r = p.add_run(f"Table {title_num}")
    set_font(r, bold=True)

    p2 = doc.add_paragraph()
    ds(p2); p2.paragraph_format.first_line_indent = Inches(0)
    r2 = p2.add_run(title_text)
    set_font(r2, italic=True)

def add_table_note(doc, note_text):
    p = doc.add_paragraph()
    ds(p); p.paragraph_format.first_line_indent = Inches(0)
    r1 = p.add_run("Note. "); set_font(r1, italic=True)
    r2 = p.add_run(note_text); set_font(r2)

def style_header_row(row, bg_color="FFFFFF"):
    """헤더 행: Bold, 하단 테두리"""
    for cell in row.cells:
        for para in cell.paragraphs:
            for run in para.runs:
                run.font.bold = True
                run.font.name = "Times New Roman"
                run.font.size = Pt(12)
        # 하단 테두리
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        tcBorders = OxmlElement('w:tcBorders')
        bottom = OxmlElement('w:bottom')
        bottom.set(qn('w:val'), 'single')
        bottom.set(qn('w:sz'), '12')
        bottom.set(qn('w:space'), '0')
        bottom.set(qn('w:color'), '000000')
        tcBorders.append(bottom)
        tcPr.append(tcBorders)

def style_cell(cell, bold=False, italic=False, align=WD_ALIGN_PARAGRAPH.LEFT):
    for para in cell.paragraphs:
        para.alignment = align
        para.paragraph_format.space_before = Pt(2)
        para.paragraph_format.space_after  = Pt(2)
        for run in para.runs:
            run.font.name = "Times New Roman"
            run.font.size = Pt(11)
            run.font.bold   = bold
            run.font.italic = italic

def set_col_width(table, col_idx, width_inches):
    for row in table.rows:
        row.cells[col_idx].width = Inches(width_inches)

def remove_table_borders(table):
    """수직 테두리 제거, 상단/하단만 유지 (APA 스타일)"""
    tbl = table._tbl
    tblPr = tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ['left', 'right', 'insideV', 'insideH']:
        b = OxmlElement(f'w:{border_name}')
        b.set(qn('w:val'), 'none')
        tblBorders.append(b)
    for border_name in ['top', 'bottom']:
        b = OxmlElement(f'w:{border_name}')
        b.set(qn('w:val'), 'single')
        b.set(qn('w:sz'), '12')
        b.set(qn('w:color'), '000000')
        tblBorders.append(b)
    tblPr.append(tblBorders)

# ── 헬퍼: 피겨 ───────────────────────────────────────────────────────────────

def add_figure(doc, fig_num, fig_title, img_path, note_text=None, width=5.5):
    """APA 7 Figure: 제목 위, 이미지, 노트 아래"""
    p_title = doc.add_paragraph()
    ds(p_title); p_title.paragraph_format.first_line_indent = Inches(0)
    p_title.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r1 = p_title.add_run(f"Figure {fig_num}"); set_font(r1, bold=True)

    p_cap = doc.add_paragraph()
    ds(p_cap); p_cap.paragraph_format.first_line_indent = Inches(0)
    r2 = p_cap.add_run(fig_title); set_font(r2, italic=True)

    p_img = doc.add_paragraph()
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img.paragraph_format.space_before = Pt(4)
    p_img.paragraph_format.space_after  = Pt(4)
    try:
        run = p_img.add_run()
        run.add_picture(img_path, width=Inches(width))
    except Exception as e:
        r = p_img.add_run(f"[Figure: {img_path}]"); set_font(r)

    if note_text:
        p_note = doc.add_paragraph()
        ds(p_note); p_note.paragraph_format.first_line_indent = Inches(0)
        r_n1 = p_note.add_run("Note. "); set_font(r_n1, italic=True)
        r_n2 = p_note.add_run(note_text); set_font(r_n2)

# ═══════════════════════════════════════════════════════════════════════════════
# 데이터
# ═══════════════════════════════════════════════════════════════════════════════

TABLE1_PAPERS = [
    ("Baer & Frese", "2003", "J. of Organizational Behavior", "PS 풍토와 과정 혁신, 기업 성과"),
    ("Chiumento", "2024", "Advances in Developing Human Resources", "HRD와 개인 수준 PS 구축"),
    ("Choi", "2004", "Creativity Research Journal", "창의적 수행의 개인·맥락 예측 요인과 PS"),
    ("Cuellar et al.", "2018", "Annals of Family Medicine", "학습 문화, PS, 번아웃"),
    ("Edmondson & Bransby", "2023", "Annual Review of Org. Psych. & OB", "PS 연구 성숙기: 주요 주제 고찰"),
    ("Edmondson & Lei", "2014", "Annual Review of Org. Psych. & OB", "PS의 역사·르네상스·미래"),
    ("Edmondson", "1999", "Administrative Science Quarterly", "팀 PS와 학습 행동 (기념비적 연구)"),
    ("Edmondson", "2002", "Int'l Handbook of Org. Teamwork", "학습 위험 관리: 팀 내 PS"),
    ("Edmondson", "2003", "Journal of Management Studies", "수술실 발언: 팀 리더와 학습"),
    ("Frazier et al.", "2017", "Personnel Psychology", "PS 메타분석 (k=136, N>21,000)"),
    ("Gerpott et al.", "2019", "Int'l J. of Human Resource Management", "연령 다양성, 지식 공유, PS"),
    ("Han et al.", "2019", "Performance Improvement Quarterly", "팀 창의성: PS와 공유 리더십"),
    ("Higgins et al.", "2012", "Journal of Educational Change", "학교 조직학습과 PS"),
    ("Hunt et al.", "2021", "Int'l J. of Mental Health Systems", "정신건강 서비스에서 PS 향상"),
    ("Huyghebaert et al.", "2018", "Advances in Developing Human Resources", "PS 풍토: HRD 목표와 직원 기능"),
    ("Nembhard & Edmondson", "2006", "Journal of Organizational Behavior", "리더 포용성과 전문 지위, PS"),
    ("Newman et al.", "2017a", "Human Resource Management Review", "PS 체계적 문헌 고찰"),
    ("Newman et al.", "2017b", "Human Resource Management Review", "PS 체계적 문헌 고찰 (동일 문헌)"),
    ("Wanless", "2016", "Research in Human Development", "인간 발달에서 PS의 역할"),
    ("Zehr", "2017", "미확인", "PS와 직원 몰입의 관계"),
]

TABLE2_BIGRAMS = [
    (1,  "psychological safety",    3080, "심리적 안전감 (핵심 개념)"),
    (2,  "employee engagement",      631, "직원 몰입"),
    (3,  "team learning",            200, "팀 학습"),
    (4,  "organizational learning",  192, "조직 학습"),
    (5,  "psychologically safe",     189, "심리적으로 안전한"),
    (6,  "team psychological",       187, "팀 심리적"),
    (7,  "learning behavior",        154, "학습 행동"),
    (8,  "human resource",           143, "인적 자원"),
    (9,  "organizational behavior",  130, "조직 행동"),
    (10, "team performance",         125, "팀 성과"),
    (11, "health care",              122, "의료 (보건의료 맥락)"),
    (12, "shared leadership",        115, "공유 리더십"),
    (13, "knowledge sharing",        112, "지식 공유"),
    (14, "quality improvement",      104, "질 향상 (의료 맥락)"),
    (15, "applied psychology",        98, "응용 심리학"),
    (16, "cognitive deficits",        84, "인지적 결함 (HRD 맥락)"),
    (17, "age diversity",             83, "연령 다양성"),
    (18, "resource management",       82, "자원 관리"),
    (19, "team creativity",           80, "팀 창의성"),
    (20, "leader inclusiveness",      79, "리더 포용성"),
]

TABLE3_TFIDF = [
    ("Zehr (2017)",              "employee engagement",       0.0892, "PS와 직원 몰입의 연결"),
    ("Higgins et al. (2012)",    "reinforces learning",       0.0733, "학습 강화 리더십"),
    ("Han et al. (2019)",        "shared leadership",         0.0721, "공유 리더십"),
    ("Gerpott et al. (2019)",    "age diversity",             0.0714, "연령 다양성"),
    ("Chiumento (2024)",         "cognitive deficits",        0.0513, "인지적 결함 (HRD)"),
    ("Huyghebaert et al. (2018)","psychological health",      0.0479, "심리적 건강"),
    ("Wanless (2016)",           "human development",         0.0449, "인간 발달"),
    ("Cuellar et al. (2018)",    "independent practices",     0.0423, "독립 의원 (의료 맥락)"),
    ("Choi (2004)",              "psychological processes",   0.0388, "심리적 과정 매개"),
    ("Edmondson (1999)",         "team efficacy",             0.0388, "팀 효능감"),
    ("Edmondson (2002)",         "collective learning",       0.0215, "집단적 학습"),
    ("Nembhard & Edmondson (2006)", "cross disciplinary",     0.0214, "초학문적 협업"),
    ("Hunt et al. (2021)",       "mental health",             0.0173, "정신건강 서비스"),
    ("Edmondson & Bransby (2023)", "established literature",  0.0165, "성숙한 문헌"),
    ("Baer & Frese (2003)",      "process innovativeness",   0.0201, "과정 혁신성"),
]

# ═══════════════════════════════════════════════════════════════════════════════
# 문서 생성
# ═══════════════════════════════════════════════════════════════════════════════

doc = Document()

for section in doc.sections:
    section.top_margin    = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin   = Inches(1)
    section.right_margin  = Inches(1)
    section.page_height   = Inches(11)
    section.page_width    = Inches(8.5)

doc.styles['Normal'].font.name = 'Times New Roman'
doc.styles['Normal'].font.size = Pt(12)

# ── TITLE PAGE ────────────────────────────────────────────────────────────────
for _ in range(6): add_blank(doc)

p = doc.add_paragraph()
ds(p); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.first_line_indent = Inches(0)
r = p.add_run("심리적 안전감 문헌의 텍스트 분석:\nTidy Text 방법론을 활용한 Bigram 빈도 및 의미 구조 탐색")
set_font(r, bold=True)

for _ in range(2): add_blank(doc)

for line, bold in [("Shinhee Park", False), ("The University of Southern Mississippi", False)]:
    p = doc.add_paragraph(); ds(p)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Inches(0)
    r = p.add_run(line); set_font(r, bold=bold)

for _ in range(2): add_blank(doc)

p = doc.add_paragraph(); ds(p)
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.first_line_indent = Inches(0)
r = p.add_run("Author Note"); set_font(r, bold=True)

p = doc.add_paragraph(); ds(p)
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.first_line_indent = Inches(0)
r = p.add_run(
    "본 보고서는 심리적 안전감 관련 학술 논문 20편에 대한 계량서지학적 텍스트 분석 결과를 담고 있습니다.\n"
    "분석은 R 4.5.2 환경에서 tidytext 패키지를 활용하여 수행되었습니다.\n"
    "Correspondence: shinheepark@usm.edu  |  Date: March 12, 2026"
)
set_font(r)
page_break(doc)

# ── ABSTRACT ──────────────────────────────────────────────────────────────────
add_heading(doc, "Abstract", level=1)

add_body(doc,
    "본 연구는 심리적 안전감(psychological safety)을 주제로 한 학술 논문 20편(1999–2024)을 대상으로 "
    "R 기반 Tidy Text 분석 방법론을 적용하여 문헌 내 핵심 개념 구조와 연구 동향을 탐색하였다. "
    "PDF 문서에서 텍스트를 추출하고 캐싱 파이프라인을 구축한 후, Bigram 빈도 분석, TF-IDF 분석, "
    "단어 네트워크 시각화, 연도별 개념 트렌드 분석을 수행하였다. "
    "분석 결과, 전체 corpus에서 'psychological safety'(3,080회), 'employee engagement'(631회), "
    "'team learning'(200회), 'organizational learning'(192회), 'learning behavior'(154회) 순으로 "
    "높은 빈도를 나타냈다. TF-IDF 분석을 통해 논문별 특징적 개념어를 식별하였으며, "
    "네트워크 시각화는 심리적 안전감이 팀 학습, 리더십, 창의성, 조직 성과와 복잡하게 연결된 "
    "개념 생태계를 형성함을 보여주었다. 본 연구는 계량서지학적 텍스트 분석 방법론의 효용성을 실증하며, "
    "심리적 안전감 연구의 개념적 지형 파악에 기여한다.", first_line=0)

add_blank(doc)
p = doc.add_paragraph(); ds(p); p.paragraph_format.first_line_indent = Inches(0)
r1 = p.add_run("Keywords: "); set_font(r1, italic=True)
r2 = p.add_run("psychological safety, tidy text analysis, bigram, TF-IDF, network visualization, systematic literature review, organizational learning")
set_font(r2)
page_break(doc)

# ── BODY TITLE ────────────────────────────────────────────────────────────────
p = doc.add_paragraph(); ds(p)
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.first_line_indent = Inches(0)
r = p.add_run("심리적 안전감 문헌의 텍스트 분석:\nTidy Text 방법론을 활용한 Bigram 빈도 및 의미 구조 탐색")
set_font(r, bold=True)
add_blank(doc)

# ── 서론 ──────────────────────────────────────────────────────────────────────
add_heading(doc, "서론", level=1)
add_body(doc,
    "심리적 안전감(psychological safety)은 조직 및 팀 연구에서 가장 활발하게 탐구되는 개념 중 하나로 자리매김하였다. "
    "Kahn(1990)이 처음으로 개인의 직무 몰입을 설명하는 심리적 조건 중 하나로 안전감을 이론화한 이후, "
    "Edmondson(1999)의 팀 심리적 안전감 연구를 기점으로 본격적인 실증 연구가 축적되어 왔다. "
    "이 개념은 팀원들이 대인 관계적 위험(interpersonal risk taking)을 감수하고 발언하며 실수로부터 "
    "학습할 수 있는 공유된 신념을 의미하며(Edmondson, 1999), 팀 학습, 조직 학습, 혁신, 창의성, "
    "직원 몰입 등 다양한 조직 성과와 연결되어 있다.")
add_body(doc,
    "지난 25년간 심리적 안전감 관련 연구는 양적으로 급격히 팽창하였으며, 이는 체계적인 문헌 검토의 필요성을 "
    "높이고 있다. 본 연구는 심리적 안전감 관련 학술 논문 20편을 대상으로 R Tidy Text 분석 방법론을 적용하여 "
    "다음의 연구 목적을 추구한다: (1) 문헌 전체에서 가장 빈번하게 등장하는 개념 쌍(bigram)을 식별하고, "
    "(2) 개별 논문의 특징적 개념어를 TF-IDF를 통해 추출하며, (3) 개념 간 연결 구조를 네트워크로 "
    "시각화하고, (4) 시간 흐름에 따른 핵심 개념의 변화 추이를 탐색한다.")
add_blank(doc)

# ══ TABLE 1 (서론 직후) ═══════════════════════════════════════════════════════
make_apa_table(doc, 1, "분석 대상 논문 목록 (N = 20)")

t1 = doc.add_table(rows=1, cols=4)
t1.alignment = WD_TABLE_ALIGNMENT.CENTER
remove_table_borders(t1)
hdr = t1.rows[0]
for cell, txt in zip(hdr.cells, ["저자", "연도", "게재 저널", "주요 연구 초점"]):
    cell.text = txt
style_header_row(hdr)
set_col_width(t1, 0, 1.2); set_col_width(t1, 1, 0.5)
set_col_width(t1, 2, 2.5); set_col_width(t1, 3, 2.3)

for auth, yr, jour, focus in TABLE1_PAPERS:
    row = t1.add_row()
    for cell, txt in zip(row.cells, [auth, yr, jour, focus]):
        cell.text = txt
        style_cell(cell)

add_table_note(doc, "동일 저자의 중복 문헌(Newman et al., 2017a/b)은 PDF 파일이 별도로 존재하여 corpus에 포함됨.")
add_blank(doc)

# ── 이론적 배경 ───────────────────────────────────────────────────────────────
add_heading(doc, "이론적 배경", level=1)
add_heading(doc, "심리적 안전감의 개념적 기원과 발전", level=2)
add_body(doc,
    "심리적 안전감의 개념적 뿌리는 1960년대 조직 변화 연구로 거슬러 올라간다. Schein과 Bennis(1965)는 "
    "조직 변화 과정에서 구성원들이 불안감 없이 새로운 행동을 시도할 수 있는 안전한 환경의 필요성을 처음 강조하였다. "
    "이후 Kahn(1990)은 개인의 직무 몰입을 결정하는 세 가지 심리적 조건—의미성(meaningfulness), "
    "안전감(safety), 가용성(availability)—을 이론화하였다. Edmondson(1999)은 이를 팀 수준으로 확장하여 "
    "팀 심리적 안전감을 \"팀이 대인 관계적 위험 감수에 안전하다는 팀원들의 공유된 신념\"으로 정의하였다.")
add_body(doc,
    "Edmondson과 Lei(2014)는 심리적 안전감 연구의 역사, 르네상스, 미래를 조망하는 리뷰 논문에서 "
    "이 개념이 개인, 팀, 조직 수준에서 다양하게 적용되어 왔음을 정리하였다. "
    "Edmondson과 Bransby(2023)는 이 분야의 연구가 성숙기에 접어들었음을 선언하며, "
    "새로운 측정 방식, 다층 분석, 그리고 심리적 안전감의 어두운 면에 대한 연구 확장을 촉구하였다.")
add_blank(doc)

add_heading(doc, "선행 요인 및 결과 요인", level=2)
add_body(doc,
    "기존 문헌은 심리적 안전감의 선행 요인으로 리더십 행동, 조직 문화, 집단 규범 등을 일관되게 제시한다. "
    "Nembhard와 Edmondson(2006)은 리더의 포용성(leader inclusiveness)이 심리적 안전감을 촉진하고 "
    "전문적 지위에 따른 발언 억제를 완화함을 발견하였다. Detert와 Burris(2007)는 관리자의 개방성이 "
    "부하직원의 발언 행동을 예측함을 보였다. 결과 요인으로는 학습 행동, 혁신, 창의성, 팀 성과가 주목받아 왔다. "
    "Frazier 등(2017)의 메타분석(k = 136, N > 21,000)은 심리적 안전감이 정보 공유, 학습 행동, 수행과 "
    "유의한 정적 관계임을 종합하였다.")
add_blank(doc)

add_heading(doc, "HRD 및 조직 학습 맥락", level=2)
add_body(doc,
    "조직 학습(organizational learning) 연구에서 심리적 안전감은 학습의 핵심 촉진 요인으로 위치한다. "
    "Edmondson(1999, 2003)은 팀의 학습 행동이 심리적 안전감에 의해 촉진됨을 실증하였다. "
    "HRD 분야에서 Huyghebaert 등(2018)은 심리적 안전 풍토가 자기결정이론의 기본 심리적 욕구를 만족시킴으로써 "
    "직원의 기능에 긍정적 영향을 미친다고 주장하였다. 교육 맥락에서 Higgins 등(2012)은 학교 조직의 학습을 "
    "심리적 안전감, 실험 정신, 학습을 강화하는 리더십의 역할로 분석하였다.")
add_blank(doc)

add_heading(doc, "텍스트 마이닝과 Tidy Text 방법론", level=2)
add_body(doc,
    "Silge와 Robinson(2017)이 개발한 Tidy Text 프레임워크는 R의 tidy data 원칙을 텍스트 분석에 적용한다. "
    "Bigram 분석은 인접한 두 단어 조합을 분석 단위로 하여 풍부한 문맥 정보를 제공하며, "
    "TF-IDF는 문서 내 단어의 상대적 중요도를 측정하여 각 문서의 특징적 어휘를 식별하는 데 효과적이다.")
add_blank(doc)

# ── 연구 방법 ─────────────────────────────────────────────────────────────────
add_heading(doc, "연구 방법", level=1)
add_heading(doc, "분석 대상 및 텍스트 추출", level=2)
add_body(doc,
    "본 연구는 심리적 안전감 관련 논문 20편(1999–2024)을 분석 대상으로 하였다. "
    "논문 선정 기준은 (1) 피어 리뷰 학술지 게재 논문 또는 학술 저서의 챕터, "
    "(2) 심리적 안전감을 주요 변인으로 다루거나 이론적으로 논의한 문헌, (3) PDF 전문 접근 가능 여부였다. "
    "분석 대상 논문의 상세 목록은 Table 1에 제시하였다.")
add_body(doc,
    "텍스트 추출에는 R pdftools 패키지를 사용하였으며, 각 PDF를 최초 1회만 파싱하여 "
    "평문 텍스트(.txt)로 캐싱하는 파이프라인을 구축하였다. "
    "전처리 단계에서는 소문자 변환, 숫자·특수문자 제거, 공백 정규화를 실시하였다. "
    "불용어 제거에는 SMART 불용어 목록에 출판사명·URL 구성 요소 등의 맞춤형 단어를 추가하여 사용하였다.")
add_blank(doc)

add_heading(doc, "분석 방법", level=2)
add_heading(doc, "Bigram 빈도 분석.", level=4)
add_body(doc,
    "tidytext::unnest_tokens(token = 'ngrams', n = 2)로 bigram을 추출하고, "
    "불용어 포함 bigram 및 2자 미만 단어를 필터링한 후 빈도를 집계하였다.", first_line=0)
add_heading(doc, "TF-IDF 분석.", level=4)
add_body(doc,
    "tidytext::bind_tf_idf()로 문서-bigram 조합별 TF-IDF 점수를 산출하였다. "
    "TF-IDF가 높은 bigram은 해당 논문에서 특별히 강조되는 개념이다.", first_line=0)
add_heading(doc, "네트워크 시각화.", level=4)
add_body(doc,
    "상위 60개 bigram으로 igraph 그래프를 구성하고, ggraph의 Fruchterman-Reingold "
    "레이아웃으로 시각화하였다. 엣지 굵기는 bigram 빈도에 비례한다.", first_line=0)
add_heading(doc, "연도별 트렌드 분석.", level=4)
add_body(doc,
    "핵심 bigram 10개를 선정하여 출판 연도별 출현 빈도를 추적하였다. "
    "tidyr::complete()로 결측 연도-bigram 조합을 0으로 보완하였다.", first_line=0)
add_blank(doc)

# ── 결과 ──────────────────────────────────────────────────────────────────────
add_heading(doc, "결과", level=1)

add_heading(doc, "Corpus 기본 통계", level=2)
add_body(doc,
    "분석 corpus는 총 20편, 출판 연도 1999–2024, 불용어 제거 후 유효 bigram 약 199,000 토큰, "
    "고유 bigram 유형 약 45,000개이다. 논문별 규모는 Zehr(2017)의 약 7,300개에서 "
    "Chiumento(2024)의 약 49,000개까지 분포하였다.")
add_blank(doc)

add_heading(doc, "전체 Bigram 빈도 분석", level=2)
add_body(doc,
    "20편 전체 corpus에서 가장 빈번한 bigram은 'psychological safety'(3,080회)로, "
    "2위 'employee engagement'(631회)의 약 5배에 달한다. "
    "상위 20개 bigram은 Table 2에 제시하였으며, Figure 1은 시각화 결과이다. "
    "'team learning'(200회), 'organizational learning'(192회), 'learning behavior'(154회)가 "
    "공히 상위권에 위치하는 것은 심리적 안전감 연구가 조직 학습 패러다임과 밀접히 연결됨을 반영한다.")
add_blank(doc)

# ══ TABLE 2 (bigram 빈도 결과 절) ════════════════════════════════════════════
make_apa_table(doc, 2, "전체 Corpus 상위 20개 Bigram 빈도")

t2 = doc.add_table(rows=1, cols=4)
t2.alignment = WD_TABLE_ALIGNMENT.CENTER
remove_table_borders(t2)
hdr2 = t2.rows[0]
for cell, txt in zip(hdr2.cells, ["순위", "Bigram", "빈도(n)", "의미"]):
    cell.text = txt
style_header_row(hdr2)
set_col_width(t2, 0, 0.5); set_col_width(t2, 1, 1.9)
set_col_width(t2, 2, 0.8); set_col_width(t2, 3, 3.3)

for rank, bigram, n, meaning in TABLE2_BIGRAMS:
    row = t2.add_row()
    for cell, txt in zip(row.cells, [str(rank), bigram, str(n), meaning]):
        cell.text = txt
        style_cell(cell)
    t2.rows[-1].cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    t2.rows[-1].cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

add_table_note(doc, "불용어(SMART 목록 + 출판사·URL 맞춤 불용어) 제거 후 집계. 빈도는 20편 논문 전체 합산값.")
add_blank(doc)

# ══ FIGURE 1 (bigram 막대그래프) ══════════════════════════════════════════════
add_figure(doc, 1,
    "전체 Corpus 상위 20개 Bigram 빈도 (막대그래프)",
    f"{FIG_DIR}/PS_bigram_top20.png",
    note_text="20편 논문 전체 합산. 불용어 제거 후 상위 20개 bigram. "
              "색상 밝기는 빈도에 비례(plasma 팔레트).",
    width=5.5)
add_blank(doc)

add_heading(doc, "TF-IDF 분석: 논문별 특징 개념어", level=2)
add_body(doc,
    "TF-IDF 분석은 각 논문이 강조하는 독특한 개념들을 효과적으로 식별하였다. "
    "논문별 최고 TF-IDF bigram과 그 의미는 Table 3에 제시하였으며, "
    "상위 8개 논문의 TF-IDF 결과는 Figure 2에 시각화하였다. "
    "Zehr(2017)는 'employee engagement'(tf-idf = 0.089)를 가장 특징적인 bigram으로 보이며, "
    "Han 등(2019)은 'shared leadership'(0.072)과 'team creativity'(0.070)가 "
    "특징어로 나타나 공유 리더십과 팀 창의성의 관계 탐구가 핵심임을 드러낸다.")
add_blank(doc)

# ══ TABLE 3 (TF-IDF) ══════════════════════════════════════════════════════════
make_apa_table(doc, 3, "논문별 최고 TF-IDF Bigram (논문당 1위)")

t3 = doc.add_table(rows=1, cols=4)
t3.alignment = WD_TABLE_ALIGNMENT.CENTER
remove_table_borders(t3)
hdr3 = t3.rows[0]
for cell, txt in zip(hdr3.cells, ["논문", "특징 Bigram", "TF-IDF", "이론적 초점"]):
    cell.text = txt
style_header_row(hdr3)
set_col_width(t3, 0, 2.0); set_col_width(t3, 1, 1.8)
set_col_width(t3, 2, 0.7); set_col_width(t3, 3, 2.0)

for paper, bigram, score, focus in TABLE3_TFIDF:
    row = t3.add_row()
    for cell, txt in zip(row.cells, [paper, bigram, f"{score:.3f}", focus]):
        cell.text = txt
        style_cell(cell)
    t3.rows[-1].cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

add_table_note(doc, "TF-IDF = Term Frequency–Inverse Document Frequency. "
               "값이 클수록 해당 논문에서 독특하게 강조되는 bigram임. "
               "참고문헌 노이즈 bigram(출판사명 등)은 제외.")
add_blank(doc)

# ══ FIGURE 2 (TF-IDF per paper) ═══════════════════════════════════════════════
add_figure(doc, 2,
    "논문별 TF-IDF 상위 5개 Bigram (상위 8개 논문)",
    f"{FIG_DIR}/PS_tfidf_by_paper.png",
    note_text="각 패널은 한 논문을 나타냄. 각 논문 내에서 TF-IDF 점수가 높은 "
              "상위 5개 bigram을 표시. 노이즈 bigram은 사전 필터링하였으나 일부 잔류 가능.",
    width=6.0)
add_blank(doc)

# ══ FIGURE 3 (bigram by paper) ════════════════════════════════════════════════
add_figure(doc, 3,
    "논문별 상위 5개 Bigram 빈도 비교 (상위 8개 논문)",
    f"{FIG_DIR}/PS_bigram_by_paper.png",
    note_text="각 패널은 한 논문을 나타냄. 빈도(n)는 해당 논문 내 출현 횟수.",
    width=6.0)
add_blank(doc)

add_heading(doc, "Bigram 네트워크 구조", level=2)
add_body(doc,
    "상위 60개 bigram으로 구성된 네트워크 시각화(Figure 4)는 심리적 안전감 개념의 복잡한 연결 구조를 드러낸다. "
    "'psychological', 'safety', 'team', 'learning'이 네트워크의 허브(hub) 노드로 다수의 엣지와 연결된다. "
    "이들은 'psychological safety', 'team learning', 'psychological meaningfulness' 등의 "
    "핵심 클러스터를 형성하며, 주변부에 'health care', 'leader inclusiveness', 'knowledge sharing' 등 "
    "맥락 특수적 개념들이 위치한다.")
add_blank(doc)

# ══ FIGURE 4 (network) ════════════════════════════════════════════════════════
add_figure(doc, 4,
    "심리적 안전감 Bigram 네트워크 시각화 (상위 60개 Bigram)",
    f"{FIG_DIR}/PS_bigram_network.png",
    note_text="노드 = 단어, 엣지 = bigram 공출현 관계. "
              "엣지 굵기와 투명도는 bigram 빈도에 비례. "
              "Fruchterman-Reingold 알고리즘으로 레이아웃 생성(seed = 42).",
    width=6.0)
add_blank(doc)

add_heading(doc, "연도별 개념 트렌드", level=2)
add_body(doc,
    "Figure 5는 핵심 bigram 10개의 출판 연도별 출현 빈도 변화를 보여준다. "
    "'psychological safety'는 전 시기에 걸쳐 안정적으로 등장하며, "
    "'team learning'과 'organizational learning'은 2000년대 초반부터 지속적으로 관찰된다. "
    "2010년대 이후에는 'creative performance', 'shared leadership' 등 다양화된 연구 맥락을 반영하는 "
    "bigram들이 등장하며, 심리적 안전감 연구가 보건의료, 교육, HRD 등 다양한 분야로 확장됨을 보여준다.")
add_blank(doc)

# ══ FIGURE 5 (yearly trend) ═══════════════════════════════════════════════════
add_figure(doc, 5,
    "핵심 Bigram의 출판 연도별 빈도 트렌드 (1999–2024)",
    f"{FIG_DIR}/PS_yearly_trend.png",
    note_text="각 점은 해당 연도에 출판된 논문들에서의 bigram 총 출현 빈도. "
              "corpus가 20편으로 제한되어 연속적 트렌드보다 개별 논문 효과가 반영될 수 있음.",
    width=6.0)
add_blank(doc)

# ── 논의 ──────────────────────────────────────────────────────────────────────
add_heading(doc, "논의", level=1)
add_body(doc,
    "본 연구의 Tidy Text 분석 결과는 심리적 안전감 연구의 개념적 지형에 대한 세 가지 중요한 통찰을 제공한다. "
    "첫째, 심리적 안전감은 팀 학습 및 조직 학습과 긴밀히 연결된 핵심 개념으로 확고히 자리매김하였다. "
    "Table 2에서 'team learning', 'organizational learning', 'learning behavior'가 공히 상위권에 "
    "위치하는 것은 Edmondson(1999)의 최초 이론화가 문헌 전반에 지속적 영향을 미치고 있음을 보여준다.")
add_body(doc,
    "둘째, Table 3의 TF-IDF 분석은 20편의 논문이 심리적 안전감이라는 공통 주제 아래 "
    "서로 다른 이론적 렌즈와 맥락적 변인을 활용하고 있음을 드러낸다. "
    "공유 리더십(Han et al., 2019), 연령 다양성(Gerpott et al., 2019), 직원 몰입(Zehr, 2017) 등 "
    "각 논문의 이론적 특수성이 TF-IDF 분석을 통해 효과적으로 식별되었다. "
    "셋째, Figure 4의 네트워크 시각화는 심리적 안전감이 리더십, 학습, 성과, 웰빙, 다양성이라는 "
    "여러 조직 현상과 연결된 다차원적 구성체임을 시각적으로 명확히 보여준다(Edmondson & Bransby, 2023).")
add_blank(doc)

# ── 한계 및 향후 연구 ─────────────────────────────────────────────────────────
add_heading(doc, "연구의 한계 및 향후 연구 방향", level=1)
add_body(doc,
    "본 연구는 몇 가지 한계를 가진다. 첫째, 분석 corpus가 20편으로 한정되어 있어 "
    "향후 더 많은 문헌을 포함한 대규모 corpus 구축이 필요하다. "
    "둘째, PDF 텍스트 추출 과정에서 참고문헌 섹션이 혼재되어 일부 노이즈 bigram이 발생하였다. "
    "셋째, 본 연구는 영어 텍스트만을 대상으로 하였으므로 비영어권 연구 포함 시 결과가 달라질 수 있다. "
    "향후 연구에서는 LDA 토픽 모델링, Word2Vec, 또는 BERT 기반 의미론적 분석으로 확장할 수 있다.")
add_blank(doc)

# ── 결론 ──────────────────────────────────────────────────────────────────────
add_heading(doc, "결론", level=1)
add_body(doc,
    "본 연구는 심리적 안전감 관련 논문 20편을 대상으로 R Tidy Text 방법론을 적용하여 문헌의 개념적 구조와 "
    "연구 동향을 탐색하였다. 분석 결과는 팀 학습과 조직 학습이 심리적 안전감 연구의 핵심 관련 개념임을 확인하며, "
    "TF-IDF 분석과 네트워크 시각화를 통해 이 분야의 이론적 다양성과 개념 생태계의 복잡성을 드러낸다. "
    "방법론적 기여 측면에서, PDF 캐싱 파이프라인과 Tidy Text 분석의 결합은 대규모 문헌 검토를 위한 "
    "효율적이고 재현 가능한 도구임을 실증하였다.")
add_blank(doc)
page_break(doc)

# ── REFERENCES ────────────────────────────────────────────────────────────────
add_heading(doc, "References", level=1)

refs = [
    "Baer, M., & Frese, M. (2003). Innovation is not enough: Climates for initiative and psychological safety, process innovations, and firm performance. Journal of Organizational Behavior, 24(1), 45–68. https://doi.org/10.1002/job.179",
    "Carmeli, A., Reiter-Palmon, R., & Ziv, E. (2010). Inclusive leadership and employee involvement in creative tasks in the workplace: The mediating role of psychological safety. Creativity Research Journal, 22(3), 250–260. https://doi.org/10.1080/10400419.2010.504654",
    "Chiumento, A. (2024). A phenomenological study of the role HRD plays in building psychological safety for individuals. [Full journal citation requires manual verification].",
    "Choi, J. N. (2004). Individual and contextual predictors of creative performance: The mediating role of psychological processes. Creativity Research Journal, 16(2–3), 187–199. https://doi.org/10.1080/10400419.2004.9651452",
    "Cuellar, A., Krist, A. H., Nichols, L. M., & Kuzel, A. J. (2018). Effect of practice ownership on work environment, learning culture, psychological safety, and burnout. The Annals of Family Medicine, 16(Suppl 1), S44–S51. https://doi.org/10.1370/afm.2198",
    "Detert, J. R., & Burris, E. R. (2007). Leadership behavior and employee voice: Is the door really open? Academy of Management Journal, 50(4), 869–884. https://doi.org/10.5465/amj.2007.26279183",
    "Edmondson, A. C. (1999). Psychological safety and learning behavior in work teams. Administrative Science Quarterly, 44(2), 350–383. https://doi.org/10.2307/2666999",
    "Edmondson, A. C. (2002). Managing the risk of learning: Psychological safety in work teams. In M. A. West, D. Tjosvold, & K. G. Smith (Eds.), International handbook of organizational teamwork and cooperative working (pp. 255–275). Blackwell. https://doi.org/10.1002/9780470696712.ch13",
    "Edmondson, A. C. (2003). Speaking up in the operating room: How team leaders promote learning in interdisciplinary action teams. Journal of Management Studies, 40(6), 1419–1452. https://doi.org/10.1111/1467-6486.00386",
    "Edmondson, A. C. (2004). Learning from mistakes is easier said than done: Group and organizational influences on the detection and correction of human error. The Journal of Applied Behavioral Science, 40(1), 66–90. https://doi.org/10.1177/0021886304263849",
    "Edmondson, A. C., & Bransby, D. P. (2023). Psychological safety comes of age: Observed themes in an established literature. Annual Review of Organizational Psychology and Organizational Behavior, 10, 55–78. https://doi.org/10.1146/annurev-orgpsych-120920-055217",
    "Edmondson, A. C., & Lei, Z. (2014). Psychological safety: The history, renaissance, and future of an interpersonal construct. Annual Review of Organizational Psychology and Organizational Behavior, 1, 23–43. https://doi.org/10.1146/annurev-orgpsych-031413-091305",
    "Frazier, M. L., Fainshmidt, S., Klinger, R. L., Pezeshkan, A., & Vracheva, V. (2017). Psychological safety: A meta-analytic review and extension. Personnel Psychology, 70(1), 113–165. https://doi.org/10.1111/peps.12183",
    "Gerpott, F. H., Lehmann-Willenbrock, N., Wenzel, R., & Voelpel, S. C. (2019). Age diversity and learning outcomes in organizational training groups: The role of knowledge sharing and psychological safety. The International Journal of Human Resource Management, 32(18), 3777–3805. https://doi.org/10.1080/09585192.2019.1640763",
    "Han, S. J., Lee, Y., & Beyerlein, M. (2019). Developing team creativity: The influence of psychological safety and relation-oriented shared leadership. Performance Improvement Quarterly, 32(2), 159–182. https://doi.org/10.1002/piq.21293",
    "Higgins, M., Ishimaru, A., Holcombe, R., & Fowler, A. (2012). Examining organizational learning in schools: The role of psychological safety, experimentation, and leadership that reinforces learning. Journal of Educational Change, 13(1), 67–94. https://doi.org/10.1007/s10833-011-9167-9",
    "Hunt, D. F., Bailey, J., Lennox, B. R., & colleagues. (2021). Enhancing psychological safety in mental health services. International Journal of Mental Health Systems, 15, Article 33. https://doi.org/10.1186/s13033-021-00439-1",
    "Huyghebaert, T., Gillet, N., Lahiani, F.-J., Dubois-Fleury, A., & Fouquereau, E. (2018). Psychological safety climate as a human resource development target: Effects on workers functioning through need satisfaction and thwarting. Advances in Developing Human Resources, 20(2), 175–191. https://doi.org/10.1177/1523422318756955",
    "Kahn, W. A. (1990). Psychological conditions of personal engagement and disengagement at work. Academy of Management Journal, 33(4), 692–724. https://doi.org/10.5465/256287",
    "Kark, R., & Carmeli, A. (2009). Alive and creating: The mediating role of vitality and aliveness in the relationship between psychological safety and creative work involvement. Journal of Organizational Behavior, 30(6), 785–804. https://doi.org/10.1002/job.571",
    "Nembhard, I. M., & Edmondson, A. C. (2006). Making it safe: The effects of leader inclusiveness and professional status on psychological safety and improvement efforts in health care teams. Journal of Organizational Behavior, 27(7), 941–966. https://doi.org/10.1002/job.413",
    "Newman, A., Donohue, R., & Eva, N. (2017). Psychological safety: A systematic review of the literature. Human Resource Management Review, 27(3), 521–535. https://doi.org/10.1016/j.hrmr.2017.01.001",
    "R Core Team. (2025). R: A language and environment for statistical computing (Version 4.5.2). R Foundation for Statistical Computing. https://www.R-project.org/",
    "Schein, E. H., & Bennis, W. G. (1965). Personal and organizational change through group methods: The laboratory approach. Wiley.",
    "Silge, J., & Robinson, D. (2017). Text mining with R: A tidy approach. O'Reilly Media. https://www.tidytextmining.com/",
    "Tynan, R. (2005). The effects of threat sensitivity and face giving on dyadic psychological safety and upward communication. Journal of Applied Social Psychology, 35(2), 223–247. https://doi.org/10.1111/j.1559-1816.2005.tb02119.x",
    "Wanless, S. B. (2016). The role of psychological safety in human development. Research in Human Development, 13(1), 6–14. https://doi.org/10.1080/15427609.2016.1141283",
    "West, M. A. (1990). The social psychology of innovation in groups. In M. A. West & J. L. Farr (Eds.), Innovation and creativity at work (pp. 309–333). Wiley.",
    "Wickham, H., Averick, M., Bryan, J., Chang, W., McGowan, L., François, R., Grolemund, G., Hayes, A., Henry, L., Hester, J., Kuhn, M., Pedersen, T. L., Miller, E., Bache, S. M., Müller, K., Ooms, J., Robinson, D., Seidel, D. P., Spinu, V., … Yutani, H. (2019). Welcome to the tidyverse. Journal of Open Source Software, 4(43), 1686. https://doi.org/10.21105/joss.01686",
    "Zehr, S. M. (2017). Safe to be engaged or engaged to be safe: The relationship between psychological safety and employee engagement. [Full journal citation requires manual verification].",
]

for ref in refs:
    add_ref(doc, ref)

doc.save(OUTPUT_PATH)
print(f"✅ 저장 완료: {OUTPUT_PATH}")
print(f"   참고문헌: {len(refs)}개")
print(f"   테이블:   3개 (Table 1–3)")
print(f"   피겨:     5개 (Figure 1–5)")
