#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gazette_docx_build.py — GAZETTE 독성동향 소식지 md → docx 빌드 스크립트

용법:
    python3 gazette_docx_build.py            # 변환 + 검증(QA 리포트) 모두 수행
    python3 gazette_docx_build.py --verify   # QA 리포트만 출력 (기존 .docx 대상)

스크립트 위치에서 GAZETTE_독성동향_*.md 를 자동으로 찾는다.
출력 .docx 는 md 와 동일 파일명(확장자만 .docx).
디자인 토큰은 GAZETTE_독성동향_DESIGN.md 에 의거(primary #3B8FB7 등).
"""

import argparse
import os
import re
import sys
import zipfile
from pathlib import Path

import docx
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

# ---------------------------------------------------------------------------
# 디자인 토큰 (DESIGN.md)
# ---------------------------------------------------------------------------
PRIMARY = "3B8FB7"        # NEB 중간 청색 — H2·논문 제목·링크·불릿 리드
SECONDARY = "86CCEC"      # 하늘색 — 불릿 글리프·커버 3행·구분선
NAVY = "092B39"           # 남색 — 커버·해석상자·불릿 리드
PANEL_SHADE = "E9F4FA"    # 패널 음영 — 통계 바·해석상자
FONT_COLOR = "64748B"     # 세미 muted — DOI 라벨·푸터
MUTED = "64748B"          # 인용·캡션·보조
HEADER_SHADE = "D7EAF5"   # 표 헤더 음영
BORDER = "E2E8F0"         # 표 경계
BODY_TEXT = "1E293B"      # 본문 텍스트
WHITE = "FFFFFF"

EAST_ASIA_FONT = "맑은 고딕"

SIZE_BODY = 10
SIZE_H2 = 13
SIZE_H3 = 11.5
SIZE_QUOTE = 9.5
SIZE_CAPTION = 9

# 검증 기준치 — (2026-10-01 게재규칙) 소식지는 초록만 진행, 논문 그림은 abstract figure가
# 있을 때만 사용. 현행 소식지에 본문 그림은 넣지 않는다(abstract figure 없음 확인).
EXPECTED_MIN_TABLES = 2
EXPECTED_IMAGES = 0

# 머리말(pageHeader) 이식 — 구판 docx header1.xml 실측 계승
# (NEB 로고 + 'nature·environment·beyond' 문구 + #86CCEC 하단 구분선, 1행 2열 표)
HEADER_TEMPLATE = "word_header1_template.xml"
HEADER_PHRASE = "Nature · Environment · Beyond"
HEADER_LINE_COLOR = "86CCEC"
# 구판 pgMar: top 1247 twips(=2.2cm)·header 397 twips(=0.7cm) — 머리말 표시 공간
HEADER_MARGIN_TOP_TWIPS = "1247"
HEADER_MARGIN_HEADER_TWIPS = "397"

MD_GLOB = "GAZETTE_독성동향_*.md"

# ---------------------------------------------------------------------------
# 저수준 헬퍼
# ---------------------------------------------------------------------------


def _set_rfonts(rpr):
    """run/스타일 rPr 에 맑은 고딕 eastAsia/ascii/hAnsi 설정."""
    rfonts = rpr.get_or_add_rFonts()
    rfonts.set(qn("w:ascii"), EAST_ASIA_FONT)
    rfonts.set(qn("w:hAnsi"), EAST_ASIA_FONT)
    rfonts.set(qn("w:eastAsia"), EAST_ASIA_FONT)


def add_run(paragraph, text, size=SIZE_BODY, bold=False, color=BODY_TEXT,
            italic=False):
    """문단에 포맷된 run 추가 (rFonts eastAsia 포함)."""
    run = paragraph.add_run(text)
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    run.font.color.rgb = RGBColor.from_string(color)
    _set_rfonts(run._element.get_or_add_rPr())
    return run


# 인라인 마크다운: **bold**, [text](url) (url 제거 → text primary 색상)
INLINE_RE = re.compile(r"(\*\*.+?\*\*|\[[^\]]+\]\([^)]*\))")
LINK_RE = re.compile(r"\[([^\]]+)\]\(([^)]*)\)")


def add_inline_runs(paragraph, text, size=SIZE_BODY, bold=False,
                    color=BODY_TEXT, italic=False):
    """**bold** / [text](url) 인라인 형식 파싱 후 runs 추가. 이모지·유니코드 보존."""
    for tok in INLINE_RE.split(text):
        if not tok:
            continue
        if tok.startswith("**") and tok.endswith("**") and len(tok) >= 4:
            add_run(paragraph, tok[2:-2], size=size, bold=True,
                    color=color, italic=italic)
        else:
            m = LINK_RE.match(tok)
            if m:
                # url 없애고 text 를 primary 색상으로
                add_run(paragraph, m.group(1), size=size, bold=bold,
                        color=PRIMARY, italic=italic)
            else:
                add_run(paragraph, tok, size=size, bold=bold, color=color,
                        italic=italic)


def add_italic_runs(paragraph, text, size=SIZE_BODY, color=BODY_TEXT):
    """*이탤릭* 범위 + **bold** + 링크 파서 — 구판 이탤릭 보존 로직."""
    for tok in re.split(r"(\*[^*\n]+\*)", text):
        if not tok:
            continue
        if tok.startswith("*") and tok.endswith("*") and len(tok) >= 2:
            add_inline_runs(paragraph, tok[1:-1], size=size, color=color,
                            italic=True)
        else:
            add_inline_runs(paragraph, tok, size=size, color=color)


def set_style_font(style, size=SIZE_BODY, color=BODY_TEXT):
    """Normal 등 스타일의 기본 폰트·크기·색 설정."""
    style.font.name = EAST_ASIA_FONT
    style.font.size = Pt(size)
    style.font.color.rgb = RGBColor.from_string(color)
    _set_rfonts(style.element.get_or_add_rPr())


def set_section_bottom_line(p, color=SECONDARY, sz="16"):
    """섹션 헤더 하단 밑줄(구판: bottom single sz=16(2pt) space=2)."""
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), sz)
    bottom.set(qn("w:space"), "2")
    bottom.set(qn("w:color"), color)
    pBdr.append(bottom)
    pPr.append(pBdr)


def para_shading(p, fill):
    """문단 배경 음영(구판 통계바·해석상자 E9F4FA)."""
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), fill)
    pPr.append(shd)


def set_para_spacing(p, before, after, line="276"):
    """문단 간격(구판: line 276 lineRule auto + before/after twips)."""
    pPr = p._p.get_or_add_pPr()
    sp = OxmlElement("w:spacing")
    sp.set(qn("w:before"), str(before))
    sp.set(qn("w:after"), str(after))
    sp.set(qn("w:line"), line)
    sp.set(qn("w:lineRule"), "auto")
    pPr.append(sp)


def set_cell_shading(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), fill)
    tcPr.append(shd)


def set_cell_valign_center(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    vAlign = OxmlElement("w:vAlign")
    vAlign.set(qn("w:val"), "center")
    tcPr.append(vAlign)


def set_table_width_100pct(table):
    tblPr = table._tbl.tblPr
    tblW = OxmlElement("w:tblW")
    tblW.set(qn("w:w"), "5000")
    tblW.set(qn("w:type"), "pct")
    tblPr.append(tblW)


def set_table_borders(table, color=BORDER, size="4", val="single"):
    tblPr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement("w:" + edge)
        el.set(qn("w:val"), val)
        if val != "none":
            el.set(qn("w:sz"), size)
            el.set(qn("w:color"), color)
        borders.append(el)
    tblPr.append(borders)


def add_horizontal_rule(doc):
    """md 구분선 `---` → 구판 푸터 분리선 문단(top 경계) + 간격."""
    p = doc.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    top = OxmlElement("w:top")
    top.set(qn("w:val"), "single")
    top.set(qn("w:sz"), "8")  # 구판 footer top boundary = 1pt
    top.set(qn("w:space"), "2")
    top.set(qn("w:color"), BORDER)
    pBdr.append(top)
    pPr.append(pBdr)
    set_para_spacing(p, 120, 40)
    return p


def get_cell_fill(cell):
    tcPr = cell._tc.tcPr
    if tcPr is None:
        return None
    shd = tcPr.find(qn("w:shd"))
    if shd is None:
        return None
    return shd.get(qn("w:fill"))


# ---------------------------------------------------------------------------
# 경로 해석
# ---------------------------------------------------------------------------


def find_md(script_dir):
    """스크립트 위치에서 GAZETTE_독성동향_*.md 자동 검색(최신 mtime 우선)."""
    candidates = sorted(
        script_dir.glob(MD_GLOB),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    # 디자인 규약(DESIGN.md)은 빌드 원문이 아니므로 후보 제외
    candidates = [c for c in candidates if "DESIGN" not in c.name]
    if not candidates:
        raise FileNotFoundError(
            f"'{MD_GLOB}' 을(를) '{script_dir}' 에서 찾지 못했습니다."
        )
    return candidates[0]


def resolve_logo(md_path):
    base = md_path.parent
    candidates = [
        base / "assets" / "NEB_logo.png",
        base / "assets" / "assets" / "NEB_logo.png",
    ]
    for c in candidates:
        if c.exists():
            return str(c)
    return str(candidates[0])


def resolve_image(md_path, rel):
    """이미지 경로를 문서(md) 위치 기준 상대경로로 해석. figures/ 폴백."""
    base = md_path.parent
    p = base / rel
    if p.exists():
        return str(p)
    p2 = base / "figures" / Path(rel).name
    if p2.exists():
        return str(p2)
    return str(p)


def parse_month_title(md_path):
    m = re.search(r"(\d{4})-(\d{2})", md_path.stem)
    if m:
        year, month = m.group(1), int(m.group(2))
        return f"{year}년 {month}월호"
    return "소식지"


def extract_cover_stats(md_text):
    """발표통수(수집편수·저널종수·확보율·발행분포) 추출."""
    n = j = pct = dist = None
    m = re.search(r"수집\s*논문\s*(\d+)편", md_text)
    if m:
        n = m.group(1)
    m = re.search(r"(\d+)종\s*저널", md_text)
    if m:
        j = m.group(1)
    m = re.search(r"(\d+)%", md_text)
    if m:
        pct = m.group(1)
    m = re.search(r"발행[^\n]*\(([^)]*)\)", md_text)
    if m:
        dist = m.group(1).strip()
    parts = []
    if n:
        parts.append(f"수집 논문 {n}편")
    if j:
        parts.append(f"{j}종 저널")
    if pct:
        parts.append(f"초록 확보 {pct}%")
    if dist:
        parts.append(f"발행분포 {dist}")
    return " · ".join(parts)


# ---------------------------------------------------------------------------
# 문서 설정 / 커버
# ---------------------------------------------------------------------------


def setup_document(doc):
    sec = doc.sections[0]
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.top_margin = Cm(2.2)
    sec.bottom_margin = Cm(2.0)
    sec.left_margin = Cm(2.0)
    sec.right_margin = Cm(2.0)
    set_style_font(doc.styles["Normal"], size=SIZE_BODY, color=BODY_TEXT)


GRAD_TOP = "86CCEC"
GRAD_MID = "3B8FB7"
GRAD_BOT = "092B39"
GRAD_STEPS = 24
BAND_H = 2312            # 밴드 행 높이(exact, twips)


def _hex3(h):
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def _grad_hex(t):
    """3-스톱 브랜드 램프(좌→우): 86CCEC → 3B8FB7(중앙) → 092B39."""
    stops = (_hex3(GRAD_TOP), _hex3(GRAD_MID), _hex3(GRAD_BOT))
    seg = min(int(t * 2), 1)
    u = t * 2 - seg
    a, b = stops[seg], stops[seg + 1]
    return "".join("%02X" % round(a[k] + (b[k] - a[k]) * u) for k in range(3))


def _para_shading(p, color):
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), color)
    pPr.append(shd)


def _para_exact(p, twips):
    pPr = p._p.get_or_add_pPr()
    sp = OxmlElement("w:spacing")
    sp.set(qn("w:before"), "0")
    sp.set(qn("w:after"), "0")
    sp.set(qn("w:line"), str(twips))
    sp.set(qn("w:lineRule"), "exact")
    pPr.append(sp)


def _frame_overlay(p, x, y, w):
    """페이지 앵커 플로트 프레임(pPr 선두 framePr) — 밴드 위 텍스트 오버레이."""
    pPr = p._p.get_or_add_pPr()
    fp = OxmlElement("w:framePr")
    fp.set(qn("w:wrap"), "none")
    fp.set(qn("w:vAnchor"), "page")
    fp.set(qn("w:hAnchor"), "page")
    fp.set(qn("w:x"), str(x))
    fp.set(qn("w:y"), str(y))
    fp.set(qn("w:w"), str(w))
    pPr.insert(0, fp)


def build_cover(doc, month_title, stats_text, logo_path, topic_label="독성동향"):
    """커버: navy 그래디언트 — 셀 shd 24단계 램프(좌 #86CCEC→우 #092B39 수평).

    밴드=1행 24열 고정 그리드(nil 경계·행높이 exact 2312 twips), 텍스트 3행은
    framePr 페이지 앵커 플로트 오버레이(원본 잉크 세로 범위 실측 좌표 기준).
    """
    table = doc.add_table(rows=1, cols=GRAD_STEPS)
    set_table_borders(table, val="none")

    tblPr = table._tbl.tblPr
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.append(tblW)
    tblW.set(qn("w:w"), "9638")
    tblW.set(qn("w:type"), "dxa")
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    look = tblPr.find(qn("w:tblLook"))
    if look is not None:
        look.addprevious(layout)
    else:
        tblPr.append(layout)

    prev = 0
    for j in range(GRAD_STEPS):
        cur = round(9638 * (j + 1) / GRAD_STEPS)
        wj = cur - prev
        prev = cur
        table.columns[j].width = wj * 635          # gridCol(EMU)
        cell = table.cell(0, j)
        tcPr = cell._tc.get_or_add_tcPr()
        tcW = tcPr.find(qn("w:tcW"))
        if tcW is None:
            tcW = OxmlElement("w:tcW")
            tcPr.append(tcW)
        tcW.set(qn("w:w"), str(wj))
        tcW.set(qn("w:type"), "dxa")
        set_cell_shading(cell, _grad_hex(j / (GRAD_STEPS - 1)))
    row = table.rows[0]
    trPr = row._tr.get_or_add_trPr()
    trH = OxmlElement("w:trHeight")
    trH.set(qn("w:val"), str(BAND_H))
    trH.set(qn("w:hRule"), "exact")
    trPr.append(trH)

    # 텍스트 3행 — 플로트 프레임(오버레이). y_offset = 원본 잉크 실측(밴드 top 기준).
    sec = doc.sections[0]
    frame_x = sec.left_margin.twips
    band_top = sec.top_margin.twips
    frame_rows = (
        (243, 420, f"{topic_label} 소식지", 16, True, WHITE),
        (834, 780, month_title, 30, True, WHITE),
        (1709, 234, "환경솔루션 사업부 · 월간 저널 모니터링 요약", 9, False, SECONDARY),
    )
    for y_off, lh, txt, sz, bold, clr in frame_rows:
        p = doc.add_paragraph()
        _frame_overlay(p, frame_x, band_top + y_off, 9638)
        _para_exact(p, lh)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_run(p, txt, size=sz, bold=bold, color=clr)

    return table


# ---------------------------------------------------------------------------
# 본문 빌드
# ---------------------------------------------------------------------------


def build_table(doc, lines):
    """`| … |` 연속 행 → docx 표 (tblW 5000 pct, 헤더 굵게+음영)."""
    rows = []
    for ln in lines:
        cells = [c.strip() for c in ln.strip().strip("|").split("|")]
        rows.append(cells)
    header = rows[0]
    # 구분 행(|---|---|) 건너뛰기
    data = [r for r in rows[1:] if not all(re.fullmatch(r"-+", c) for c in r)]

    ncols = len(header)
    table = doc.add_table(rows=1 + len(data), cols=ncols)
    set_table_width_100pct(table)
    set_table_borders(table, color=BORDER, size="4", val="single")

    # 구판 실측: 클러스터 표 3열 = 4800/2600/2238 dxa, 나머지 등분
    W_TOTAL = 9638
    if ncols == 3:
        col_dxa = [4800, 2600, 2238]
    elif ncols == 5:
        col_dxa = [913, 2089, 2342, 1423, 2871]
    else:
        col_dxa = [W_TOTAL // ncols] * ncols
        col_dxa[-1] += W_TOTAL - sum(col_dxa)
    # 표 폭을 구판처럼 dxa 고정 + 중앙 정렬(5000 pct는 남용 시 열폭 붕괴)
    tblPr = table._tbl.tblPr
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is not None:
        tblW.set(qn("w:w"), str(W_TOTAL))
        tblW.set(qn("w:type"), "dxa")
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    look = tblPr.find(qn("w:tblLook"))
    if look is not None:
        look.addprevious(layout)
    else:
        tblPr.append(layout)
    jc = OxmlElement("w:jc")
    jc.set(qn("w:val"), "center")
    tw_el = tblPr.find(qn("w:tblW"))
    if tw_el is not None:
        tw_el.addnext(jc)
    else:
        tblPr.append(jc)
    # gridCol을 배분과 동일하게 교체 — tblLayout fixed에서 실제 열폭 결정 요소
    grid = table._tbl.find(qn("w:tblGrid"))
    for gc, w in zip(grid.findall(qn("w:gridCol")), col_dxa):
        gc.set(qn("w:w"), str(w))
    # 헤더 행 셀 폭 고정
    for j, h in enumerate(header):
        cell = table.cell(0, j)
        tcPr = cell._tc.get_or_add_tcPr()
        tcW = tcPr.find(qn("w:tcW"))
        if tcW is not None:
            tcW.set(qn("w:w"), str(col_dxa[j]))
            tcW.set(qn("w:type"), "dxa")
        set_cell_shading(cell, HEADER_SHADE)
        p = cell.paragraphs[0]
        add_run(p, h, size=SIZE_BODY, bold=True, color=BODY_TEXT)

    # 데이터 행
    for i, r in enumerate(data, start=1):
        for j in range(ncols):
            val = r[j] if j < len(r) else ""
            cell = table.cell(i, j)
            tcPr = cell._tc.get_or_add_tcPr()
            tcW = tcPr.find(qn("w:tcW"))
            if tcW is not None:
                tcW.set(qn("w:w"), str(col_dxa[j]))
                tcW.set(qn("w:type"), "dxa")
            p = cell.paragraphs[0]
            add_inline_runs(p, val, size=SIZE_BODY, color=BODY_TEXT)

    return table


def is_list_item(stripped):
    return bool(re.match(r"^(?:[-•]|\d+\.)\s+", stripped))


def build_body(doc, md_text, md_path, stats):
    lines = md_text.splitlines()
    i = 0
    total = len(lines)

    while i < total:
        raw = lines[i]
        stripped = raw.strip()

        # h1 (# ) → 커버에서 소비, 본문 제거
        if stripped.startswith("# "):
            i += 1
            continue

        # 커버 발표통수 원천 라인(수집 논문 N편) → 커버 소비 후 제거
        if re.match(r"^\*\*수집\s*논문\s+\d+편\*\*", stripped):
            i += 1
            continue

        # h3 (### ) — 불릿 리드·시그널 헤더(구판: 세미 볼드 파랑 11.5pt)
        if stripped.startswith("### "):
            p = doc.add_paragraph()
            set_para_spacing(p, before=200, after=60)
            add_inline_runs(p, stripped[4:].strip(), size=SIZE_H3, bold=True,
                            color=PRIMARY)
            stats["h3"] += 1
            i += 1
            continue

        # h2 (## ) — 섹션 헤더(구판: 파랑 13pt 볼드 + 2pt 하늘색 밑줄)
        if stripped.startswith("## "):
            p = doc.add_paragraph()
            set_section_bottom_line(p)
            set_para_spacing(p, before=280, after=80)
            add_inline_runs(p, stripped[3:].strip(), size=SIZE_H2, bold=True,
                            color=PRIMARY)
            stats["h2"] += 1
            i += 1
            continue

        # 구분선 (---)
        if re.fullmatch(r"-{3,}", stripped):
            add_horizontal_rule(doc)
            stats["hr"] += 1
            i += 1
            continue

        # 표 (| … |)
        if stripped.startswith("|"):
            block = []
            while i < total and lines[i].strip().startswith("|"):
                block.append(lines[i])
                i += 1
            build_table(doc, block)
            stats["tables"] += 1
            continue

        # 이미지 (![alt](path)) — 게재규칙(2026-10-01): 논문 그림은 그 논문에 abstract
        # figure(graphical abstract)가 있을 때만 소스 md에 넣는다. 본문 figure·로고는
        # 소식지에 넣지 않는다(로고는 머리말 전용). 소스 md에서 그림을 이미 필터링함.
        if stripped.startswith("!["):
            m = re.match(r"!\[[^\]]*\]\(([^)]+)\)", stripped)
            if m:
                img_path = resolve_image(md_path, m.group(1))
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = p.add_run()
                run.add_picture(img_path, width=Cm(15.5))
                stats["images"] += 1
            i += 1
            continue

        # 인용 (> ) — 구판 2종: 통계바(패널, 라벨 bold muted 9.5pt)/해석상자(
        # 패널+좌측 하늘색 2.25pt, 남색 10pt)
        if stripped.startswith(">"):
            p = doc.add_paragraph()
            content = stripped.lstrip(">").strip()
            m_lead = re.match(r"^\*\*(.+?)\*\*\s*[:：]?\s*(.*)$", content)
            para_shading(p, PANEL_SHADE)
            set_para_spacing(p, 120, 160)
            is_stat_bar = bool(m_lead and m_lead.group(1).startswith("수집 논문"))
            if is_stat_bar:
                add_run(p, m_lead.group(1), size=9.5, bold=True,
                        color=FONT_COLOR)
                add_inline_runs(p, " " + m_lead.group(2), size=9.5,
                                color=FONT_COLOR)
            else:
                pPr = p._p.get_or_add_pPr()
                pBdr = OxmlElement("w:pBdr")
                left = OxmlElement("w:left")
                left.set(qn("w:val"), "single")
                left.set(qn("w:sz"), "18")
                left.set(qn("w:space"), "2")
                left.set(qn("w:color"), SECONDARY)
                pBdr.append(left)
                pPr.append(pBdr)
                if m_lead:
                    add_run(p, m_lead.group(1) + ": ", size=SIZE_BODY,
                            bold=True, color=NAVY)
                    add_inline_runs(p, m_lead.group(2), size=SIZE_BODY,
                                    color=NAVY)
                else:
                    add_inline_runs(p, content, size=SIZE_BODY, color=NAVY)
            stats["quotes"] += 1
            i += 1
            continue

        # 목록 (불릿/번호) — 구판: '· ' 글리프(하늘색 볼드)·리드 남색 볼드·들여쓰기
        if is_list_item(stripped):
            indent = len(raw) - len(raw.lstrip())
            m_num = re.match(r"^(\d+)\.\s+(.*)$", stripped)
            text_only = re.sub(r"^(?:[-•]|\d+\.\s+)\s*", "", stripped)
            m_lead = re.match(r"^\*\*([^*]+)\*\*[:：]\s*(.*)$", text_only)
            p = doc.add_paragraph()
            if m_lead and m_lead.group(1) == "해석":
                # 해석상자(구판: E9F4FA 음영 + 좌측 하늘색 2.25pt)
                pPr = p._p.get_or_add_pPr()
                pBdr = OxmlElement("w:pBdr")
                left = OxmlElement("w:left")
                left.set(qn("w:val"), "single")
                left.set(qn("w:sz"), "18")
                left.set(qn("w:space"), "2")
                left.set(qn("w:color"), SECONDARY)
                pBdr.append(left)
                pPr.append(pBdr)
                para_shading(p, PANEL_SHADE)
                set_para_spacing(p, 120, 160)
                add_run(p, "해석: ", size=SIZE_BODY, bold=True, color=NAVY)
                add_inline_runs(p, m_lead.group(2), size=SIZE_BODY, color=NAVY)
            elif m_num:
                m_ml = re.match(r"^(\d+\.\s+)\*\*([^*]+)\*\*\s*(.*)$",
                                stripped)
                set_para_spacing(p, 240, 60)
                if m_ml:
                    add_run(p, m_ml.group(1) + m_ml.group(2), size=SIZE_H3,
                            bold=True, color=PRIMARY)
                    add_inline_runs(p, " " + m_ml.group(3), size=SIZE_BODY,
                                    color=BODY_TEXT)
                else:
                    add_run(p, m_num.group(1) + " ", size=SIZE_BODY, bold=True,
                            color=PRIMARY)
                    add_inline_runs(p, m_num.group(2), size=SIZE_BODY,
                                    color=BODY_TEXT)
            else:
                set_para_spacing(p, 0, 100)
                p.paragraph_format.left_indent = (Cm(0.4) if indent >= 2
                                                  else Cm(0.6))
                if m_lead:
                    add_run(p, "·  ", size=SIZE_BODY, bold=True,
                            color=SECONDARY)
                    add_run(p, m_lead.group(1), size=SIZE_BODY, bold=True,
                            color=NAVY)
                    add_inline_runs(p, ": " + m_lead.group(2), size=SIZE_BODY,
                                    color=BODY_TEXT)
                else:
                    add_run(p, "·  ", size=SIZE_BODY, bold=True,
                            color=SECONDARY)
                    add_inline_runs(p, text_only, size=SIZE_BODY,
                                    color=BODY_TEXT)
            stats["bullets"] += 1
            i += 1
            continue

        # 캡션(*Fig. N:* … 이탤릭 라인) — 구판: 좌 ind 227, 9pt muted 이탤릭
        if re.match(r"^\*Fig\.", stripped) and stripped.endswith("*") \
                and not stripped.startswith("**"):
            p = doc.add_paragraph()
            set_para_spacing(p, 40, 80)
            p.paragraph_format.left_indent = Cm(0.4)
            add_inline_runs(p, stripped[1:-1].strip(), size=SIZE_CAPTION,
                            color=MUTED, italic=True)
            stats["captions"] += 1
            i += 1
            continue

        # 기타 전체-이탤릭 라인 — 저널·발행 라인(구판: 중앙 9.5pt muted)
        if stripped.startswith("*") and stripped.endswith("*") \
                and not stripped.startswith("**"):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_para_spacing(p, 40, 80)
            add_inline_runs(p, stripped[1:-1].strip(), size=9.5, bold=True,
                            color=MUTED)
            stats["captions"] += 1
            i += 1
            continue

        # 저자 저작권 라인(— © …) — 구판: 좌 ind, 9pt muted
        if stripped.startswith("—"):
            p = doc.add_paragraph()
            set_para_spacing(p, 40, 160)
            p.paragraph_format.left_indent = Cm(0.4)
            add_inline_runs(p, stripped, size=SIZE_CAPTION, color=MUTED)
            stats["credits"] = stats.get("credits", 0) + 1
            i += 1
            continue

        # DOI 라인 — 구판: 'DOI: ' 라벨 muted + 값 파랑, 9pt, ind 340
        if stripped.startswith("DOI:"):
            p = doc.add_paragraph()
            set_para_spacing(p, 0, 160)
            p.paragraph_format.left_indent = Cm(0.6)
            val = stripped[4:].strip()
            m_link = re.match(r"\[([^\]]+)\]\([^)]*\)", val)
            if m_link:
                val = m_link.group(1)
            add_run(p, "DOI: ", size=SIZE_CAPTION, bold=True, color=FONT_COLOR)
            add_run(p, val, size=SIZE_CAPTION, bold=False, color=PRIMARY)
            stats["doi"] = stats.get("doi", 0) + 1
            i += 1
            continue

        # 일반 본문 문단 (양쪽 정렬) — *이탤릭* 범위 보존
        if stripped:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            set_para_spacing(p, 0, 100)
            add_italic_runs(p, stripped, size=SIZE_BODY, color=BODY_TEXT)
            stats["paragraphs"] += 1
        i += 1
        continue


# ---------------------------------------------------------------------------
# 검증 (QA 리포트)
# ---------------------------------------------------------------------------


def collect_all_text(doc):
    texts = []
    for p in doc.paragraphs:
        texts.append(p.text)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    texts.append(p.text)
    return texts


def inject_page_header(docx_path, script_dir):
    """docx zip에 구판 머리말(word/header1.xml + 로고 rels)을 이식한다.

    python-docx로 빌드한 신규 docx에는 머리말이 아예 없어 구판 대비
    누락되었으므로(2026-10-01 사건), 저장 직후 zip 레벨에서 구판 XML을
    그대로 복원해 표·색·EMU 값이 100% 동일하게 한다. 중복 실행 시
    재이식하지 않는다. sectPr에 headerReference + 구판 여백(top 1247 /
    header 397 twips)도 보정한다.

    커버 로고는 구판에 없으므로(2026-10-01 유저 지적 2) 넣지 않는다.
    머리말 로고는 word/media/image1.png가 필요하므로, docx에 이미지가
    하나도 없으면(그림 0장 호수) 로고를 media/image1.png로 주입한다.
    """
    template_path = Path(script_dir) / HEADER_TEMPLATE
    if not template_path.exists():
        raise FileNotFoundError(
            f"머리말 템플릿 없음: {template_path} (구판 header1.xml 필요)")
    header_xml = template_path.read_bytes()

    with zipfile.ZipFile(str(docx_path), "r") as zin:
        entries = {n: zin.read(n) for n in zin.namelist()}

    if any(re.fullmatch(r"word/header\d+\.xml", n) for n in entries):
        return False

    entries["word/header1.xml"] = header_xml

    # 신규 머리말 자산: rId1=NEB 워드마크(블루→그린 그레디언트·유저 제작
    # 2026-10-01)·rId2=logo.full.png — 태그라인 4행 문구는 워드마크로
    # 대체되어 폐지(유저 선택).
    entries["word/media/image1.png"] = (
        Path(script_dir) / "assets" / "NEB_wordmark.png").read_bytes()
    entries["word/media/image2.png"] = (
        Path(script_dir) / "assets" / "NEB_logo_full.png").read_bytes()
    entries["word/_rels/header1.xml.rels"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/'
        'package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/'
        'relationships/image" Target="media/image1.png"/>'
        '<Relationship Id="rId2" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/'
        'relationships/image" Target="media/image2.png"/></Relationships>'
    ).encode("utf-8")

    ct = entries["[Content_Types].xml"].decode("utf-8")
    if "header+xml" not in ct:
        ct = ct.replace(
            "</Types>",
            '<Override PartName="/word/header1.xml" ContentType='
            '"application/vnd.openxmlformats-officedocument.'
            'wordprocessingml.header+xml"/></Types>')
    entries["[Content_Types].xml"] = ct.encode("utf-8")

    rels = entries["word/_rels/document.xml.rels"].decode("utf-8")
    used = [int(m) for m in re.findall(r'Id="rId(\d+)"', rels)]
    rid = f"rId{max(used) + 1 if used else 1}"
    rels = rels.replace(
        "</Relationships>",
        f'<Relationship Id="{rid}" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/'
        'relationships/header" Target="header1.xml"/></Relationships>')
    entries["word/_rels/document.xml.rels"] = rels.encode("utf-8")

    doc_xml = entries["word/document.xml"].decode("utf-8")
    sect_open = re.search(r"<w:sectPr\b[^>]*>", doc_xml)
    if sect_open is None:
        raise RuntimeError("document.xml sectPr 없음 — headerReference 결련 불가")
    doc_xml = doc_xml.replace(
        sect_open.group(0),
        sect_open.group(0) +
        f'<w:headerReference w:type="default" r:id="{rid}"/>', 1)

    def _fix_pgmar(mm):
        tag = mm.group(0)
        tag = re.sub(r'\bw:top="\d+"',
                     f'w:top="{HEADER_MARGIN_TOP_TWIPS}"', tag)
        if re.search(r'\bw:header="\d+"', tag):
            tag = re.sub(r'\bw:header="\d+"',
                         f'w:header="{HEADER_MARGIN_HEADER_TWIPS}"', tag)
        else:
            tag = tag.replace(
                "<w:pgMar",
                f'<w:pgMar w:header="{HEADER_MARGIN_HEADER_TWIPS}"', 1)
        return tag

    doc_xml = re.sub(r"<w:pgMar\b[^/]*/>", _fix_pgmar, doc_xml)
    entries["word/document.xml"] = doc_xml.encode("utf-8")

    tmp_path = str(docx_path) + ".part"
    with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in entries.items():
            zout.writestr(name, data)
    os.replace(tmp_path, str(docx_path))
    return True


def check_page_header(docx_path):
    """pageHeader 3요소 검증: header xml 존재·문구 이미지·로고 이미지
    (2026-10-01 이미지화: 문구는 PNG 이미지로 주입 — 텍스트 매칭 대신
    rId1=태그라인 PNG·rId2=로고 PNG 두 Relationship 실존으로 검증)."""
    with zipfile.ZipFile(str(docx_path), "r") as z:
        names = z.namelist()
        headers = [n for n in names if re.fullmatch(r"word/header\d+\.xml", n)]
        header_ok = len(headers) >= 1
        phrase_ok = False
        logo_ok = False
        for h in headers:
            xml = z.read(h).decode("utf-8", "ignore")
            if HEADER_PHRASE in xml:
                phrase_ok = True
            base = h.split("/")
            rels_name = f"{base[0]}/_rels/{base[1]}.rels"
            if rels_name not in names:
                continue
            rels_txt = z.read(rels_name).decode("utf-8", "ignore")
            rels = re.findall(r'<Relationship Id="(rId\d+)"[^>]*'
                              r'Target="([^"]+)"', rels_txt)
            media_targets = []
            for rid, tgt in rels:
                tgt = tgt.lstrip("/")
                media = os.path.normpath(
                    os.path.join("word", tgt)).replace("\\", "/")
                media_targets.append((rid, media in names))
            # 문구(rId1)·로고(rId2) 둘 다 이미지로 주입된 상태
            d = dict((rid, ok) for rid, ok in media_targets)
            phrase_ok = phrase_ok or d.get("rId1", False)
            logo_ok = logo_ok or d.get("rId2", False)
    return {"header_ok": header_ok, "phrase_ok": phrase_ok,
            "logo_ok": logo_ok}


def verify(docx_path, stats=None):
    doc = Document(str(docx_path))

    # Markdown 잔재 검사
    remnants = {"##": 0, "**": 0, "|---": 0}
    for text in collect_all_text(doc):
        for key in remnants:
            remnants[key] += text.count(key)

    tables = len(doc.tables)
    images = len(doc.inline_shapes)

    # 커버 navy 셀 실측색 추출 — 그래디언트(문단 shd 램프) 또는 단색(tcPr shd)
    cover_fill = None
    grad_fills = []
    if tables >= 1:
        t = doc.tables[0]
        if len(t.rows) >= 1:
            cover_fill = get_cell_fill(t.cell(0, 0))
            for c in t.rows[0].cells:
                tcPr = c._tc.find(qn("w:tcPr"))
                shd = tcPr.find(qn("w:shd")) if tcPr is not None else None
                if shd is not None:
                    grad_fills.append(shd.get(qn("w:fill")))
    grad_ok = (len(grad_fills) == GRAD_STEPS
               and grad_fills[0] == GRAD_TOP
               and grad_fills[-1] == GRAD_BOT)
    cover_ok = (cover_fill == NAVY) or grad_ok

    remnant_ok = all(v == 0 for v in remnants.values())
    tables_ok = tables >= EXPECTED_MIN_TABLES
    images_ok = images == EXPECTED_IMAGES
    hdr = check_page_header(docx_path)
    header_ok = all(hdr.values())

    overall = remnant_ok and tables_ok and images_ok and cover_ok and header_ok

    return {
        "docx_path": str(docx_path),
        "remnants": remnants,
        "remnant_ok": remnant_ok,
        "tables": tables,
        "tables_ok": tables_ok,
        "images": images,
        "images_ok": images_ok,
        "cover_fill": cover_fill,
        "cover_ok": cover_ok,
        "grad_fills": grad_fills,
        "grad_ok": grad_ok,
        "header": hdr,
        "header_ok": header_ok,
        "overall": overall,
        "stats": stats or {},
    }


def print_report(rep):
    s = rep.get("stats", {})
    def mark(ok):
        return "PASS" if ok else "FAIL"

    print("=== GAZETTE docx QA 리포트 ===")
    print(f"파일: {rep['docx_path']}")
    print("[구조] "
          f"h2(파티션)={s.get('h2', 0)} h3={s.get('h3', 0)} "
          f"본문문단={s.get('paragraphs', 0)} 불릿={s.get('bullets', 0)} "
          f"인용={s.get('quotes', 0)} 구분선={s.get('hr', 0)} "
          f"테이블={rep['tables']} 이미지(플로트)={rep['images']} "
          f"캡션={s.get('captions', 0)}")
    r = rep["remnants"]
    print(f"[Markdown 잔재] '##'={r['##']} '**'={r['**']} '|---'={r['|---']} "
          f"→ {mark(rep['remnant_ok'])}")
    print(f"[표] {rep['tables']}개 (기준 ≥{EXPECTED_MIN_TABLES}) "
          f"→ {mark(rep['tables_ok'])}")
    print(f"[이미지] {rep['images']}개 (기준 ={EXPECTED_IMAGES}: 게재규칙 — 소식지는 초록만, "
          f"abstract figure 있을 때만) → {mark(rep['images_ok'])}")
    print(f"[커버 색] 기대 fill='{NAVY}' 실측='{rep['cover_fill']}' "
          f"그래디언트={len(rep.get('grad_fills', []))}steps → "
          f"{mark(rep['cover_ok'])}")
    h = rep.get("header", {})
    print(f"[머리말] header_xml={h.get('header_ok')} 문구={h.get('phrase_ok')} "
          f"로고={h.get('logo_ok')} ('{HEADER_PHRASE}') → "
          f"{mark(rep.get('header_ok'))}")
    print(f"종합: {mark(rep['overall'])}")
    return rep["overall"]


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------


def ensure_content_types(docx_path: str) -> None:
    """python-docx 재개시 전 [Content_Types].xml에 png Default 보증."""
    import zipfile, os
    tmp = docx_path + ".tmp"
    with zipfile.ZipFile(docx_path) as zin:
        ct = zin.read("[Content_Types].xml").decode("utf-8")
        if 'Extension="png"' in ct:
            return
        ct = ct.replace("</Types>", '<Default Extension="png" ContentType="image/png"/></Types>')
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for n in zin.namelist():
                zout.writestr(n, ct.encode("utf-8") if n == "[Content_Types].xml" else zin.read(n))
    os.replace(tmp, docx_path)


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="GAZETTE 독성동향 소식지 md → docx 빌드")
    parser.add_argument("--verify", action="store_true",
                        help="QA 리포트만 출력 (기존 .docx 대상)")
    parser.add_argument("--md", default=None,
                        help="빌드할 소식지 md 경로 (기본: 스크립트 폴더에서 GAZETTE_독성동향_*.md 최신)")
    args = parser.parse_args(argv)

    script_dir = Path(__file__).resolve().parent
    if args.md:
        md_path = Path(args.md).expanduser()
        if not md_path.is_absolute():
            md_path = script_dir / md_path
        if not md_path.exists():
            print(f"지정 md 없음: {md_path}", file=sys.stderr)
            return 2
    else:
        md_path = find_md(script_dir)
    docx_path = md_path.with_suffix(".docx")

    if args.verify:
        if not docx_path.exists():
            print(f"대상 .docx 없음: {docx_path}", file=sys.stderr)
            return 2
        rep = verify(docx_path)
        ok = print_report(rep)
        return 0 if ok else 1

    # 변환
    md_text = md_path.read_text(encoding="utf-8")
    month_title = parse_month_title(md_path)
    topic_m = re.match(r"GAZETTE_(.+)_\d{4}-\d{2}", md_path.stem)
    topic_label = topic_m.group(1) if topic_m else "독성동향"
    stats_text = extract_cover_stats(md_text)
    logo_path = resolve_logo(md_path)

    stats = {
        "h2": 0, "h3": 0, "paragraphs": 0, "bullets": 0,
        "quotes": 0, "hr": 0, "tables": 0, "images": 0, "captions": 0,
    }

    doc = Document()
    setup_document(doc)
    build_cover(doc, month_title, stats_text, logo_path, topic_label)
    build_body(doc, md_text, md_path, stats)

    # 문서 끝에 본문과 커버 사이 빈 문단 추가(커버 아래 본문 시작 구분)
    doc.save(str(docx_path))
    # 구판 머리말(로고+문구+구분선) 이식 — 빌드마다 자동 적용
    inject_page_header(docx_path, script_dir)

    # CT 보증(머리말 이식이 zip을 재작성하며 png Default 누락 소지) → verify 직전 시행
    ensure_content_types(str(docx_path))

    # 검증
    rep = verify(docx_path, stats=stats)
    ok = print_report(rep)
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
