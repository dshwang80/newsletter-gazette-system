---
version: "1.0 (2026-10-01)"
name: NEB TIDE
description: 독성학 저널 동향 소식지 전용 디자인 시스템. kain_sp2 전향 지침서 디자인 양식(POPS 토큰형 YAML + Components + Do/Don't)의 편집 스타일을 계승하되, 소식지 스케일(A4 편면 4쪽 내외, 월간 요약형)에 맞춰 재정의한 것 — 지침서 양식 복제 아님.
docx_spec: DESIGN_report.md(후향 지침서 docx 사양)의 페이지·폰트·검증 항목 구조를 소식지 규모로 축소·변형해 차용.
source_of_truth:
  - /mnt/c/dshwang/kain_article/POPS_전문가포럼_DESIGN.md  # 양식 원형(토큰 구조)
  - /mnt/c/dshwang/data_lab_mirror/위해성평가_방법론/08_지침서/디자인/DESIGN_report.md  # docx 사양 원형
  - kain_sp2 전향_지침서_제2권_해설서_v2.0_ppt디자인.docx  # 실측: 11표·12이미지, 색 006793·0B81B0·404040, 한컴 고딕·나눔고딕
source_doc: 볼트 "02. Research/독성 동향/월간 소식지 md"
colors:
  primary: "#3B8FB7"      # NEB 중간 청색(로고 리뉴얼 2026-09-27) — H2·논문 제목·링크
  secondary: "#86CCEC"    # NEB 하늘색 — 언더라인·머리말 구분선
  navy: "#092B39"         # NEB 짙은 남색 — titleBlock 배경·부서명·강조 텍스트
  accent: "#C00000"       # 위해·경고 표시(인용 수치 대조 등)
  highlight: "#F59E0B"    # 숫자·통계 강조
  dark: "#092B39"         # 캡션·보조 텍스트(기존 #404040 → NEB navy 통합)
  bodyText: "#1E293B"
  lightBg: "#E9F4FA"      # 표 배경·강조 상자(하늘색 유도 틴트)
  headerShade: "#D7EAF5"  # 표 헤더 음영(하늘색 유도 틴트)
  border: "#E2E8F0"
  muted: "#64748B"
  white: "#FFFFFF"
typography:
  font:
    eastAsia: "맑은 고딕"     # DESIGN_report 사양 계승 — Word 열람 안전성
    latin: "맑은 고딕"
  h1: {fontSize: 16pt, weight: Bold, align: left}      # 소식지 제목
  h2: {fontSize: 13pt, weight: Bold}                  # 주제 클러스터 제목
  h3: {fontSize: 11.5pt, weight: Bold}                # 개별 논문 제목
  body: {fontSize: 10pt, weight: Regular, align: justify}
  tableCell: {fontSize: 10pt, headerBold: true}
  caption: {fontSize: 9pt}                            # 출처·메타 표기
  sectionNum: {fontSize: 30pt, weight: Bold}          # 월 표기 대형 숫자 (POPS 5rem 대신 소식지 스케일 축소)
layout:
  page: "A4 세로 (21.0 × 29.7 cm)"
  margins: {left: "2.0cm", right: "2.0cm", top: "2.2cm", bottom: "2.0cm"}
  tableWidth: "항상 페이지 폭 100% (w:tblW 5000 pct)"
  volumeSize: "A4 4쪽 내외 (월간 요약형) — 지침서 스케일과 다름, 본 문서가 재정의"
components:
  titleBlock:
    backgroundColor: "{colors.primary}"
    textColor: "{colors.white}"
  clusterHeader:
    backgroundColor: "{colors.white}"
    textColor: "{colors.primary}"
    borderBottom: "{colors.secondary}"
  entryCard:
    backgroundColor: "{colors.white}"
    textColor: "{colors.bodyText}"
    titleColor: "{colors.primary}"
  statBar:
    backgroundColor: "{colors.lightBg}"
    textColor: "{colors.bodyText}"
  warningBadge:
    backgroundColor: "{colors.accent}"
    textColor: "{colors.white}"
  captionLine:
    backgroundColor: "{colors.white}"
    textColor: "{colors.muted}"
---

## Overview

월간 독성학 저널 소식지(요약 4쪽 이내·md 소스)의 docx 렌더용 디자인 토큰 문서. kain_sp2 전향 지침서(POPS 원형)의 **편집 스타일만 참고**하고, 소식지 규모에 맞춰 크기·구성을 재정의함.

**편집 스타일 계승점 (POPS 원형 → 소식지)**

| POPS 원형 | 이 문서(소식지) | 계승/변형 |
|---|---|---|
| YAML 토큰(colors/typography/layout/components) | 동일 구조 유지 | 값 전부 재계산 |
| 다크/라이트 샌드위치(표지·섹션 구분 다크) | 미적용 — 인쇄 소식지 특성 | 표지만 Primary 블록, 본문 백색 |
| 2열 카드 그리드 | 1열 항목 카드(제목: 요약: 출처:) | 슬라이드→문서 |
| Pretendard 단일 패밀리 | 맑은 고딕 단일 체계(DESIGN_report 계승) | Word 안전성 우선 |
| 섹션 대형 숫자(5rem) | 30pt 월·섹션 번호 | 축소 변형 |

**소식지 전용 변형 (지침서와 다른 점)**

- 지침서의 법령형 문체·조문 체계·TOC 필드 **해당 없음** — 소식지는 요약 행·클러스터 헤더·출처 라인 구성
- 표 2종만: 편집 요약 통계표(statBar), 출처 정리표(captionLine)
- 렌더 검증은 "docx→PDF→PNG→vision 0결함" 루프 사용 (render-QA)

## Components (소식지 docx 구성 요소)

### pageHeader — 모든 페이지 머리말 (2026-10-01 워드마크 개편)

- 1행 2열 테이블: 좌셀 **NEB 워드마크 PNG(rId1=word/media/image1.png**·assets/NEB_wordmark.png — 유저 제작 원본(img_264e1142a0a5.jpg) 비흰색 bbox 트림·캔버스 비율 유지·300dpi 대응·표시 크기 **65% 축소판 7.24×0.74cm**(원판 11.14×1.14cm; extent 2607264×266089 EMU)·tcPr vAlign=center), 우셀 NEB 로고(**rId2=word/media/image2.png**·logo_assets/엔이비/logo.full.png — 마크+NEB 텍스트, 표시 폭 3.30cm·vAlign=center)
- [이력] 좌셀 초판=텍스트 문구(맑은 고딕 9pt navy)→태그라인 4행 PNG(neb_tagline.png 1316×216)→**워드마크(유저 지시 '그림 생성 네가 이상해서 내가 했어'·4행 문구 폐지 선택)**
- 표 하단 경계 1pt 단선 secondary #86CCEC, 표 너비 9638 dxa 고정, 표 중앙 정렬
- header xml 1파일(header1.xml)·모든 페이지 적용(default 헤더)·문구는 docx 텍스트 없음(전량 이미지)
- 검증: rId1=media/image1.png(워드마크)·rId2=media/image2.png(로고) 해시 일치 — 누락·불일치 있으면 QA FAIL

### titleBlock — 표지 상단 영역

- Primary 풀폭 배경, 백색 텍스트
- 표지 밴드 그래디언트 (2026-10-01 유저 최종 확정·이후 월간 고정): **수평(왼쪽→오른쪽) 3-스톱 램프 #86CCEC → #3B8FB7 → #092B39**, 구현 방식=1행 24열 셀 배경(shd fill) — 문단-shd LibreOffice 세로 렌더 문제로 폐기. 검증은 실렌더 픽셀 프로파일(좌단 밝음→우단 진함, 위아래 동일)
- 행 1: "독성동향 소식지" (h1 16pt Bold, 백색)
- 행 2: 발행월 "2026년 10월호" (sectionNum 30pt Bold, 백색)
- 행 3: 이번 호 수록 요약 건수·대상 저널명 caption 9pt (lightBg #E8F1F8 위 muted 텍스트)

### clusterHeader — 주제 클러스터 헤더

- 백색 배경, Primary 텍스트(h2 13pt Bold)
- 아래에 Secondary 2pt 언더라인(borderBottom)·좌측 섹션 번호(축소 16pt)

### entryCard — 개별 논문 항목 (소식지 기본 단위)

- 백색 배경, border 1px #E2E8F0, 내부 패딩 8pt
- 제목: h3 11.5pt Bold, Primary
- 요약: body 10pt, bodyText, 양쪽 정렬 (2~3문장 요약)
- 메타: caption 9pt muted — 저널명·발행일·DOI
- 인용 수치 강조: highlight #F59E0B Bold
- 기준초과·위해 표기: accent #C00000 Bold

### 게재규칙 — 논문 그림·출처 (2026-10-01 유저 확정, 전문)

1. **소식지는 초록만으로 진행** — 논문 본문 그림을 기본 게재하지 않는다(그림 없는 초록 텍스트만).
2. **논문 그림은 abstract figure가 있을 경우에만 사용** — 본문 figure는 소식지에 넣지 않는다. abstract figure(graphical abstract·teaser)가 존재하는 논문만 그림 게재 대상.
3. **출처 표기 = '참고문헌' + DOI만** — '출처' 라벨 사용 금지, 라이선스 표기(© · CC 등) 전량 삭제. 각 논문 블록 메타 라인은 `참고문헌: [DOI](https://doi.org/...)` 형식.

- QA 반영: 빌드 verify의 이미지 기준 `EXPECTED_IMAGES`는 소식지 그림 사용 여부에 따라 갱신(현행 0 — abstract figure 미확보), 라이선스 문자열 0건 검증은 본 규칙에 연동.

### statBar — 편집 통계 상자 (표 1)

- headerShade #DCE6F1 헤더행 + Bold, 셀 lightBg 아님(백색)
- 표 너비 100pct, 셀 10pt
- 항목: 수집 건수·채택 건수·주제 클러스터 수·대상 저널

### captionLine — 페이지 하단 발행 라인

- muted 9pt, border 상단 1px, "발행: NEB 환경솔루션 사업부 · 기준일: YYYY-MM-DD"
- (용어 규칙) '출처' 표현 폐지 — 논문 저작물 출처는 '참고문헌' 라벨로만 표기(게재규칙 §3)

## Do's and Don'ts

- **Do** 토큰 참조 사용 — 색/크기 하드코딩 금지, `colors.primary` 등 토큰 이름으로 기록
- **Do** 다크/라이트는 표지(Primary 블록)와 본문(백색)만 — 지침서의 샌드위치 패턴 미적용
- **Do** 맑은 고딕 단일 폰트 체계 유지 (WSL/Windows Word 열람 안전)
- **Do** 소식지 본문은 요약형 2~3문장 — 논문 원문 단락 복제 금지
- **Don't** POPS 팔레트(#002B5C·#055076 등) 그대로 도입 금지 — 이 문서 팔레트만 사용
- **Don't** 지침서 조문 구조·TOC 필드·법령형 문체 소식지에 적용 금지
- **Don't** 글꼴·크기 임의 변형 — 11.5/10/9pt 체계만, 표 헤더 Bold만
- **Don't** Markdown 흔적(`##`, `**`, `|---`) docx 잔존 — 전수 0건 검증 (render-QA 루프)

## 검증 항목 (render-QA 연계)

1. 색 토큰과 실측 docx 색 일치(토큰 이름 → HEX 1:1)
2. 표 100pct·headerShade 적용 확인
3. h1/h2/h3/body/caption 사이즈 체계 일치
4. Markdown 흔적 0건·캡션 병합 0건
5. 출처 라인 존재·기준일 기재
