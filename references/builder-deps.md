# 빌더 의존성 (gazette_docx_build.py)

- 소스: /mnt/c/dshwang/kain_article/gazette_docx_build.py (37,084자) — 패키지에 포함 배포.
- import: argparse, docx(python-docx), os, re, sys, zipfile — 표준+python-docx만.
- 의존 에셋(같은 디렉 배치 필요):
  - assets/NEB_logo.png (24,910 B)
  - assets/NEB_logo_full.png (35,251 B) — 머리말 워드마크(rId2)
  - assets/NEB_wordmark.png, neb_tagline.png — 커버/머리말 보조
  - word_header1_template.xml (3,732 B) — 머리말 헤더 XML 템플릿
- HEADER_PHRASE = "Nature · Environment · Beyond" (60행)
- resolve_logo: md 경로 base에서 assets/NEB_logo.png 탐색.
- CLI: --md <file> (빌드+QA), --verify (기존 docx 사후 검증 — 단독 시 계수 0 표기 특성).
