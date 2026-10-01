# NEB TIDE

**NEB TIDE** = **N**eural **E**ducation **B**ibliography **T**rends, **I**nsights, **D**igest, **E**ducation

NEB 월간 저널 큐레이션 소식지(독성·환경·해양 트랙)를 자동 제작하는 시스템.
이전 명칭은 GAZETTE(2026-09/10호에 사용 — 레거시 표기), 2026-11호부터 **TIDE** 로 발행.

## 트랙
- **TIDE-Tox** — 독성학 8종 저널(ToxSci·ArchTox 등) 월간 동향
- **TIDE-Eco** — 환경·해양(MPB·ES&T·AquaticTox) 월간 동향

## 파이프라인
수집(OpenAlex SEP) → papers json → 클러스터링(12개) → 6선(직전호 DOI 0중복) → md 6단 조입(표지 인용구·브리핑·시그널 심화·클러스터 표·6선 카드·발행 각주) → DOCX 빌드(gazette_docx_build.py, QA 리포트) → 3처 배포(data_lab WSL·Windows 미러·Obsidian).

## 구성
```
templates/GAZETTE_독성동향_DESIGN_v1.0.md   # 디자인 토큰(표지 그라디언트 86CCEC→092B39 등)
gazette_docx_build.py                       # md → docx 빌더 (QA 검증 포함)
references/pipeline.md                      # 월별 조립 절차
examples/GAZETTE_독성동향_2026-10_sample.md  # 10월호 발행 실물 예시
```

## 용법
```
python3 gazette_docx_build.py            # 스크립트 폴더의 TIDE_독성동향_*.md 최신 자동 빌드
python3 gazette_docx_build.py --md /path/to/TIDE_독성동향_2026-11.md
python3 gazette_docx_build.py --verify   # QA 리포트만
```
2026-09/10호는 레거시 GAZETTE_ 접두어 — `--md` 로 직접 지정.

## 디자인
- 표지 그라디언트: #86CCEC → #3B8FB7 → #092B39 (수평 24셀)
- NEB 워드마크/로고: assets (유저 제작)
- QA 3요소(구조·링크·통계) 자동 검증

License: MIT
