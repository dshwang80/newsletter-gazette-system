# Newsletter Gazette System

월간 학술 저널 동향 소식지(GAZETTE) 자동 제작 시스템 — 수집 → 원문 조립 → DOCX 빌드 → 다중 위치 배포.

**실증 사례**: 「GAZETTE 독성동향」 2026-09호·2026-10호 (examples/ 참조)

## 구성

| 경로 | 내용 |
|---|---|
| `SKILL.md` | 에이전트 스킬 정의(발행 규칙·6단 구조·배포 절차·함정) |
| `references/pipeline.md` | 수집 파이프라인 상세(OpenAlex 질의·초록 복원·클러스터링·6선 선정) |
| `references/builder-deps.md` | DOCX 빌더 의존성 |
| `scripts/gazette_docx_build.py` | DOCX 빌더(커버 그라디언트·머리말 로고·QA 리포트) |
| `assets/` | NEB 로고 4종 + word_header1_template.xml 머리말 템플릿 |
| `templates/GAZETTE_독성동향_DESIGN_v1.0.md` | 디자인 시스템 v1.0(색·서체·레이아웃 YAML 사양) |
| `examples/` | 9월호(완성형)·10월호(중복 최소화 정책 적용) 실제 호 |

## 워크플로 요약

1. **수집**: 월간 크론(매월 1일)이 8개 저널의 전월 게재 논문을 OpenAlex에서 수집 → papers_{yymm}.json
2. **조립**: 수집 데이터를 6단 구조 md로 조립(표지 인용구→브리핑→시그널 심화→클러스터 표→6선 카드→발행 각주)
3. **빌드**: `python3 scripts/gazette_docx_build.py --md GAZETTE_xxx.md` → DOCX + QA 리포트
4. **배포**: md+docx를 3개 위치(작업 디렉·미러·Obsidian 볼트)에 배포, sha256 해시 대조 검증

## 발행 규칙

- 발행월 1일 발행 = 직전월까지의 데이터만 (컷 고정)
- 당월호에 이미 게재된 논문과 DOI 0중복
- 중복 최소화 정책: 표·헤더 수치=당월 수집 전체, 시그널·6선=직전호 이후 신규 위주

## 사용법

```bash
# DOCX 빌드 (QA 리포트 자동 생성)
python3 scripts/gazette_docx_build.py --md GAZETTE_독성동향_2026-10.md

# 기존 DOCX 검증
python3 scripts/gazette_docx_build.py --verify
```

의존성: Python 3.10+, python-docx. 자세한 절차는 SKILL.md와 references/ 참조.

## 배경

이 시스템은 Hermes Agent의 kain_news 프로필에서 운용되는 월간 큐레이션 체인(수집 크론 → 소식지 제작 → 3처 배포)에서 추출한 것. 저널 8종(ToxSci·ArchTox·SETAC-ETC·ALTEX·ChemResTox·ToxLett·EHP·FrontTox) 월간 스캔 실증.
