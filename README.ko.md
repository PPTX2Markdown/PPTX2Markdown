# pptx2markdown

[![PyPI](https://img.shields.io/pypi/v/pptx2markdown)](https://pypi.org/project/pptx2markdown/)
[![Python](https://img.shields.io/pypi/pyversions/pptx2markdown)](https://pypi.org/project/pptx2markdown/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/LICENSE)

PowerPoint(`.pptx` / `.ppt`) 프레젠테이션을 깔끔하고 구조화된 Markdown으로 변환합니다 — RAG 파이프라인과 문서 처리를 위해 설계되었습니다.

[English README](https://github.com/PPTX2Markdown/PPTX2Markdown#readme)

## 주요 기능

- **텍스트 & 헤딩** — 슬라이드 구조, 폰트 크기, placeholder 상속(슬라이드 → 레이아웃 → 마스터)을 이용한 헤딩 레벨 추론
- **테이블** — 네이티브 PPTX 테이블을 Markdown 테이블로 렌더링
- **차트 & SmartArt** — [chart2md](https://pypi.org/project/chart2md/)와 [smartart2md](https://pypi.org/project/smartart2md/)로 Markdown 변환
- **임베디드 첨부 파일** — PDF, 오디오, Office, ZIP, OLE Packager payload를 링크된 원본 파일로 보존
- **수식** — OMML 수식을 [omml2latex](https://pypi.org/project/omml2latex/)로 LaTeX 변환
- **이미지** — PPTX 패키지에서 로컬 asset으로 복사하고 결정론적 링크 생성
- **읽기 순서** — 네이티브 도형 좌표를 이용한 결정론적 재귀 XY-cut

## 설치

```bash
pip install pptx2markdown
```

### AI 에이전트에게 설치 맡기기

환경 확인, 격리된 CLI 설치, 선택적 LibreOffice 설정, 설치 검증을
로컬 코딩 에이전트에게 맡길 수 있습니다. 아래 문장을 복사해 전달하고,
에이전트가 명령을 설명하면 내용을 확인한 뒤 승인하세요.

```text
Read https://raw.githubusercontent.com/PPTX2Markdown/PPTX2Markdown/main/.github/agent-install.md and install pptx2markdown for this machine. You may run the commands needed for installation after explaining them. Ask me before any administrator, sudo, password, system package-manager, shell-profile, PATH, or LibreOffice change.
```

Codex에서는 `/plan`으로 시작하면 설치 범위와 구형 `.ppt` 지원 필요 여부를
먼저 질문받을 수 있습니다. 전체 절차는
[`.github/agent-install.md`](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/.github/agent-install.md)에 있습니다.

**LibreOffice**(선택)는 구형 `.ppt` 입력과 EMF/WMF 이미지 변환에 사용됩니다:

```bash
# macOS
brew install libreoffice
# Ubuntu / Debian
sudo apt-get install -y libreoffice libreoffice-impress
```

Windows에서는 PowerPoint가 설치되어 있으면 `.ppt` 변환에 COM 자동화를 우선 사용하며, 없으면 `SOFFICE_PATH` 환경변수에 LibreOffice 실행 파일 경로를 지정하세요.

## 빠른 시작

```bash
# 파일 하나 변환 → ./output/deck/deck.md
pptx2markdown deck.pptx

# 현재 디렉터리의 모든 .pptx/.ppt 변환
pptx2markdown

# 출력 디렉터리 지정
pptx2markdown deck.pptx -o converted/

# Markdown 대신 표준 중간표현 JSON 출력
pptx2markdown deck.pptx --output-format json
```

Python에서:

```python
import pptx2markdown

pptx2markdown.convert("deck.pptx", output_dir="converted")
```

## 옵션

### 읽기 순서

읽기 순서는 항상 도형 bbox 기반 재귀 XY-cut으로 결정합니다. 더 이상 기하학적으로
분할할 수 없으면 top-left 순서를 사용하며, 원본 XML index는 좌표가 같을 때의
결정론적 tie-break와 bbox 누락 객체의 fallback으로만 사용합니다.

### 기타 플래그

| 플래그 | 기본값 | 설명 |
| --- | --- | --- |
| `-o, --output-dir` | `./output` | 변환 결과 출력 위치 |
| `--work-dir` | `./.pptx2markdown` | 중간 파일(추출, 캐시) 위치 |
| `--output-format` | `markdown` | 최종 출력 형식 (`markdown`/`json`) |
| `--headings` | `auto` | 헤딩 판정 방식 (`auto`/`strict`) |
| `--placeholder-inheritance` | `style` | 레이아웃/마스터 스타일 상속 범위 (`none`/`geometry`/`style`) |
| `--inherited-shapes` | `visible` | 레이아웃/마스터 도형 반영 (`none`/`visible`/`all`) |
| `--ppt-converter` | `auto` | `.ppt` 변환 백엔드 (`powerpoint`/`libreoffice`) |
| `--verbose` | off | 디버그 로깅 |

전체 목록은 `pptx2markdown --help`로 확인하세요.

## 출력

생성 경로는 하나의 공통 레이아웃을 사용합니다. 최종 문서는 `output/` 아래에만
저장하고, 압축 해제 패키지·구조 분석·테이블 파이프라인·캐시는 모두
`.pptx2markdown/` 아래에만 저장합니다. `--output-dir` 또는 `--work-dir`를
지정해도 해당 루트만 바뀌며 내부 레이아웃은 동일합니다.

```
output/
├── convert_manifest.json
└── <deck-name>/
    ├── <deck-name>.md     # 또는 <deck-name>.json
    └── media/             # 복사된 이미지 asset

.pptx2markdown/
├── target_slides/         # 압축 해제된 PPTX 패키지
├── structure_analysis/    # 재정렬 XML 및 구조 분석 파일
├── table_pipeline/        # 독립 테이블 파이프라인 산출물
└── .cache/                # 재사용 가능한 중간 캐시
```

`convert_manifest.json`에는 슬라이드별 상태, 경고, 블록 통계가 기록됩니다.
JSON 출력은 Markdown 렌더러가 사용하는 것과 동일한 `PresentationDocument`
중간표현입니다. 각 슬라이드는 읽기 순서대로 정렬된 블록을 가지며 블록에는
`kind`, `content`, `shape_id`, 선택적 `heading_level`, EMU `bbox`,
`source_part` 필드가 들어갑니다. 입력 정보는 절대경로 없이 파일명과
형식만 기록합니다.

## 문서

- [메인 컨버터 내부 구조](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/docs/main_converter.md)
- [출력 스키마](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/docs/output_schema.md)
- [구조 분석기](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/docs/structure_analyzer.md)

## 라이선스

[MIT](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/LICENSE)
