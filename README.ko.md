# pptx2markdown

[![Quality](https://github.com/PPTX2Markdown/PPTX2Markdown/actions/workflows/quality.yml/badge.svg)](https://github.com/PPTX2Markdown/PPTX2Markdown/actions/workflows/quality.yml)
[![PyPI](https://img.shields.io/pypi/v/pptx2markdown)](https://pypi.org/project/pptx2markdown/)
[![Python](https://img.shields.io/pypi/pyversions/pptx2markdown)](https://pypi.org/project/pptx2markdown/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/LICENSE)

[English](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/README.md) | 한국어

`pptx2markdown`은 PowerPoint 프레젠테이션을 검색, RAG, 문서 처리에 적합한
결정론적 Markdown 또는 JSON으로 변환합니다. 네이티브 OOXML 구조와 도형 좌표를
직접 파싱하며 OCR, 비전 모델, API 키, 클라우드 서비스가 필요하지 않습니다.

## 왜 pptx2markdown인가요?

- **하나의 정적 파이프라인** — `.pptx`의 XML, relationship, 내장 패키지 part를
  직접 파싱합니다.
- **결정론적 읽기 순서** — 의미론 또는 모델 기반 보정 없이 기하학적 재귀
  XY-cut으로 네이티브 도형을 정렬합니다.
- **하나의 출력 계약** — Markdown과 JSON을 동일한 버전 기반
  `PresentationDocument`에서 렌더링합니다.
- **풍부한 네이티브 콘텐츠** — 텍스트, 헤딩, 목록, 표, 차트, SmartArt, 수식,
  이미지, 발표자 노트, 내장 첨부파일을 보존합니다.
- **검토 가능한 회귀 테스트** — 실제 PPTX fixture를 Linux, macOS, Windows에서
  JSON 및 Markdown 골든 출력과 바이트 단위로 비교합니다.

## 지원 콘텐츠

| 콘텐츠 | 처리 방식 |
| --- | --- |
| 텍스트와 헤딩 | 텍스트를 보존하고 네이티브 placeholder 및 스타일 근거로 헤딩을 추론 |
| 목록 | 순서 있는 목록과 순서 없는 목록 구조를 보존 |
| 표 | 네이티브 PowerPoint 표를 Markdown 표로 렌더링 |
| 차트 | [chart2md](https://pypi.org/project/chart2md/)를 통해 네이티브 차트를 변환 |
| SmartArt | [smartart2md](https://pypi.org/project/smartart2md/)를 통해 다이어그램 콘텐츠를 변환 |
| 수식 | [omml2latex](https://pypi.org/project/omml2latex/)를 통해 OMML 수식을 LaTeX로 변환 |
| 이미지 | 내장 이미지 part를 결정론적인 로컬 asset 경로로 복사 |
| 발표자 노트 | 화면에 보이는 슬라이드 콘텐츠와 분리해 보존 |
| 첨부파일 | 복구 가능한 PDF, 오디오, 비디오, Office, ZIP, 3D, OLE payload를 링크 파일로 보존 |

이미지 내부 텍스트 OCR, 오디오·비디오 전사, 애니메이션 재구성, connector 의미
해석, 픽셀 단위 슬라이드 재현은 지원하지 않습니다. PowerPoint 검토 댓글도
의도적으로 제외합니다.

## 요구사항

- Python 3.12 이상
- `.pptx`: Microsoft Office 또는 LibreOffice 불필요
- 구형 `.ppt`: Windows의 Microsoft PowerPoint 또는 LibreOffice
- EMF/WMF 변환: 변환이 필요할 때 LibreOffice 사용

## 설치

### AI 에이전트에게 설치 맡기기

로컬 코딩 에이전트가 환경을 확인하고, 격리 설치 방식을 선택하며, 필요한 경우
LibreOffice를 설정하고 CLI까지 검증할 수 있습니다. 아래 프롬프트 중 하나를
전달하고 모든 권한 요청을 직접 검토하세요.

일반 에이전트:

```text
Read https://raw.githubusercontent.com/PPTX2Markdown/PPTX2Markdown/main/.github/agent-install.ko.md and install pptx2markdown for this machine. Use Korean throughout the installation. You may run the commands needed for installation after explaining them. Ask me before any administrator, sudo, password, system package-manager, shell-profile, PATH, or LibreOffice change.
```

Codex 대화형 계획:

```text
/plan
Read https://raw.githubusercontent.com/PPTX2Markdown/PPTX2Markdown/main/.github/agent-install.ko.md and install pptx2markdown for this machine. Use Korean throughout the installation. You may run the commands needed for installation after explaining them. Ask me before any administrator, sudo, password, system package-manager, shell-profile, PATH, or LibreOffice change.
```

전체 절차는 [에이전트 설치 안내](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/.github/agent-install.ko.md)를 참고하세요.

### 수동 설치

격리된 CLI는 `uv` 사용을 권장합니다.

```bash
uv tool install pptx2markdown
```

또는 `pipx`나 활성화된 가상환경을 사용합니다.

```bash
pipx install pptx2markdown
# 또는 활성화된 가상환경 안에서
python -m pip install pptx2markdown
```

LibreOffice는 선택 사항입니다.

```bash
# macOS
brew install --cask libreoffice

# Ubuntu / Debian
sudo apt-get install libreoffice libreoffice-impress
```

Windows의 `.ppt` 변환은 PowerPoint COM 자동화를 사용할 수 있으면 우선 사용하고,
그렇지 않으면 LibreOffice를 시도합니다. 특정 LibreOffice 실행 파일을 사용하려면
`SOFFICE_PATH`를 지정할 수 있습니다.

## 빠른 시작

```bash
# 파일 하나 -> ./output/deck/deck.md
pptx2markdown deck.pptx

# 현재 디렉터리의 모든 .pptx/.ppt
pptx2markdown

# 여러 파일과 사용자 지정 출력 루트
pptx2markdown first.pptx second.pptx -o converted/

# Markdown 대신 표준 JSON
pptx2markdown deck.pptx --output-format json
```

일괄 변환 시 `~$*.pptx` 형식의 임시 Office 잠금 파일은 조용히 건너뜁니다.

Python에서 사용하기:

```python
import pptx2markdown

exit_code = pptx2markdown.convert(
    "deck.pptx",
    output_dir="converted",
    output_format="markdown",
)
if exit_code != 0:
    raise RuntimeError("conversion failed")
```

Markdown 출력 예시:

```markdown
[Page_1]

# 분기 실적

- 매출 18% 증가
- 영업이익률 24% 달성

| 지역 | 매출 |
| --- | ---: |
| APAC | $12.4M |
```

## CLI 옵션

| 플래그 | 기본값 | 설명 |
| --- | --- | --- |
| `-o, --output-dir` | `./output` | 최종 Markdown/JSON과 복사된 asset |
| `--work-dir` | `./.pptx2markdown` | 압축 해제 패키지, 분석 파일, 캐시 |
| `--output-format` | `markdown` | `markdown` 또는 `json` |
| `--headings` | `auto` | `auto` 또는 placeholder 전용 `strict` 헤딩 판정 |
| `--placeholder-inheritance` | `style` | `none`, `geometry`, `style` 상속 |
| `--inherited-shapes` | `visible` | `none`, `visible`, `all` 레이아웃/마스터 도형 |
| `--ppt-converter` | `auto` | `.ppt`용 `auto`, `powerpoint`, `libreoffice` |
| `--verbose` | 꺼짐 | 디버그 로깅 |

정확한 CLI 목록은 `pptx2markdown --help`로 확인하세요.

## 출력 구조

```text
output/
├── convert_manifest.json
└── <deck-name>/
    ├── <deck-name>.md       # 또는 <deck-name>.json
    ├── media/
    └── attachments/

.pptx2markdown/
├── target_slides/
├── structure_analysis/
├── table_pipeline/
└── .cache/
```

`convert_manifest.json`에는 파일 및 슬라이드 상태, 경고, 실패, 블록 통계가
기록됩니다. JSON 출력은 패키지에 포함된 `PresentationDocument 1.0` 스키마를
따릅니다. 입력 경로는 basename만 남기고 생성 링크에는 이식 가능한 `/` 구분자를
사용합니다.

## 읽기 순서와 결정론

읽기 순서는 항상 네이티브 도형 bbox 기반 재귀 XY-cut을 사용합니다. 기하학적
영역을 열과 행으로 나누고, 더 분할할 수 없으면 top-left 순서를 사용합니다.
원본 XML 순서는 좌표가 같은 경우의 결정론적 tie-break 또는 유효한 좌표가 없는
객체의 fallback으로만 사용합니다.

동일한 입력과 옵션의 JSON 및 Markdown 출력은 지원 운영체제에서 바이트 단위로
같아야 합니다. 저장소에는 검토 가능한 PPTX fixture 17개와 두 출력 형식의 예상
결과 및 asset hash가 함께 들어 있습니다.

## 보안 모델

일반 `.pptx` 변환은 로컬에서 동작하며 네트워크나 모델을 호출하지 않습니다.
OOXML 압축 해제 단계에서는 경로 탈출, 링크, 암호화된 entry, relationship 탈출,
설정된 압축 한도를 넘는 패키지를 거부합니다. 추출된 첨부파일은 입력
프레젠테이션의 데이터이므로 열기 전에 직접 확인하세요. 구형 `.ppt`와 일부 벡터
이미지 변환은 선택한 외부 Office 변환기를 실행합니다.

## 문서

- [출력 스키마](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/docs/output_schema.ko.md)
- [메인 변환기](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/docs/main_converter.ko.md)
- [구조 분석기](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/docs/structure_analyzer.ko.md)
- [골든 PPTX 모음](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/tests/fixtures/golden/README.ko.md)

## 개발

```bash
python -m pip install -e .
python -m unittest discover -v
python scripts/update_goldens.py --check
ruff check src tests scripts
ruff format --check src tests scripts
```

파서 동작을 의도적으로 바꿀 때는 새 골든 스냅샷을 승인하기 전에 생성된
Markdown과 JSON diff를 직접 검토해야 합니다.

## 라이선스

이 프로젝트는 [MIT 라이선스](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/LICENSE)로 배포합니다.
외부 골든 fixture의 저작권 표시는
[`THIRD_PARTY_NOTICES.ko.md`](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/tests/fixtures/golden/THIRD_PARTY_NOTICES.ko.md)에 정리되어 있습니다.
