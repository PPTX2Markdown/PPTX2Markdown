# pptx2markdown

[![PyPI](https://img.shields.io/pypi/v/pptx2markdown)](https://pypi.org/project/pptx2markdown/)
[![Python](https://img.shields.io/pypi/pyversions/pptx2markdown)](https://pypi.org/project/pptx2markdown/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

PowerPoint(`.pptx` / `.ppt`) 프레젠테이션을 깔끔하고 구조화된 Markdown으로 변환합니다 — RAG 파이프라인과 문서 처리를 위해 설계되었습니다.

[English README](README.md)

## 주요 기능

- **텍스트 & 헤딩** — 슬라이드 구조, 폰트 크기, placeholder 상속(슬라이드 → 레이아웃 → 마스터)을 이용한 헤딩 레벨 추론
- **테이블** — 네이티브 PPTX 테이블을 Markdown 테이블로 렌더링
- **차트 & SmartArt** — [chart2md](https://pypi.org/project/chart2md/)와 [smartart2md](https://pypi.org/project/smartart2md/)로 Markdown 변환
- **수식** — OMML 수식을 [omml2latex](https://pypi.org/project/omml2latex/)로 LaTeX 변환
- **이미지** — asset으로 복사하거나, VLM(Gemini / OpenAI / OpenRouter / 로컬 Qwen2.5-VL)으로 Markdown 설명 생성
- **읽기 순서** — 기본은 XML 순서, 선택적으로 XY-cut 또는 [Surya](https://github.com/VikParuchuri/surya) 레이아웃 기반 재정렬

## 설치

```bash
pip install pptx2markdown
```

선택 extras (각각 PyTorch를 포함하므로 수 GB 설치):

```bash
pip install "pptx2markdown[surya]"      # --reading-order surya
pip install "pptx2markdown[local-vlm]"  # --image-vlm-provider local (Qwen2.5-VL)
pip install "pptx2markdown[all]"        # 전부
```

**LibreOffice**(선택)는 구형 `.ppt` 입력과 EMF/WMF 이미지 변환, Surya 모드의 PDF 렌더링에 사용됩니다:

```bash
# macOS
brew install libreoffice
# Ubuntu / Debian
sudo apt-get install -y libreoffice libreoffice-impress
```

Windows에서는 PowerPoint가 설치되어 있으면 `.ppt` 변환에 COM 자동화를 우선 사용하며, 없으면 `SOFFICE_PATH` 환경변수에 LibreOffice 실행 파일 경로를 지정하세요.

## 빠른 시작

```bash
# 파일 하나 변환 → ./output/xml/deck/result.md
pptx2markdown deck.pptx

# 현재 디렉터리의 모든 .pptx/.ppt 변환
pptx2markdown

# 출력 디렉터리 지정
pptx2markdown deck.pptx -o converted/
```

Python에서:

```python
import pptx2markdown

pptx2markdown.convert("deck.pptx", output_dir="converted")
```

## 옵션

### 읽기 순서

```bash
pptx2markdown deck.pptx --reading-order xml    # 기본: 슬라이드 XML 순서
pptx2markdown deck.pptx --reading-order xycut  # 도형 좌표 기반 재귀 XY-cut
pptx2markdown deck.pptx --reading-order surya  # Surya 레이아웃 모델 ([surya] extra 필요)
```

### VLM 이미지 설명

기본적으로 이미지는 asset 링크로 처리됩니다. `--image-vlm-provider`를 지정하면 문서성 이미지(테이블, 다이어그램, 스크린샷)를 Markdown 설명으로 변환합니다:

```bash
# Gemini — GEMINI_API_KEY를 .env 또는 환경변수에 설정
pptx2markdown deck.pptx --image-vlm-provider gemini

# OpenAI (기본 모델: gpt-4.1-mini) — OPENAI_API_KEY 필요
pptx2markdown deck.pptx --image-vlm-provider openai

# OpenRouter (기본 모델: google/gemini-2.5-flash) — OPENROUTER_API_KEY 필요
pptx2markdown deck.pptx --image-vlm-provider openrouter

# 로컬 Qwen2.5-VL — [local-vlm] extra 필요, GPU 권장
pptx2markdown deck.pptx --image-vlm-provider local --image-vlm-model 3b
```

VLM 결과는 `~/.cache/pptx2markdown/`에 캐시됩니다(`PPTX2MARKDOWN_CACHE_DIR`로 변경 가능). 캐시를 무시하려면 `--ignore-image-vlm-cache`를 사용하세요.

### 기타 플래그

| 플래그 | 기본값 | 설명 |
| --- | --- | --- |
| `-o, --output-dir` | `./output` | 변환된 Markdown 출력 위치 |
| `--work-dir` | `./.pptx2markdown` | 중간 파일(추출, 캐시) 위치 |
| `--not-strict` | off | 완화된 헤딩 판정 |
| `--placeholder-inheritance` | `style` | 레이아웃/마스터 스타일 상속 범위 (`none`/`geometry`/`style`) |
| `--inherited-shapes` | `visible` | 레이아웃/마스터 도형 반영 (`none`/`visible`/`all`) |
| `--ppt-converter` | `auto` | `.ppt` 변환 백엔드 (`powerpoint`/`libreoffice`) |
| `--verbose` | off | 디버그 로깅 |

전체 목록은 `pptx2markdown --help`로 확인하세요.

## 출력

```
output/
└── xml/                    # 읽기 순서 모드별 폴더
    ├── convert_manifest.json
    └── <deck-name>/
        ├── result.md
        └── media/          # 복사된 이미지 asset
```

`convert_manifest.json`에는 슬라이드별 상태, 경고, 블록 통계가 기록됩니다.

## 문서

- [메인 컨버터 내부 구조](docs/main_converter.md)
- [이미지 VLM 파이프라인](docs/image_pipeline.md)
- [구조 분석기](docs/structure_analyzer.md)
- [Surya 파이프라인](docs/surya_pipeline.md)

## 라이선스

[MIT](LICENSE)
