# PPTX2Markdown

PPTX를 Markdown으로 변환하는 파이프라인입니다.

## 개요

이 프로젝트는 PPTX 파일을 분석해 Markdown으로 변환합니다. 텍스트, 테이블, 이미지 블록을 각각 적절한 방식으로 처리하며, 이미지의 경우 VLM 기반 변환도 지원합니다.

### 주요 구성 요소

- `main_converter`: 메인 엔드투엔드 실행기
- `surya_pipeline`: Surya 기반 reading-order 재정렬
- `structure_analyzer`: XML 기반 reading-order 분석
- `table_pipeline`: 네이티브 테이블 추출 및 렌더링
- `image_pipeline`: 이미지 블록을 Markdown으로 변환하는 VLM 계층

## 설치

### 1. uv 기반 설치

#### 요구 사항

- uv
- LibreOffice
  - EMF/WMF 이미지를 PNG로 변환할 때 필요
- 로컬 Qwen 모드 사용 시
  - `requirements.txt` 설치 필요
  - GPU 환경 권장
- Gemini 모드 사용 시
  - 프로젝트 루트 `.env` 파일에 `GEMINI_API_KEY` 설정 필요
- OpenAI 모드 사용 시
  - 프로젝트 루트 `.env` 파일에 `OPENAI_API_KEY` 설정 필요

#### 저장소 클론 후 환경 구성

```bash
git clone https://github.com/PPTX2Markdown/PPTX2Markdown.git
cd PPTX2Markdown
uv sync
```

#### 가상환경 적용

macOS / Linux:

```bash
source .venv/bin/activate
```

Windows PowerShell:

```powershell
.venv\Scripts\Activate.ps1
```

#### LibreOffice 설치

macOS:

```bash
brew install libreoffice
```

Ubuntu / Debian:

```bash
sudo apt-get update
sudo apt-get install -y libreoffice libreoffice-impress fonts-dejavu-core libglib2.0-0 libgl1 libgomp1 libsm6 libxext6
```

Windows PowerShell:
다음 웹 페이지에서 직접 다운로드 후 환경변수 설정

- 설치 페이지 : https://www.libreoffice.org/download/download-libreoffice/?type=win-x86_64&version=25.8.4&lang=ss

```powershell
$env:SOFFICE_PATH="C:\Program Files\LibreOffice\program\soffice.exe"
```

### 2. Docker 기반 설치

#### 요구 사항

- Docker
- Docker Compose
- NVIDIA GPU를 사용할 경우
  - NVIDIA driver
  - NVIDIA Container Toolkit

#### 포함 파일

- [Dockerfile](/mnt/c/study/graduation/PPTX2Markdown/Dockerfile)
- [docker-compose.yml](/mnt/c/study/graduation/PPTX2Markdown/docker-compose.yml)

#### 구성 특징

- 저장소 루트가 컨테이너의 `/workspace`로 마운트됩니다.
- Hugging Face / torch / pip 캐시는 compose 볼륨으로 유지됩니다.
- 기본 compose 설정은 GPU용 PyTorch wheel(`cu128`)을 설치하고 `gpus: all`로 실행합니다.

#### 빌드 및 실행

```bash
docker compose build
docker compose up -d
```

#### 컨테이너 셸 진입

```bash
docker compose exec app bash
```

## 사용 방법

### 입력 경로

기본 입력 폴더:

- `main_converter/target_pptx/`

기본 동작:

- 인자를 생략하면 `main_converter/target_pptx/*.pptx` 전체를 처리합니다.
- 인자를 지정하면 해당 `.pptx` 파일만 처리합니다.

### 기본 실행

```bash
python main_converter/run_pptx_to_markdown.py [INPUT_PPTX ...]
```

예시:

```bash
python main_converter/run_pptx_to_markdown.py
python main_converter/run_pptx_to_markdown.py sample1.pptx sample2.pptx
```

### Reading Order 옵션

```bash
python main_converter/run_pptx_to_markdown.py --reading-order xml sample1.pptx
python main_converter/run_pptx_to_markdown.py --reading-order surya sample1.pptx
```

### 이미지 VLM 옵션

이미지 블록을 VLM으로 Markdown 변환하려면 `--image-vlm-provider`를 지정합니다.

- 자세한 옵션 설명은 [`image_pipeline/README.md`](image_pipeline/README.md) 참고
- 변환 실패 또는 비문서성 이미지인 경우 Markdown 이미지 링크로 fallback
- 이미지 VLM 결과는 `.cache/image_pipeline/` 디스크 캐시에 저장되어 재실행 시 재사용됨
- 캐시를 무시하고 처음부터 다시 계산하려면 `--ignore-image-vlm-cache` 사용

Gemini 사용:

프로젝트 루트 `.env`:

```dotenv
GEMINI_API_KEY=your-api-key
```

그 다음 실행:

```bash
python main_converter/run_pptx_to_markdown.py \
  --image-vlm-provider gemini \
  sample1.pptx
```

캐시를 무시하고 처음부터 다시 실행:

```bash
python main_converter/run_pptx_to_markdown.py \
  --image-vlm-provider gemini \
  --ignore-image-vlm-cache \
  sample1.pptx
```

로컬 Qwen 사용:

```bash
python main_converter/run_pptx_to_markdown.py \
  --image-vlm-provider local \
  --image-vlm-model 3b \
  sample1.pptx
```

OpenAI 사용:

프로젝트 루트 `.env`:

```dotenv
OPENAI_API_KEY=your-api-key
```

기본 모델 `gpt-4.1-mini` 사용:

```bash
python main_converter/run_pptx_to_markdown.py \
  --image-vlm-provider openai \
  sample1.pptx
```

명시적으로 모델 지정:

```bash
python main_converter/run_pptx_to_markdown.py \
  --image-vlm-provider openai \
  --image-vlm-model gpt-4.1-mini \
  sample1.pptx
```

## 출력

기본 출력 경로:

- `main_converter/output/xml/<package>/result.md`
- `main_converter/output/surya/<package>/result.md`

추가 산출물:

- `convert_manifest.json`
- 복사된 이미지 asset

이미지 VLM이 Markdown으로 변환한 결과 앞에는 디버깅용 source marker가 붙습니다.

예시:

```md
[image-vlm-source: image6.png]
```

## 주의사항

- 로컬 `7b` 모델은 GPU 없이 사용하기 어렵습니다.
- 로컬 Qwen 모드는 최초 실행 시 Hugging Face 캐시에 모델을 다운로드합니다.
- Gemini 모드는 네트워크 연결과 유효한 API 키가 필요합니다.
- Gemini가 `quota exceeded`를 반환하면 이번 실행에서는 남은 이미지 요청을 즉시 fallback 처리합니다.

## 빠른 체크

```bash
python -m py_compile \
  image_pipeline/service.py \
  main_converter/converter_models.py \
  main_converter/slide_converter.py \
  main_converter/run_pptx_to_markdown.py
```
