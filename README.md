# PPTX2Markdown

PPTX를 Markdown으로 변환하는 파이프라인입니다.

주요 구성:

- `main_converter`: 메인 엔드투엔드 실행기
- `surya_pipeline`: Surya 기반 reading-order 재정렬
- `structure_analyzer`: XML 기반 reading-order 분석
- `table_pipeline`: 네이티브 테이블 추출/렌더링
- `image_pipeline`: 이미지 블록을 Markdown으로 바꾸는 VLM 계층

## 이미지 처리 백엔드

이미지 블록은 세 가지 방식으로 처리됩니다.

1. 기본값: 이미지 링크 그대로 유지
2. 로컬 VLM: Qwen2.5-VL 사용
3. API VLM: Gemini API 사용

Gemini 모드는 별도 Python 패키지를 추가로 요구하지 않습니다. `urllib` 기반 REST 호출을 사용합니다.

## 요구사항

- Python `3.10` 또는 `3.11`
- LibreOffice
  - EMF/WMF 이미지를 PNG로 변환할 때 필요
- 로컬 Qwen 모드 사용 시:
  - `requirements.txt` 설치
  - GPU 환경 권장
- Gemini 모드 사용 시:
  - `GEMINI_API_KEY` 환경변수

## 설치

### 로컬 venv

```bash
python3.11 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip setuptools wheel
python -m pip install -r requirements.txt
```

### LibreOffice

- macOS

```bash
brew install libreoffice
```

- Ubuntu / Debian

```bash
sudo apt-get update
sudo apt-get install -y libreoffice libreoffice-impress fonts-dejavu-core libglib2.0-0 libgl1 libgomp1 libsm6 libxext6
```

- Windows PowerShell

```powershell
$env:SOFFICE_PATH="C:\Program Files\LibreOffice\program\soffice.exe"
```

## 입력 경로

기본 입력 폴더:

- `main_converter/target_pptx/`

기본 동작:

- 인자를 생략하면 `main_converter/target_pptx/*.pptx` 전체를 처리합니다.
- 인자를 주면 지정한 `.pptx`만 처리합니다.

예시:

```bash
python3 main_converter/convert_slides_to_md.py
python3 main_converter/convert_slides_to_md.py sample1.pptx sample2.pptx
```

## 기본 실행

```bash
python3 main_converter/convert_slides_to_md.py [INPUT_PPTX ...]
```

reading-order 지정:

```bash
python3 main_converter/convert_slides_to_md.py --reading-order xml sample1.pptx
python3 main_converter/convert_slides_to_md.py --reading-order surya sample1.pptx
```

## 이미지 VLM 옵션

공통 옵션:

- `--image-vlm-provider {local,gemini}`
- `--image-vlm-model MODEL`
- `--image-vlm-prompt "..." `
- `--image-vlm-max-new-tokens 1024`
- `--image-vlm-api-key-env GEMINI_API_KEY`

동작 규칙:

- `--image-vlm-provider local` 에서는 `--image-vlm-model`을 지정해야 이미지 변환이 켜집니다.
- `--image-vlm-provider gemini` 에서는 `--image-vlm-model`을 생략하면 기본값 `gemini-2.5-flash`를 사용합니다.
- 변환 실패 또는 비문서성 이미지인 경우 Markdown 이미지 링크로 fallback 합니다.

### 로컬 Qwen 예시

3B 별칭 사용:

```bash
python3 main_converter/convert_slides_to_md.py \
  --image-vlm-provider local \
  --image-vlm-model 3b \
  sample1.pptx
```

직접 Hugging Face 모델 ID 지정:

```bash
python3 main_converter/convert_slides_to_md.py \
  --image-vlm-provider local \
  --image-vlm-model Qwen/Qwen2.5-VL-7B-Instruct \
  sample1.pptx
```

### Gemini 예시

환경변수 설정:

```bash
export GEMINI_API_KEY="your-api-key"
```

기본 Gemini 모델 사용:

```bash
python3 main_converter/convert_slides_to_md.py \
  --image-vlm-provider gemini \
  sample1.pptx
```

명시적으로 모델 지정:

```bash
python3 main_converter/convert_slides_to_md.py \
  --image-vlm-provider gemini \
  --image-vlm-model gemini-2.5-flash \
  sample1.pptx
```

API 키 환경변수 이름을 바꾸고 싶으면:

```bash
MY_GEMINI_KEY="your-api-key" \
python3 main_converter/convert_slides_to_md.py \
  --image-vlm-provider gemini \
  --image-vlm-api-key-env MY_GEMINI_KEY \
  sample1.pptx
```

### 프롬프트 / 토큰 제어

```bash
python3 main_converter/convert_slides_to_md.py \
  --image-vlm-provider gemini \
  --image-vlm-model gemini-2.5-flash \
  --image-vlm-prompt "Custom prompt here" \
  --image-vlm-max-new-tokens 768 \
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

## Docker 사용

현재 저장소에는 [Dockerfile](/mnt/c/study/graduation/PPTX2Markdown/Dockerfile)과 [docker-compose.yml](/mnt/c/study/graduation/PPTX2Markdown/docker-compose.yml)이 포함되어 있습니다.

기본 빌드 및 실행:

```bash
docker compose build
docker compose up -d
docker compose exec app bash
```

컨테이너 쉘 진입:

```bash
docker compose exec app bash
```

로컬 `venv`처럼 계속 작업하려면, 위처럼 먼저 쉘에 들어간 뒤 그 안에서 `python3 ...` 명령을 반복 실행하면 됩니다. 아래 Docker 예시는 모두 이 셸 기준입니다.

볼륨/동작:

- 저장소 루트가 컨테이너의 `/workspace`로 마운트됩니다.
- Hugging Face / torch / pip 캐시는 compose 볼륨으로 유지됩니다.
- 기본 compose 설정은 GPU용 PyTorch wheel(`cu128`)을 설치하고 `gpus: all`로 실행합니다.
- Docker Host에는 NVIDIA driver와 NVIDIA Container Toolkit이 준비되어 있어야 합니다.

### Docker 셸에서 기본 실행

호스트에서 `main_converter/target_pptx/`에 파일을 넣은 뒤:

```bash
python3 main_converter/convert_slides_to_md.py
```

특정 파일만:

```bash
python3 main_converter/convert_slides_to_md.py sample1.pptx
```

Surya reading-order:

```bash
python3 main_converter/convert_slides_to_md.py --reading-order surya sample1.pptx
```

### Docker 셸에서 로컬 Qwen 사용

`3b` 예시:

```bash
python3 main_converter/convert_slides_to_md.py \
  --image-vlm-provider local \
  --image-vlm-model 3b \
  sample1.pptx
```

직접 Hugging Face 모델 ID 지정:

```bash
python3 main_converter/convert_slides_to_md.py \
  --image-vlm-provider local \
  --image-vlm-model Qwen/Qwen2.5-VL-7B-Instruct \
  sample1.pptx
```

주의:

- Dockerfile은 기본적으로 GPU용 PyTorch wheel(`cu128`)을 설치합니다.
- 호스트가 GPU를 컨테이너에 전달하지 못하면 `torch.cuda.is_available()`는 `False`가 됩니다.
- 큰 로컬 VLM은 GPU 없이 실행하면 매우 느릴 수 있습니다.

GPU 사용 여부 확인:

```bash
python3 -c "import torch; print(torch.__version__); print(torch.version.cuda); print(torch.cuda.is_available()); print(torch.cuda.device_count())"
```

### Docker 셸에서 Gemini API 사용

호스트 셸에 `GEMINI_API_KEY`가 이미 잡혀 있다면, 가장 간단한 방법은 키를 같이 넘겨서 쉘에 들어가는 것입니다.

```bash
docker compose exec -e GEMINI_API_KEY="$GEMINI_API_KEY" app bash
```

이미 컨테이너 쉘 안에 들어와 있다면, 먼저 키를 export:

```bash
export GEMINI_API_KEY="your-api-key"
```

그 다음 실행:

```bash
python3 main_converter/convert_slides_to_md.py \
  --image-vlm-provider gemini \
  sample1.pptx
```

모델 지정:

```bash
python3 main_converter/convert_slides_to_md.py \
  --image-vlm-provider gemini \
  --image-vlm-model gemini-2.5-flash \
  sample1.pptx
```

다른 환경변수 이름 사용:

```bash
export MY_GEMINI_KEY="$MY_GEMINI_KEY"
python3 main_converter/convert_slides_to_md.py \
  --image-vlm-provider gemini \
  --image-vlm-api-key-env MY_GEMINI_KEY \
  sample1.pptx
```

자주 쓸 거면 [docker-compose.yml](/mnt/c/study/graduation/PPTX2Markdown/docker-compose.yml)의 `environment:`에 직접 추가해도 됩니다.

예시:

```yaml
environment:
  PYTHONUNBUFFERED: "1"
  HF_HOME: /root/.cache/huggingface
  HF_HUB_DISABLE_PROGRESS_BARS: "1"
  TOKENIZERS_PARALLELISM: "false"
  GEMINI_API_KEY: ${GEMINI_API_KEY}
```

그 뒤 실행:

```bash
export GEMINI_API_KEY="your-api-key"
docker compose up -d
docker compose exec app bash
export GEMINI_API_KEY="$GEMINI_API_KEY"
python3 main_converter/convert_slides_to_md.py --image-vlm-provider gemini sample1.pptx
```

## 이미지 처리 정책

- 일반 이미지 블록은 이미지 VLM이 켜져 있으면 Markdown 변환을 시도합니다.
- 표 위에 겹쳐진 이미지와 표 셀 내부 배경 이미지도 같은 이미지 VLM 경로로 처리합니다.
- VLM이 빈 결과를 내거나 `불필요한 정보` 성격으로 판단되면 이미지 링크로 남깁니다.
- EMF/WMF는 LibreOffice로 PNG 변환 후 처리합니다.
- 투명 배경 이미지는 회색 배경으로 평탄화한 뒤 처리합니다.

## 주의사항

- 로컬 `7b` 모델은 GPU 없이 쓰기 어렵습니다.
- 로컬 Qwen 모드는 최초 실행 시 Hugging Face 캐시에 모델을 내려받습니다.
- Gemini 모드는 네트워크 연결과 유효한 API 키가 필요합니다.

## 빠른 체크

```bash
python3 -m py_compile \
  image_pipeline/service.py \
  main_converter/converter_models.py \
  main_converter/slide_converter.py \
  main_converter/convert_slides_to_md.py
```
