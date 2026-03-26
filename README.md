# pptx2markdown

PPTX를 Markdown으로 변환하는 파이프라인입니다.

## 프로젝트 개요

- `main_converter`: 메인 엔드투엔드 실행기
- `surya_pipeline`: Surya reading-order 재정렬 파이프라인
- `structure_analyzer`: XML 기반 reading-order 분석기
- `table_pipeline`: PPT 네이티브 테이블 추출/파싱/렌더링
- `image_pipeline`: Qwen2.5-VL 기반 이미지 Markdown 변환

## 목차

- [빠른 시작](#빠른-시작)
- [환경 준비](#환경-준비)
- [실행 방법](#실행-방법)
- [입력 파일 규칙](#입력-파일-규칙)
- [출력 경로](#출력-경로)
- [이미지 처리 정책](#이미지-처리-정책)
- [Docker 사용](#docker-사용)
- [Surya 연동](#surya-연동)
- [트러블슈팅](#트러블슈팅)

## 빠른 시작

가장 단순한 흐름은 이렇습니다.

1. `.pptx` 파일을 `main_converter/target_pptx/`에 넣습니다.
2. 환경을 준비합니다.
3. 아래 명령으로 실행합니다.

전체 `.pptx` 자동 처리:

```bash
python3 main_converter/convert_slides_to_md.py --image-vlm-model 3b
```

특정 파일만 처리:

```bash
python3 main_converter/convert_slides_to_md.py --image-vlm-model 3b sample1.pptx
```

## 환경 준비

지원 Python 버전:

- `3.10`
- `3.11` 권장

### 로컬 venv

1. 시스템 의존성 설치

- macOS (Homebrew)

```bash
brew install libreoffice
```

- Ubuntu/Debian

```bash
sudo apt-get update
sudo apt-get install -y libreoffice libreoffice-impress fonts-dejavu-core libglib2.0-0 libgl1 libgomp1 libsm6 libxext6
```

- Windows (PowerShell)
  - LibreOffice 설치: <https://www.libreoffice.org/download/download-libreoffice/>
  - 기본 설치 경로: `C:\Program Files\LibreOffice\program\soffice.exe`
  - 필요 시:

```powershell
$env:SOFFICE_PATH="C:\Program Files\LibreOffice\program\soffice.exe"
```

2. 가상환경 생성 및 의존성 설치

```bash
python3.11 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip setuptools wheel
python -m pip install -r requirements.txt
```

3. 설치 확인

```bash
python3 main_converter/convert_slides_to_md.py --help
```

주의:

- `7b` 모델은 GPU 환경이 훨씬 현실적입니다.
- 첫 `--image-vlm-model` 실행 시 Hugging Face 캐시에 모델이 다운로드됩니다.
- EMF/WMF 이미지는 LibreOffice로 PNG 래스터화 후 처리합니다.

## 실행 방법

기본 실행 형식:

```bash
python3 main_converter/convert_slides_to_md.py [--per-slide] [--reading-order {xml,surya}] [--strict] [--reuse-surya-cache] [--image-vlm-model {3b,7b}] [--image-vlm-prompt "..."] [--image-vlm-max-new-tokens 1024] [INPUT_PPTX ...]
```

### 자주 쓰는 예시

기본 실행:

```bash
python3 main_converter/convert_slides_to_md.py
```

특정 PPTX만 변환:

```bash
python3 main_converter/convert_slides_to_md.py sample1.pptx sample2.pptx
```

Surya reading order 사용:

```bash
python3 main_converter/convert_slides_to_md.py --reading-order surya sample1.pptx
```

이미지를 Qwen2.5-VL 3B로 Markdown 변환:

```bash
python3 main_converter/convert_slides_to_md.py --image-vlm-model 3b sample1.pptx
```

이미지를 Qwen2.5-VL 3B로 Markdown 변환:

```bash
python3 main_converter/convert_slides_to_md.py --image-vlm-model 3b sample1.pptx
```

프롬프트와 토큰 수를 직접 지정:

```bash
python3 main_converter/convert_slides_to_md.py \
  --image-vlm-model 3b \
  --image-vlm-prompt "Convert this image into concise Markdown for RAG ingestion. If the image is not useful, answer exactly: 불필요한 정보" \
  --image-vlm-max-new-tokens 768 \
  sample1.pptx
```

### 주요 옵션

- `[INPUT_PPTX]`
  - `*.pptx` 파일을 0개 이상 전달할 수 있습니다.
  - 생략하면 `main_converter/target_pptx/*.pptx` 전체를 자동 처리합니다.
- `--per-slide`
  - 슬라이드별 Markdown(`per_slide/slideN.md`)도 생성합니다.
- `--reading-order {xml,surya}`
  - `xml`: XML 기반 순서 사용
  - `surya`: Surya reorder 결과 사용
- `--strict`
  - XML 모드에서 stricter heading 규칙을 적용합니다.
- `--reuse-surya-cache`
  - 기존 Surya `structure_ready` 출력을 재사용합니다.
- `--image-vlm-model {3b,7b}`
  - 일반 이미지와 표 overlay 이미지를 로컬 Qwen2.5-VL로 Markdown 변환합니다.
- `--image-vlm-prompt`
  - 이미지 VLM 프롬프트를 덮어씁니다. 기본 프롬프트는 표 이미지를 Markdown 표로 재구성하려고 시도하고, 유의미한 문서 정보가 없으면 `불필요한 정보`를 반환하도록 유도합니다.
- `--image-vlm-max-new-tokens`
  - 이미지 1장당 최대 생성 토큰 수를 지정합니다.

## 입력 파일 규칙

입력 `.pptx`는 아래 순서로 찾습니다.

1. 전달한 경로 그대로
2. `main_converter/<INPUT_PPTX>`
3. `main_converter/target_pptx/<INPUT_PPTX>`
4. `main_converter/target_slides/<INPUT_PPTX>`

권장 방식:

- `.pptx` 파일을 `main_converter/target_pptx/`에 넣고 파일명만 넘깁니다.
- 절대경로나 상대경로를 직접 넘겨도 됩니다.

중요:

- 입력을 하나라도 명시하면 자동 전체 탐색은 하지 않습니다.
- 예: `sample3.pptx`만 넘기면 `sample1.pptx`, `sample2.pptx`가 있어도 처리하지 않습니다.

예시:

```bash
python3 main_converter/convert_slides_to_md.py --image-vlm-model 3b sample1.pptx
```

또는:

```bash
python3 main_converter/convert_slides_to_md.py --image-vlm-model 3b /workspace/main_converter/target_pptx/sample1.pptx
```

입력이 없으면 전체 처리:

```bash
python3 main_converter/convert_slides_to_md.py --image-vlm-model 3b
```

입력 파일이 없으면 즉시 종료하고 확인한 경로를 출력합니다:

```text
Input .pptx file not found.
- requested: sample3.pptx
  checked: /workspace/main_converter/sample3.pptx
  checked: /workspace/main_converter/target_pptx/sample3.pptx
...
Available .pptx files under main_converter/target_pptx:
- sample1.pptx
- sample2.pptx
```

## 출력 경로

- `main_converter/output/<reading-order>/<package>/result.md`
- `main_converter/output/<reading-order>/<package>/per_slide/slideN.md`
- `main_converter/output/<reading-order>/<package>/media/*`
- `main_converter/output/<reading-order>/convert_manifest.json`

`<reading-order>`는 `xml | surya`입니다.

## 이미지 처리 정책

- `graphicFrame/a:tbl` 형태의 실제 PPT 표는 `table_pipeline`으로 Markdown 표로 변환합니다.
- 표 위에 겹쳐진 overlay 이미지는 셀 안에서 VLM Markdown 주입을 먼저 시도합니다.
- 일반 `pic` 이미지 블록은 `--image-vlm-model` 지정 시 VLM Markdown 변환을 시도합니다.
- VLM이 변환한 Markdown 앞에는 임시 디버깅용으로 `[image-vlm-source: image6.png]` 같은 source marker를 남깁니다.
- `--image-vlm-model`을 주지 않으면 이미지는 Markdown 이미지 링크(`![](...)`)로 남깁니다.
- 투명 배경 이미지는 OCR/VLM 가독성을 위해 `RGB(192, 192, 192)` 배경에 합성한 뒤 처리합니다.
- VLM이 Markdown을 비워서 반환하면 오류로 보지 않고 `status=no_markdown`으로 기록한 뒤 이미지 링크로 폴백합니다. 이런 경우는 RAG 문서화 가치가 낮은 이미지로 간주합니다.
- 실제 VLM 오류일 때만 경고를 남기고 이미지 링크로 폴백합니다.

## Docker 사용

빌드 및 실행:

```bash
docker compose build
docker compose up -d
docker compose exec app python3 main_converter/convert_slides_to_md.py --help
```

컨테이너 진입:

```bash
docker compose exec app bash
```

참고:

- Docker 이미지는 `requirements.txt`를 기준으로 설치합니다.
- Hugging Face 캐시는 `hf_cache` 볼륨에 유지됩니다.
- 컨테이너가 CPU-only면 `7b`는 매우 느릴 수 있습니다.

## Surya 연동

`--reading-order surya` 실행 시:

1. `main_converter/target_slides`를 Surya 슬라이드 루트로 사용
2. `surya_pipeline/run_surya_pipeline.py` 실행
3. `surya_pipeline/output/structure_ready/<package>/slideN.reordered.xml` 사용

이미지 Markdown 변환은 reading-order와 별개이며, `--image-vlm-model`을 지정했을 때만 동작합니다.

## 트러블슈팅

### Docker credential 에러

WSL/SSH 환경에서는 Docker credential helper 때문에 아래와 비슷한 에러가 날 수 있습니다.

```text
error getting credentials
```

이 경우 `~/.docker/config.json`에서 `credsStore`를 제거하고 다시 빌드합니다.

예:

```bash
printf '{\n  "auths": {}\n}\n' > ~/.docker/config.json
docker compose build --no-cache
```

### WSL에서 docker 명령 자체가 없음

Docker Desktop의 `Settings > Resources > WSL Integration`에서 현재 distro를 활성화해야 합니다.

### `torchvision` 관련 ImportError

Qwen2.5-VL processor 로딩에는 `torchvision`이 필요합니다. 컨테이너는 현재 이를 포함하도록 구성되어 있으므로, 이미지 재빌드 후 다시 실행합니다.

```bash
docker compose build --no-cache
docker compose up -d --force-recreate
```

### 입력 파일을 못 찾는 경우

- 파일이 실제로 `main_converter/target_pptx/`에 있는지 확인합니다.
- 파일명을 직접 넘겼다면 오타 여부를 먼저 확인합니다.

```bash
ls -l main_converter/target_pptx
```
