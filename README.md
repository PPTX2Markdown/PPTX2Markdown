# pptx2markdown

PPTX를 Markdown으로 변환하는 파이프라인입니다.

## 핵심 모듈

- `main_converter`: 메인 엔드투엔드 실행기
- `surya_pipeline`: Surya reading-order 재정렬 파이프라인
- `structure_analyzer`: XML 기반 reading-order 분석기
- `table_pipeline`: PPT 네이티브 테이블 추출/파싱/렌더링
- `image_pipeline`: Qwen2.5-VL 기반 이미지 전체 Markdown 변환

## 현재 이미지 처리 정책

- `graphicFrame/a:tbl` 형태의 실제 PPT 표는 기존 `table_pipeline`으로 그대로 Markdown 표로 변환합니다.
- 표 위에 겹쳐진 overlay 이미지가 있으면 해당 셀에 링크 대신 VLM Markdown 주입을 먼저 시도합니다.
- `pic` 이미지 블록은 `--image-vlm-model` 옵션을 주면 분류 없이 전부 로컬 Qwen2.5-VL 모델로 보내 Markdown으로 변환합니다.
- 임시 디버깅용으로 VLM이 변환한 Markdown 앞에는 원본 이미지명(`image6.png` 등)을 `[image-vlm-source: ...]` 형태로 남깁니다.
- `--image-vlm-model`을 주지 않으면 이미지는 기존처럼 Markdown 이미지 링크(`![](...)`)로 남깁니다.
- 투명 배경 이미지가 들어오면 OCR/VLM 가독성을 위해 `RGB(192, 192, 192)` 배경에 합성한 뒤 처리합니다.
- 이미지 VLM 변환이 실패한 경우에는 경고를 남기고 원본 이미지 링크로 폴백합니다.

## 실행 명령

저장소 루트 기준:

```bash
python3 main_converter/convert_slides_to_md.py [--per-slide] [--reading-order {xml,surya}] [--strict] [--reuse-surya-cache] [--image-vlm-model {3b,7b}] [--image-vlm-prompt "..."] [--image-vlm-max-new-tokens 1024] [INPUT_PPTX ...]
```

### 주요 옵션

- `[INPUT_PPTX]`
  - `*.pptx` 파일을 0개 이상 전달할 수 있습니다.
  - 생략하면 `main_converter/target_pptx/*.pptx` 전체를 자동 처리합니다.
  - 개별 입력은 아래 순서로 찾습니다.
    - 전달한 경로 그대로
    - `main_converter/<INPUT_PPTX>`
    - `main_converter/target_pptx/<INPUT_PPTX>`
    - `main_converter/target_slides/<INPUT_PPTX>`
  - 파일이 없으면 즉시 종료하고, 어떤 경로들을 확인했는지 출력합니다.
- `--per-slide`
  - 패키지 통합 결과(`result.md`)와 함께 슬라이드별 Markdown(`per_slide/slideN.md`)을 생성합니다.
- `--reading-order {xml,surya}`
  - 읽기 순서 전략 선택
  - `xml`(기본): XML 기반 순서 사용
  - `surya`: Surya 파이프라인 결과(`slideN.reordered.xml`) 사용
- `--strict`
  - XML 모드에서 stricter heading 규칙을 적용합니다.
  - `surya` 모드에서는 무시됩니다.
- `--reuse-surya-cache`
  - 기존 Surya `structure_ready` 결과를 재사용합니다.
- `--image-vlm-model {3b,7b}`
  - 모든 이미지 블록을 로컬 Qwen2.5-VL 모델로 Markdown 변환합니다.
  - `3b` -> `Qwen/Qwen2.5-VL-3B-Instruct`
  - `7b` -> `Qwen/Qwen2.5-VL-7B-Instruct`
- `--image-vlm-prompt`
  - 이미지 VLM에 전달할 프롬프트를 덮어씁니다.
- `--image-vlm-max-new-tokens`
  - 이미지 1장당 최대 생성 토큰 수를 지정합니다.

## 빠른 사용 예시

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

모든 이미지를 Qwen2.5-VL 3B로 Markdown 변환:

```bash
python3 main_converter/convert_slides_to_md.py --image-vlm-model 3b sample1.pptx
```

모든 이미지를 Qwen2.5-VL 7B로 Markdown 변환:

```bash
python3 main_converter/convert_slides_to_md.py --image-vlm-model 7b sample1.pptx
```

프롬프트와 토큰 수를 직접 지정:

```bash
python3 main_converter/convert_slides_to_md.py \
  --image-vlm-model 3b \
  --image-vlm-prompt "Convert this image into concise Markdown. Return Markdown only." \
  --image-vlm-max-new-tokens 768 \
  sample1.pptx
```

### 입력 파일 규칙

- 가장 단순한 방법은 `.pptx` 파일을 `main_converter/target_pptx/`에 넣고 파일명만 넘기는 것입니다.
- 절대경로나 상대경로를 직접 넘겨도 됩니다.
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

입력 파일이 없으면 이런 식으로 즉시 종료합니다:

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

## 환경 재현

팀 기본 경로는 `env(.venv)`입니다. 로컬 환경 구축이 번거롭거나 OS 차이 이슈가 있으면 Docker를 사용합니다.

지원 Python 버전:

- `3.10`
- `3.11` 권장

### A. 로컬 venv

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

2. 가상환경 생성/활성화 + Python 의존성 설치

```bash
python3.11 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip setuptools wheel
python -m pip install -r requirements.txt
```

3. 실행

```bash
python3 main_converter/convert_slides_to_md.py --help
```

주의:

- `7b` 모델은 메모리 사용량이 커서 GPU 환경이 훨씬 현실적입니다.
- 첫 `--image-vlm-model` 실행 시 Hugging Face 캐시에 모델 가중치가 다운로드됩니다.
- EMF/WMF 이미지는 LibreOffice로 PNG 래스터화한 뒤 모델에 전달합니다.

### B. Docker

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
- 컨테이너도 기본은 CPU 런타임이므로 `7b`는 매우 느릴 수 있습니다.

## 출력 경로

- `main_converter/output/<reading-order>/<package>/result.md`
- `main_converter/output/<reading-order>/<package>/per_slide/slideN.md`
- `main_converter/output/<reading-order>/<package>/media/*`
- `main_converter/output/<reading-order>/convert_manifest.json`

`<reading-order>`는 `xml | surya`입니다.

## Surya 연동 요약

`--reading-order surya` 실행 시:

1. `main_converter/target_slides`를 Surya 슬라이드 루트로 사용
2. `surya_pipeline/run_surya_pipeline.py` 실행
3. `surya_pipeline/output/structure_ready/<package>/slideN.reordered.xml` 사용

이미지 Markdown 변환은 reading-order와 별개이며, `--image-vlm-model`을 지정했을 때만 동작합니다.
