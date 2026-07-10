# image_pipeline

이미지 블록을 Markdown으로 변환하는 VLM 계층입니다. `main_converter/run_pptx_to_markdown.py`에서 옵션으로 활성화해 사용합니다.

## 빠른 사용법

### 공통 옵션

- `--image-vlm-provider {local,gemini,openai,openrouter}`
- `--image-vlm-model MODEL`
- `--image-vlm-prompt "..."`
- `--image-vlm-max-new-tokens 1024`
- `--image-vlm-api-key-env GEMINI_API_KEY`
- `--ignore-image-vlm-cache`

### 동작 규칙

- `--image-vlm-provider local`에서는 `--image-vlm-model`을 지정해야 이미지 변환이 활성화됩니다.
- `--image-vlm-provider gemini`에서는 `--image-vlm-model`을 생략하면 기본값 `gemini-2.5-flash`를 사용합니다.
- `--image-vlm-provider openai`에서는 `--image-vlm-model`을 생략하면 기본값 `gpt-4.1-mini`를 사용합니다.
- `--image-vlm-provider openrouter`에서는 `--image-vlm-model`을 생략하면 기본값 `google/gemini-2.5-flash`를 사용합니다.
- 변환 실패 또는 비문서성 이미지인 경우 Markdown 이미지 링크로 fallback 합니다.

## 캐시 동작

- 이미지 VLM 결과는 메모리 캐시와 디스크 캐시를 함께 사용합니다.
- 디스크 캐시는 저장소 루트의 `.cache/image_pipeline/` 아래에 저장됩니다.
- `markdown` 또는 `no_markdown` 결과만 디스크에 저장합니다.
- `quota exceeded` 같은 일시적 실패 결과는 디스크에 고정하지 않습니다.
- 이미지 파일 내용이 바뀌면 경로, 수정 시각, 파일 크기를 기준으로 새 캐시 키가 만들어집니다.

디스크 캐시를 무시하고 처음부터 다시 계산하려면:

```bash
python main_converter/run_pptx_to_markdown.py \
  --image-vlm-provider gemini \
  --ignore-image-vlm-cache \
  sample1.pptx
```

## 로컬 Qwen 사용

3B 별칭 사용:

```bash
python main_converter/run_pptx_to_markdown.py \
  --image-vlm-provider local \
  --image-vlm-model 3b \
  sample1.pptx
```

직접 Hugging Face 모델 ID 지정:

```bash
python main_converter/run_pptx_to_markdown.py \
  --image-vlm-provider local \
  --image-vlm-model Qwen/Qwen2.5-VL-7B-Instruct \
  sample1.pptx
```

## Gemini 사용

환경변수 설정:

```dotenv
GEMINI_API_KEY=your-api-key
```

`image_pipeline`의 Gemini 클라이언트는 프로젝트 루트의 `.env` 파일을 자동으로 읽습니다.

기본 Gemini 모델 사용:

```bash
python main_converter/run_pptx_to_markdown.py \
  --image-vlm-provider gemini \
  sample1.pptx
```

명시적으로 모델 지정:

```bash
python main_converter/run_pptx_to_markdown.py \
  --image-vlm-provider gemini \
  --image-vlm-model gemini-2.5-flash \
  sample1.pptx
```

API 키 환경변수 이름을 바꾸는 경우:

```bash
MY_GEMINI_KEY="your-api-key" \
python main_converter/run_pptx_to_markdown.py \
  --image-vlm-provider gemini \
  --image-vlm-api-key-env MY_GEMINI_KEY \
  sample1.pptx
```

## OpenAI 사용

환경변수 설정:

```dotenv
OPENAI_API_KEY=your-api-key
```

기본 OpenAI 모델 사용:

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

API 키 환경변수 이름을 바꾸는 경우:

```bash
MY_OPENAI_KEY="your-api-key" \
python main_converter/run_pptx_to_markdown.py \
  --image-vlm-provider openai \
  --image-vlm-model gpt-4.1-mini \
  --image-vlm-api-key-env MY_OPENAI_KEY \
  sample1.pptx
```

## OpenRouter 사용

환경변수 설정:

```dotenv
OPENROUTER_API_KEY=your-api-key
```

기본 OpenRouter 모델 사용:

```bash
python main_converter/run_pptx_to_markdown.py \
  --image-vlm-provider openrouter \
  sample1.pptx
```

명시적으로 모델 지정:

```bash
python main_converter/run_pptx_to_markdown.py \
  --image-vlm-provider openrouter \
  --image-vlm-model google/gemini-2.5-flash \
  sample1.pptx
```

API 키 환경변수 이름을 바꾸는 경우:

```bash
MY_OPENROUTER_KEY="your-api-key" \
python main_converter/run_pptx_to_markdown.py \
  --image-vlm-provider openrouter \
  --image-vlm-model google/gemini-2.5-flash \
  --image-vlm-api-key-env MY_OPENROUTER_KEY \
  sample1.pptx
```

### 재시도와 속도 제한 관련 환경변수

```bash
export GEMINI_MIN_REQUEST_INTERVAL_SEC=1.5
export GEMINI_MAX_RETRIES=2
export GEMINI_BASE_BACKOFF_SEC=2
export GEMINI_MAX_BACKOFF_SEC=30
```

- `429 Too Many Requests` 또는 일시적 `5xx` 응답이 오면 자동 재시도합니다.
- `GEMINI_MIN_REQUEST_INTERVAL_SEC`로 요청 간 최소 간격을 강제할 수 있습니다.
- quota가 빡빡하면 `GEMINI_MIN_REQUEST_INTERVAL_SEC=3` 또는 `5`로 늘리는 편이 안정적입니다.
- Gemini가 `quota exceeded`를 반환하면 이번 실행에서는 남은 Gemini 이미지 요청을 재시도하지 않고 즉시 fallback 합니다.
- 긴 재시도 대기 중에는 남은 시간을 로그로 출력합니다.

예시 로그:

```text
[sample1] [md-convert] Processed: slide2.xml
  [image-vlm] Disk cache hit: image4.png (provider=gemini)
  [image-vlm] Gemini request: image5.png (model=gemini-2.5-flash, attempt=1/3)
  [image-vlm] Gemini retry scheduled: image5.png (model=gemini-2.5-flash, status=429, wait=59.4s, next_attempt=2/3)
  [image-vlm] Gemini waiting: image5.png (model=gemini-2.5-flash, remaining=59.4s)
```

## 프롬프트와 토큰 제어

```bash
python main_converter/run_pptx_to_markdown.py \
  --image-vlm-provider gemini \
  --image-vlm-model gemini-2.5-flash \
  --image-vlm-prompt "Custom prompt here" \
  --image-vlm-max-new-tokens 768 \
  sample1.pptx
```

## 주의사항

- 로컬 `7b` 모델은 GPU 없이 사용하기 어렵습니다.
- 로컬 Qwen 모드는 최초 실행 시 Hugging Face 캐시에 모델을 다운로드합니다.
- Gemini 모드는 네트워크 연결과 유효한 API 키가 필요합니다.
