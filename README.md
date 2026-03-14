# pptx2markdown

PPTX(또는 PPTX 추출 패키지)를 Markdown으로 변환하는 파이프라인입니다.  
현재 문서는 `surya_pipeline`을 제외한 기본(XML 기반) 파이프라인 기준입니다.

## 전체 파이프라인

기본 엔드투엔드 실행은 `main_converter`가 담당합니다.

1. 입력 수집
2. 슬라이드 구조 분석(`structure_analyzer`)
3. 슬라이드별 텍스트/이미지/테이블 변환
4. 패키지 단위 Markdown 결과 및 manifest 출력

## 모듈 구성

- `main_converter`
  - 역할: 전체 변환 오케스트레이션, 최종 Markdown 생성
  - 상세: `main_converter/README.md`
- `structure_analyzer`
  - 역할: 슬라이드 객체 읽기 순서 분석, reordered XML 생성
  - 상세: `structure_analyzer/README.md`
- `table_extractor`
  - 역할: 슬라이드 XML에서 `<a:tbl>` 추출
  - 상세: `table_extractor/README.md`
- `table_parser`
  - 역할: 테이블 XML을 JSON으로 파싱하고 Markdown/HTML/CSV 렌더링
  - 상세: `table_parser/README.md`
- `image_pipeline`
  - 역할: 이미지 블록 분류(PaddleOCR) + 표 이미지 Surya 추출
  - 상세: `image_pipeline/README.md`

## 입력/출력 요약

- 주요 입력
  - `main_converter/raw_pptx/*.pptx`
  - 또는 `main_converter/target_pptx/<package>/ppt/slides/slide*.xml`
- 주요 출력
  - `main_converter/output/xml/<package>/result.md`
  - `main_converter/output/xml/convert_manifest.json`
  - `main_converter/output/xml/<package>/media/*`
  - (옵션) `main_converter/output/xml/<package>/per_slide/slideN.md`

## 빠른 사용법

`main_converter` 디렉터리 기준으로 실행합니다.

### 1) raw_pptx 일괄 변환

```bash
cd main_converter
python3 convert_slides_to_md.py --raw --reading-order xml
```

### 2) target_pptx 패키지 일괄 변환

```bash
cd main_converter
python3 convert_slides_to_md.py --reading-order xml
```

(임시 사진 테이블 처리 로직 실행...)
```bash
cd main_converter
python3 convert_slides_to_md.py --raw --reading-order xml --image-table-pipeline
```


### 3) 특정 패키지만 변환

```bash
cd main_converter
python3 convert_slides_to_md.py target_pptx/sample1 --reading-order xml
```

### 4) 슬라이드별 파일도 같이 생성

```bash
cd main_converter
python3 convert_slides_to_md.py --raw --reading-order xml --per-slide
```

## 수동 테이블 파이프라인(선택)

테이블만 따로 검증/가공할 때 사용합니다.

1. 추출
```bash
python3 table_extractor/extract_table.py
```

2. 파싱
```bash
python3 table_parser/parse_table.py
```

3. 렌더링
```bash
python3 table_parser/tableMaker.py
```

## 참고

- Python 3.10+ 권장
- 경로는 상대경로 기준으로 작성되어 있습니다.
- `surya_pipeline` 관련 옵션/흐름은 이 README 범위에서 제외했습니다.

## Quick output check

If you want one markdown file directly:

```bash
python3 main_converter/convert_slides_to_md.py --image-table-pipeline --output-file ../out.md
```

Then check:
- `../out.md` (single merged output)
- `output/xml/<package>/result.md` (per package)
- `output/xml/convert_manifest.json` (detailed run logs)

## Docker (venv-like workflow)

For development, edit files on host and run inside Docker.
Project is bind-mounted, so source changes are reflected immediately.

### 1) Build
```bash
docker compose build
```

### 2) Start
```bash
docker compose up -d
```

### 3) Enter container
```bash
docker compose exec app bash
```

### 4) Run converters
```bash
cd /workspace/main_converter
python convert_slides_to_md.py --raw --reading-order xml --image-table-pipeline
```

```bash
cd /workspace/surya_pipeline
python run_surya_pipeline.py --surya-dir .
```

### 5) Stop
```bash
docker compose down
```

Notes:
- Rebuild required when `Dockerfile` changes.
- Rebuild not required for `.py` source edits.
- See `DOCKER.md` for the full workflow.
