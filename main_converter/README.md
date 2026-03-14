# main_converter

`main_converter/convert_slides_to_md.py`는 PPTX(또는 PPTX를 풀어둔 패키지)를 슬라이드 순서대로 읽어 Markdown으로 변환하는 메인 모듈입니다.

아래 경로 예시는 저장소 루트에서 실행하는 기준입니다.

## 입력

다음 입력 형식을 지원합니다.

- 패키지 디렉터리
  - 형식: `target_pptx/<package>/ppt/slides/slide*.xml` 구조를 가진 디렉터리
  - 예: `python3 main_converter/convert_slides_to_md.py target_pptx/sample1`
- `.pptx` 파일
  - 형식: `.pptx` 파일 경로
  - 동작: 내부적으로 `target_pptx/<파일명>/`로 자동 추출 후 처리
  - 예: `python3 main_converter/convert_slides_to_md.py raw_pptx/demo.pptx`
- `--raw` 모드
  - 형식: 인자 없이 `--raw`만 전달
  - 동작: `raw_pptx/*.pptx` 전체를 읽어 자동 추출 후 처리
  - 예: `python3 main_converter/convert_slides_to_md.py --raw`
- 입력 인자 생략
  - 동작: `target_pptx/*` 아래 패키지를 모두 처리

## 출력

기본 출력 루트는 `output/<reading-order>/` 입니다.

- 패키지별 결과
  - `output/<reading-order>/<package>/result.md`
  - 내용: 해당 패키지의 모든 슬라이드를 합친 Markdown
- 변환 매니페스트
  - `output/<reading-order>/convert_manifest.json`
  - 내용: 패키지/슬라이드별 성공 여부, 블록 통계, 경고, 이미지 해석 결과
- 미디어 파일
  - `output/<reading-order>/<package>/media/*`
  - 내용: Markdown에 참조되는 이미지 파일 복사본
- 옵션
  - `--per-slide`
    - 각 슬라이드 별 개별 Markdown도 같이 제공합니다.
    - `output/<reading-order>/<package>/per_slide/slideN.md`

## 사용법

### 1) 기본 실행 (target_pptx/\*)

```bash
python3 main_converter/convert_slides_to_md.py
```

### 2) 특정 패키지만 실행

```bash
python3 main_converter/convert_slides_to_md.py target_pptx/sample1
```

### 3) `.pptx` 직접 입력

```bash
python3 main_converter/convert_slides_to_md.py raw_pptx/demo.pptx
```

### 4) raw 폴더 일괄 처리

```bash
python3 main_converter/convert_slides_to_md.py --raw
```

### 5) 슬라이드별 파일도 함께 생성

```bash
python3 main_converter/convert_slides_to_md.py --per-slide
```

### 6) 읽기 순서 모드 선택

기본은 `xml` 모드입니다.

```bash
python3 main_converter/convert_slides_to_md.py --reading-order xml
```

Surya 기반 순서를 쓰려면:

```bash
python3 main_converter/convert_slides_to_md.py --reading-order surya --surya-dir surya_pipeline
```

이미 만들어진 Surya 결과를 재사용하려면:

```bash
python3 main_converter/convert_slides_to_md.py --reading-order surya --surya-dir surya_pipeline --use-existing-surya-output
```

## 주요 옵션 요약

- `--raw`: `raw_pptx`의 `.pptx`를 자동 추출해서 처리
- `--per-slide`: `per_slide/slideN.md` 파일 추가 생성
- `--reading-order {xml|surya}`: 읽기 순서 생성 방식 선택
- `--surya-dir <path>`: Surya 파이프라인 루트 경로 지정
- `--reuse-surya-cache`: Surya 파이프라인 재실행 없이 기존 캐시 우선 사용
- `--use-existing-surya-output`: 기존 `output/structure_ready`를 바로 사용

## 종료 코드

- `0`: 전체 성공
- `1`: 하나 이상 슬라이드 변환 실패

## Image Table Pipeline (PaddleOCR + Surya)

Use this option when slide pictures may contain tables:

```bash
python3 main_converter/convert_slides_to_md.py --image-table-pipeline
```

Behavior:
- For each `pic` block, classify with PaddleOCR layout model first.
- If predicted as non-table, keep existing image markdown behavior.
- If predicted as table, run Surya table recognition and render markdown table directly.

## Quick output check

You can write one merged markdown file directly:

```bash
python3 convert_slides_to_md.py --image-table-pipeline --output-file ../../out.md
```

Check outputs:
- `../../out.md` (single merged output)
- `output/xml/<package>/result.md` (per package)
- `output/xml/convert_manifest.json` (run details)
