# main_converter

`main_converter/run_pptx_to_markdown.py`는 PPTX를 Markdown으로 변환하는 메인 진입점입니다.

## 현재 입력 정책

1. 인자를 주면

- `.pptx` 입력은 자동으로 `target_slides/<stem>/`으로 추출 후 처리
- 패키지 선택자(`sample1`)는 `target_slides/sample1` 패키지를 처리
- 여러 개 혼합 입력 가능

2. 인자를 주지 않으면

- `target_pptx/*.pptx` 전체를 자동 추출/처리

## 출력

출력 루트: `output/<reading-order>/`

- 패키지 결과: `output/<reading-order>/<package>/result.md`
- 매니페스트: `output/<reading-order>/convert_manifest.json`
- 이미지 복사본: `output/<reading-order>/<package>/media/*`

## 기본 실행 예시

저장소 루트 기준:

```bash
python main_converter/run_pptx_to_markdown.py
```

`main_converter` 디렉토리에서 바로 실행:

```bash
python run_pptx_to_markdown.py
```

특정 `.pptx`:

```bash
python main_converter/run_pptx_to_markdown.py sample3.pptx sample4.pptx
```

# Reading Order 모드

XML 모드(기본):

```bash
python main_converter/run_pptx_to_markdown.py --reading-order xml
```

Surya 모드:

```bash
python main_converter/run_pptx_to_markdown.py --reading-order surya
```

Surya 모드에서 특정 입력:

```bash
python main_converter/run_pptx_to_markdown.py --reading-order surya sample3.pptx sample4.pptx
```

## Surya 연동 동작

`--reading-order surya` 실행 시 `main_converter`가 내부적으로:

1. `target_slides`를 Surya 슬라이드 루트로 사용 (`target_pptx`는 `.pptx` 입력 소스로 사용)
2. `surya_pipeline/run_surya_pipeline.py` 호출
3. `output/structure_ready/<package>/slideN.reordered.xml`를 읽어 변환

기본적으로 Surya 파이프라인은 `.venv/bin/python`을 우선 사용합니다(존재 시).

## 주요 옵션

- `--reading-order {xml|surya|xycut}`
- 기본값은 `strict` heading 판정
- `--not-strict` (xml 모드 전용)
- `--reuse-surya-cache`
- `--image-table-pipeline`
