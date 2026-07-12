# main_converter

`main_converter/run_pptx_to_markdown.py`는 PPTX를 공통 중간표현으로 변환한 뒤
Markdown 또는 JSON으로 출력하는 메인 진입점입니다.

## 현재 입력 정책

1. 인자를 주면

- `.pptx` 입력은 자동으로 `target_slides/<stem>/`으로 추출 후 처리
- 패키지 선택자(`sample1`)는 `target_slides/sample1` 패키지를 처리
- 여러 개 혼합 입력 가능

2. 인자를 주지 않으면

- `target_pptx/*.pptx` 전체를 자동 추출/처리

## 출력

출력 루트: `output/`

- 패키지 결과: `output/<package>/<package>.md`
- JSON 옵션 결과: `output/<package>/<package>.json`
- 매니페스트: `output/convert_manifest.json`
- 이미지 복사본: `output/<package>/media/*`

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

JSON 중간표현을 최종 출력으로 저장:

```bash
pptx2markdown sample3.pptx --output-format json
```

슬라이드 변환기는 `PresentationDocument` 안에 `SlideDocument`와 순서가 보장된
`ContentBlock`을 생성합니다. Markdown 출력은 이 중간표현만 해석하므로 JSON과
Markdown 출력 경로가 별도의 파싱 로직을 갖지 않습니다.

# Reading Order

모든 입력은 shape bbox 기반 재귀 XY-cut으로 정렬합니다. 별도의 읽기 순서 옵션은
없습니다. 구조 분석기는 column cut을 우선하고, 기하학적 분할이 불가능할 때
top-left 및 XML index fallback을 사용합니다.

## 주요 옵션

- `--output-format {markdown|json}`
- `--headings {auto|strict}` (`auto`가 기본값)

이미지는 PPTX 패키지의 media part를 그대로 복사하고 정적 asset 링크로
렌더링합니다. OCR/VLM/API 호출 경로는 없습니다.
