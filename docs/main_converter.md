# main_converter

`main_converter/run_pptx_to_markdown.py`는 PPTX를 공통 중간표현으로 변환한 뒤
Markdown 또는 JSON으로 출력하는 메인 진입점입니다.

## 현재 입력 정책

1. 인자를 주면

- `.pptx` 입력은 자동으로 `<work-dir>/target_slides/<stem>/`으로 추출 후 처리
- 여러 개 혼합 입력 가능

2. 인자를 주지 않으면

- 현재 디렉터리와 `<work-dir>/target_pptx/`의 `.pptx`/`.ppt`를 자동 추출/처리

기본 `<work-dir>`은 저장소 루트의 `.pptx2markdown/`입니다. 경로 조회만으로
`target_slides`나 `target_pptx` 폴더를 만들지 않으며 실제 파일을 쓸 때만 생성합니다.

## 출력

출력 루트: `output/`

- 패키지 결과: `output/<package>/<package>.md`
- JSON 옵션 결과: `output/<package>/<package>.json`
- 매니페스트: `output/convert_manifest.json`
- 이미지 복사본: `output/<package>/media/*`

중간 산출물은 모두 `.pptx2markdown/` 아래에 모입니다.

- 압축 해제 패키지: `.pptx2markdown/target_slides/*`
- 구조 분석: `.pptx2markdown/structure_analysis/*`
- 테이블 파이프라인: `.pptx2markdown/table_pipeline/*`
- 캐시: `.pptx2markdown/.cache/*`

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
스키마 1.0은 입력 basename/형식, 양수 페이지, 닫힌 block kind, heading
불변조건, EMU bbox, slide/layout/master 출처를 규정합니다.

# Reading Order

모든 입력은 shape bbox 기반 재귀 XY-cut으로 정렬합니다. 별도의 읽기 순서 옵션은
없습니다. 구조 분석기는 column cut을 우선하고, 기하학적 분할이 불가능할 때
top-left 및 XML index fallback을 사용합니다.

## 주요 옵션

- `--output-format {markdown|json}`
- `--headings {auto|strict}` (`auto`가 기본값)

이미지는 PPTX 패키지의 media part를 그대로 복사하고 정적 asset 링크로
렌더링합니다. OCR/VLM/API 호출 경로는 없습니다.
