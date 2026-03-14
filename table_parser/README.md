# table_parser

`table_parser`는 추출된 테이블 XML을 JSON 그리드로 파싱하고, 이를 Markdown/HTML/CSV로 렌더링하는 모듈입니다.

아래 경로 예시는 저장소 루트에서 실행하는 기준입니다.

## 입력

이 모듈은 2개 스크립트로 구성됩니다.

- `parse_table.py`
  - 입력 형식: OpenXML 테이블 XML 파일(`*.xml`, `<a:tbl>` 포함)
  - 인자 생략 시: `table_extractor/extract_results/*.xml` 전체를 자동 탐색
- `tableMaker.py`
  - 입력 형식:
    - `parse_table.py`가 만든 parsed-table JSON (`rows`, `n_rows`, `n_cols` 등 포함)
    - 또는 일반 dense JSON rows (`list[list[str]]` 또는 `{"rows": [...]}`)
  - 인자 생략 시: `table_parser/parsing_results/*.json`(manifest 제외) 일괄 처리

## 출력

- `parse_table.py` 출력
  - `table_parser/parsing_results/<input_stem>_grid.json`
  - `table_parser/parsing_results/manifest.json`
  - JSON에는 `rows`, `origin_cells`, `column_widths`, `n_rows`, `n_cols` 등이 포함됩니다.
- `tableMaker.py` 출력
  - 단일 입력 모드:
    - `-o` 지정 시 해당 파일로 저장
    - 미지정 시 stdout 출력
  - 배치 모드(인자 없음):
    - `table_parser/tables/<json_stem>.md` 생성

## 사용법

### 1) XML -> JSON 파싱

```bash
python3 table_parser/parse_table.py table_extractor/extract_results/slide2_0001.xml
```

복수 파일:

```bash
python3 table_parser/parse_table.py table_extractor/extract_results/slide2_0001.xml table_extractor/extract_results/slide10_0001.xml
```

기본 입력 경로 일괄 처리:

```bash
python3 table_parser/parse_table.py
```

### 2) JSON -> Markdown/HTML/CSV 렌더링

Markdown 파일로 저장:

```bash
python3 table_parser/tableMaker.py table_parser/parsing_results/slide2_0001_grid.json --mode markdown-flat -o table_parser/tables/slide2_0001.md
```

HTML 출력:

```bash
python3 table_parser/tableMaker.py table_parser/parsing_results/slide2_0001_grid.json --mode html
```

CSV 출력:

```bash
python3 table_parser/tableMaker.py table_parser/parsing_results/slide2_0001_grid.json --mode csv
```

배치 변환(인자 없음):

```bash
python3 table_parser/tableMaker.py
```

## 주요 옵션 요약

- `tableMaker.py --mode {markdown-flat|html|csv}`: 출력 포맷 선택
- `tableMaker.py --header-rows {auto|N}`: 헤더 행 수 지정
- `tableMaker.py --fill-merged {none|horizontal|vertical|both}`: 병합셀 채움 방식
- `tableMaker.py -o <path>`: 출력 파일 저장 경로

## 종료 코드

- `0`: 성공
- `1`: 입력 없음/파싱 실패/쓰기 실패 등 오류 발생


## Image table parser module

New file: `table_parser/image_table_pipeline.py`

Purpose:
- Classify image as table/non-table with PaddleOCR layout model.
- If table, extract parsed-table JSON via Surya.
- Render markdown using existing `tableMaker` internals.

Main entrypoint:
- `extract_table_markdown_from_image(path, header_rows=1)`
