# structure_analyzer

`structure_analyzer/extract_structure_analysis.py`는 슬라이드 XML의 객체 순서를 분석해, 구조 분석 JSON과 재정렬된 XML을 생성합니다.

슬라이드 객체에 직접 좌표/스타일이 없고 placeholder만 있는 경우에는 PPTX 상속 체인
`slide -> slideLayout -> slideMaster`를 따라 placeholder geometry/font 정보를 보완합니다.
기본 placeholder 상속 모드는 markdown 변환 목적의 `style`이며, layout/master-only shape는
슬라이드 쇼에 보이는 `visible` shape만 본문에 포함합니다.

## 내부 구조 (리팩토링)

구조 판단 로직은 역할별 모듈로 나뉘어 있습니다.

- `constants.py`: XML namespace/태그/placeholder 상수
- `structure.py`: `SlideObject`, `OrderContext`, 읽기 순서 판단, heading depth/score 계산
- `xml_primitives.py`: XML 공통 유틸 (`local_name`, bbox 추출 등)
- `text_rules.py`: 텍스트 normalize/번호형 heading 규칙
- `extractor.py`: XML -> `SlideObject` 추출
- `pipeline.py`: 리포트 생성, XML 재정렬, 파일 입출력
- `extract_structure_analysis.py`: 구조 분석 CLI 엔트리포인트
- `check_native_table_support.py`: native table 우선 전략 점검용 보조 CLI

`extractor.py`, `pipeline.py`, `check_native_table_support.py`는 공통 XML 유틸을 공유합니다.

아래 경로 예시는 저장소 루트에서 실행하는 기준입니다.

## 입력

다음 입력 형식을 지원합니다.

- 슬라이드 XML 파일 경로(1개 이상)
  - 형식: `*.xml`
  - 예: `python3 structure_analyzer/extract_structure_analysis.py structure_analyzer/target_slides/slide1.xml`
- 입력 인자 생략
  - 동작: `./target_slides/*.xml` 전체 처리
  - 참고: 이 기본 경로는 **실행 위치(cwd)** 기준입니다.

## 출력

기본 출력 경로는 `./output` 입니다. (`--output-dir`로 변경 가능)

- 슬라이드별 구조 분석 JSON
  - `<output-dir>/<slide_stem>.structure_analysis.json`
  - 주요 필드:
    - `schema_version`
    - `structure_order`
    - `raw_xml_order`
    - `ordered_xml_indexes`
    - `counts`, `confidence`, `xml_tables`, `xml_images`
- 슬라이드별 재정렬 XML
  - `<output-dir>/<slide_stem>.reordered.xml`
- 실행 매니페스트
  - `<output-dir>/structure_analysis_manifest.json`

## 사용법

### 1) 단일/복수 파일 처리

```bash
python3 structure_analyzer/extract_structure_analysis.py structure_analyzer/target_slides/slide1.xml structure_analyzer/target_slides/slide2.xml
```

### 2) 기본 입력 경로 일괄 처리

```bash
python3 structure_analyzer/extract_structure_analysis.py
```

### 3) 출력 경로 지정

```bash
python3 structure_analyzer/extract_structure_analysis.py structure_analyzer/target_slides/slide1.xml --output-dir structure_analyzer/output
```

## XY-cut 순서

구조 분석기는 항상 XY-cut을 사용하며 텍스트, 번호, heading, placeholder 종류를
순서 결정에 사용하지 않습니다. 각 재귀 영역에서 X/Y bbox projection을 계산하고,
촘촘한 세로 흐름이 아니면 전역 X 공백으로 열을 먼저 분리합니다. X 분리가 없으면
첫 Y 공백으로 위/아래를 나눈 뒤 하위 영역을 다시 분석합니다.

## 주요 옵션 요약

- `--output-dir <path>`: 결과(JSON/XML/manifest) 저장 디렉터리
- `--strict`: XML 기반 strict heading 규칙 적용
- `--placeholder-inheritance none|geometry|style`: placeholder 상속 반영 깊이 지정
- `--inherited-shapes none|visible|all`: layout/master-only shape materialization 범위 지정
  - legacy alias: `semantic=visible`

## Strict 규칙 문서

- `--strict` 규칙/테스트 매트릭스: `structure_analyzer/STRICT_HEADING_RULES.md`
