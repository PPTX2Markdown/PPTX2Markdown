# 구조 분석기

[English](structure_analyzer.md) | 한국어

구조 분석기는 슬라이드 XML을 읽고, 선택적으로 placeholder 상속을 해석하고,
결정론적 읽기 순서를 계산한 뒤 진단 JSON과 재정렬 XML을 기록합니다. 메인
변환기가 이 단계를 자동으로 호출합니다.

## 모듈 구성

- `constants.py`: OOXML namespace, 태그, placeholder 상수
- `xml_primitives.py`: 공통 XML 및 bbox helper
- `text_rules.py`: 텍스트 정규화와 헤딩 문법 규칙
- `extractor.py`: XML에서 `SlideObject` 추출
- `structure.py`: `SlideObject`, XY-cut 순서, 헤딩 분석
- `pipeline.py`: 분석 보고서, 재정렬 XML, 매니페스트, 파일 입출력
- `extract_structure_analysis.py`: 진단 CLI
- `check_native_table_support.py`: 네이티브 표 점검 도구

## 읽기 순서 정책

읽기 순서는 항상 기하학적 재귀 XY-cut을 사용합니다. 텍스트, 번호, 헤딩,
placeholder 역할, 장식 label을 살펴 순서를 바꾸지 않습니다.

각 재귀 영역에서 X/Y projection을 확인합니다. 적절한 경우 촘촘한 세로 흐름을
유지하고, 그렇지 않으면 열 분할을 먼저 드러낸 뒤 아래쪽 영역을 재귀적으로
처리합니다. 좌표로 더 나눌 수 없으면 top-left 순서와 안정적인 XML index
tie-break를 사용합니다.

좌표 허용치는 작은 OOXML 반올림 차이를 흡수합니다. 저장소 PPTX 회귀 테스트는
1~3px 정렬 흔들림, 작은 양수 gutter, 1px 경계 겹침, 허용치를 넘는 겹침, 서로
다른 텍스트 상자 높이를 포함합니다. 별도의 의미론적 row-clustering 단계는
사용하지 않습니다.

## Placeholder 상속

직접 좌표나 스타일이 완전하지 않은 객체는 다음 체인에서 일치하는 placeholder
정보를 상속할 수 있습니다.

```text
slide -> slideLayout -> slideMaster
```

메인 변환기의 기본 정책은 `style` 상속과 `visible` 레이아웃/마스터 전용
도형입니다.

## 진단 CLI

하나 이상의 슬라이드 XML 처리:

```bash
python -m pptx2markdown.structure_analyzer.extract_structure_analysis \
  slide1.xml slide2.xml
```

입력을 생략하면 `<work-dir>/target_slides`를 탐색합니다.

```bash
python -m pptx2markdown.structure_analyzer.extract_structure_analysis
```

work 루트 또는 출력 디렉터리 지정:

```bash
python -m pptx2markdown.structure_analyzer.extract_structure_analysis \
  slide1.xml --work-dir .pptx2markdown --output-dir analysis
```

## 진단 출력

- `<slide>.structure_analysis.json`
- `<slide>.reordered.xml`
- `structure_analysis_manifest.json`

JSON 보고서에는 스키마 버전, 기하학적 구조 순서, 원본 XML 순서, 정렬된 XML
index, 객체 수, 신뢰도 정보, 감지한 네이티브 표와 이미지가 들어갑니다.

## 옵션

```text
--output-dir <path>
--work-dir <path>
--strict
--placeholder-inheritance none|geometry|style
--inherited-shapes none|visible|all
```

`--strict`는 헤딩 감지와 후보 threshold를 바꾸며 다른 읽기 순서 알고리즘을
선택하지 않습니다.

## 회귀 테스트

- `tests/test_xycut_reading_order.py`: 좌표 중심 단위 사례
- `tests/test_xycut_pptx_fixture.py`: 실제 PPTX 좌표 허용치 fixture
- `tests/test_xycut_layout_pptx_fixture.py`: 열, divider, 회전, 표, tie, 스키마,
  바이트 결정론
- `tests/fixtures/golden/synthetic_reading_order.pptx`: 검토된 저장소 골든 출력

승인 정책은 [골든셋 문서](../tests/fixtures/golden/README.ko.md)를 참고하세요.
