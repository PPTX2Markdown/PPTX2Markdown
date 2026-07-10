# Surya Pipeline

`surya_pipeline`은 PPTX를 PDF로 변환한 뒤 Surya layout 결과와 PPTX XML 객체를 매칭하여 `main_converter`가 사용할 reading order와 heading hint를 생성하는 파이프라인이다.

이 파이프라인의 최종 산출물은 다음 옵션에서 사용된다.

```bash
cd main_converter
python run_pptx_to_markdown.py --reading-order surya target_pptx/Heading_Test.pptx
```

Surya layout 결과만으로 heading을 판단하려면 다음처럼 실행한다.

```bash
cd main_converter
python run_pptx_to_markdown.py \
  --reading-order surya \
  --headings surya \
  target_pptx/Heading_Test.pptx
```

## 파이프라인 요약

입력 PPTX 하나는 PPTX stem 단위의 output package로 변환된다.

```text
PPTX
-> PDF
-> 00_raw_surya_result.json
-> 01_surya_bboxes.json
-> 02_xml_bboxes.json
-> 03_match_candidates.json
-> 04_normalized.json
-> 05_structure_ready/
```

출력 구조:

```text
surya_pipeline/output/{slideName}/
  00_raw_surya_result.json
  01_surya_bboxes.json
  02_xml_bboxes.json
  03_match_candidates.json
  04_normalized.json
  05_structure_ready/
    manifest.json
    slide1.reordered.xml
    slide1.structure_analysis.json
    slide2.reordered.xml
    slide2.structure_analysis.json
    ...
```

`05_structure_ready/`가 `main_converter --reading-order surya`에서 직접 읽는 최종 결과이다.

## 실행 방법

기본 입력 대상은 `surya_pipeline/target_pptx/` 바로 아래에 있는 `.pptx` 파일이다.
대상 이름으로 `Heading_Test`를 넘기면 `target_pptx/Heading_Test.pptx`를 찾는다.
하위 디렉터리는 재귀적으로 탐색하지 않으며, 압축해제된 PPTX XML 폴더를 기본 입력으로 받지는 않는다.

실행 중 PPTX는 `target_slides/{PPTX stem}/` 아래에 압축해제되어 XML bundle로 저장된다.
이 XML bundle은 기본적으로 중간 산출물이지만, `--prefer-existing-target-slides`를 주면 기존 bundle을 우선 재사용할 수 있다.

기본 실행:

```bash
cd surya_pipeline
python run_surya_pipeline.py
```

특정 PPTX만 실행:

```bash
cd surya_pipeline
python run_surya_pipeline.py Heading_Test
```

기존 산출물을 무시하고 다시 생성:

```bash
cd surya_pipeline
python run_surya_pipeline.py --force Heading_Test
```

다른 PPTX 입력 디렉터리 사용:

```bash
cd surya_pipeline
python run_surya_pipeline.py \
  --target-pptx-dir ../main_converter/target_pptx \
  Heading_Test
```

기존 PPTX XML bundle을 재사용:

```bash
cd surya_pipeline
python run_surya_pipeline.py \
  --target-pptx-dir ../main_converter/target_pptx \
  --target-slides-dir ../main_converter/target_slides \
  --prefer-existing-target-slides \
  Heading_Test
```

`main_converter`에서 이미 생성된 Surya cache를 재사용:

```bash
cd main_converter
python run_pptx_to_markdown.py \
  --reading-order surya \
  --reuse-surya-cache \
  target_pptx/Heading_Test.pptx
```

## 문서

세부 설계와 현재 작업 맥락은 `docs/` 아래 문서로 분리해두었다.

| 문서 | 내용 |
| --- | --- |
| [docs/overview.md](./docs/overview.md) | 현재 작업 맥락, 목표, Surya pipeline이 필요한 이유 |
| [docs/output_flow.md](./docs/output_flow.md) | 00~05 산출물의 생성 흐름과 각 JSON/XML의 의미 |
| [docs/matching_logic.md](./docs/matching_logic.md) | Surya bbox와 PPTX XML bbox를 비교하고 매칭하는 방식 |
| [docs/heading_reading_order.md](./docs/heading_reading_order.md) | Surya 기반 reading order와 heading 판정 규칙 |
| [docs/known_issues.md](./docs/known_issues.md) | 현재 남아있는 문제, 원인 분석, 개선 방향 |

기존 heading 규칙 문서였던 [surya_heading_rules.md](./surya_heading_rules.md)는 호환용 링크 문서로 유지하며, 실제 내용은 [docs/heading_reading_order.md](./docs/heading_reading_order.md)에 정리한다.

## 구현 파일

주요 구현 파일은 다음과 같다.

```text
run_surya_pipeline.py
  PPTX -> PDF -> Surya CLI -> normalized -> structure_ready 전체 실행을 담당한다.

normalize_surya_results.py
  Surya raw result와 PPTX XML을 읽어 01~04 산출물을 만든다.

layout_matcher.py
  PPTX XML bbox 추출, Surya bbox 정규화, bbox 비교, candidate scoring,
  최종 matching, heading hint 계산, slide XML reorder helper를 담은 공통 로직이다.

build_structure_ready_from_normalized.py
  04_normalized.json의 reading_order를 사용해 PPTX slide XML을 재정렬하고,
  main_converter가 읽을 수 있는 reordered XML과 sidecar JSON으로 변환한다.
```

`main_converter` 쪽 연결은 다음 파일에서 처리한다.

```text
main_converter/reading_order_pipeline.py
  surya_pipeline/output/{slideName}/05_structure_ready 위치를 찾는다.

main_converter/run_pptx_to_markdown.py
  --reading-order surya, --headings surya 옵션을 받는다.

main_converter/slide_converter.py
  sidecar JSON의 heading hint와 unmatched marker를 Markdown 렌더링에 반영한다.
```
