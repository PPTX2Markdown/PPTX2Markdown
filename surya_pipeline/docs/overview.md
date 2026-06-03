# Surya Pipeline 작업 맥락

Surya pipeline은 PPTX XML만으로는 안정적으로 판단하기 어려운 reading order와 heading hint를 보완하기 위해 만든 중간 처리 파이프라인이다. PPTX를 PDF로 렌더링한 뒤 Surya layout 모델로 화면상의 layout block을 감지하고, 그 결과를 다시 PPTX XML 객체와 매칭해 `main_converter`가 사용할 수 있는 구조로 정규화한다.

이 문서는 Surya pipeline의 배경, 전체 처리 흐름, 주요 설계 판단을 설명한다. 각 단계에서 생성되는 산출물의 자세한 의미는 [output_flow.md](./output_flow.md), Surya bbox와 PPTX XML bbox의 매칭 방식은 [matching_logic.md](./matching_logic.md), heading 및 reading order 규칙은 [heading_reading_order.md](./heading_reading_order.md), 현재 한계와 개선 후보는 [known_issues.md](./known_issues.md)에 더 자세히 정리되어 있다.

## 왜 필요한가

기존 XML 기반 변환은 PPTX XML의 객체 순서와 placeholder 정보를 주로 사용한다. 이 방식은 PPTX 내부 구조가 잘 정돈되어 있을 때 안정적이지만, 실제 발표 자료에서는 XML 구조와 사람이 보는 화면 구조가 자주 어긋난다.

예를 들어 `slide.xml`의 객체 순서는 사람이 읽는 순서와 다를 수 있다. placeholder가 아닌 일반 text box가 시각적으로는 제목처럼 쓰이기도 하고, 2-column 또는 3-column 레이아웃에서는 단순한 row-major 정렬만으로 자연스러운 읽기 순서를 얻기 어렵다. XML은 shape, text, table, image 같은 객체 단위 정보는 갖고 있지만, 렌더링된 화면에서 어떤 block이 heading인지, 어떤 순서로 읽혀야 하는지는 직접 알려주지 않는다.

Surya layout 모델은 PDF로 렌더링된 화면을 보고 layout block을 감지한다. 따라서 XML 내부 순서가 아니라 실제 화면상의 배치와 layout label을 기준으로 reading order와 heading hint를 보완할 수 있다. 다만 Surya는 PPTX 객체의 shape id, placeholder 여부, 원본 text body를 알지 못하므로, Surya 결과만으로 최종 변환을 수행하지 않고 PPTX XML과 다시 연결하는 단계가 필요하다.

heading 후보와 reading order를 어떤 기준으로 해석하는지는 [heading_reading_order.md](./heading_reading_order.md)에서 더 자세히 다룬다.

## 전체 처리 흐름

Surya pipeline의 큰 흐름은 다음과 같다.

```text
PPTX
-> PDF
-> Surya layout result
-> Surya bbox 정규화
-> PPTX XML bbox 추출
-> Surya/XML bbox matching
-> normalized reading_order
-> structure_ready
-> main_converter
```

먼저 PPTX를 PDF로 변환하고, Surya CLI가 PDF 페이지에서 layout block을 감지한다. 그 다음 Surya bbox를 PPTX 좌표계에 맞게 정규화하고, PPTX XML에서는 slide 객체들의 bbox를 추출한다. 두 bbox 집합을 비교해 Surya block이 어떤 PPTX 객체에 대응하는지 후보를 만들고, 최종 매칭 결과를 기반으로 `04_normalized.json`을 생성한다.

`04_normalized.json`은 Surya와 XML을 연결한 표준 중간 표현이다. 여기에는 page별 `reading_order`, 매칭된 shape 정보, heading hint, unmatched block, decorative object 정보가 함께 들어간다. 마지막으로 `build_structure_ready_from_normalized.py`가 이 reading order를 사용해 slide XML을 재정렬하고, `main_converter`가 직접 읽을 수 있는 `05_structure_ready/` 패키지를 만든다.

각 번호별 산출물의 파일 구조와 주요 필드는 [output_flow.md](./output_flow.md)에 정리되어 있다.

## Surya와 XML을 함께 쓰는 이유

Surya와 PPTX XML은 서로 다른 장점을 갖고 있다. Surya는 렌더링된 화면을 기준으로 block 위치와 layout label을 제공하므로 사람이 보는 구조에 가깝다. 반면 PPTX XML은 shape id, placeholder, 텍스트, 이미지, 테이블 같은 원본 객체 정보를 갖고 있어 실제 Markdown 변환에 필요한 identity와 content를 제공한다.

이 파이프라인은 두 정보를 bbox 기반으로 매칭해 결합한다. 즉 Surya에서 얻은 시각적 순서와 heading 신호를 사용하되, 최종 출력 대상은 PPTX XML 객체로 유지한다. 이 방식은 Surya가 잘못 감지한 block을 그대로 출력하거나, XML 순서만 믿고 시각적 순서를 놓치는 문제를 줄이기 위한 절충이다.

매칭 과정에서는 Surya bbox와 XML bbox를 같은 PPTX EMU 좌표계로 맞춘 뒤 overlap, containment, center distance, 크기 유사도 같은 신호를 비교한다. placeholder bbox가 slide.xml에 직접 없을 때는 layout/master inheritance를 따라가며 bbox를 보완한다. 후보 점수, decorative 판정, fallback 규칙 같은 세부 로직은 [matching_logic.md](./matching_logic.md)를 참고한다.

## 산출물 구조

산출물은 PPTX stem 단위의 package로 묶인다.

```text
surya_pipeline/output/{slideName}/
  00_raw_surya_result.json
  01_surya_bboxes.json
  02_xml_bboxes.json
  03_match_candidates.json
  04_normalized.json
  05_structure_ready/
```

번호가 붙은 구조를 쓰는 이유는 단계별 디버깅을 쉽게 하기 위해서다. Surya 원본 결과, 정규화된 bbox, XML bbox, match candidate, 최종 normalized 결과를 순서대로 확인할 수 있으므로, 특정 객체가 누락되었거나 잘못 매칭되었을 때 어느 단계에서 문제가 생겼는지 추적하기 쉽다.

최종적으로 `main_converter`가 읽는 것은 `05_structure_ready/`이다. 이 폴더에는 reordered slide XML과 sidecar JSON이 들어가며, Markdown 변환 단계는 여기서 reading order와 heading hint를 가져온다. 각 산출물의 정확한 역할과 스키마는 [output_flow.md](./output_flow.md)를 본다.

## Heading과 reading order

현재 목표는 Surya 결과를 단순 디버깅용으로 보관하는 것이 아니라, 실제 변환 로직의 입력으로 사용할 수 있는 reading order와 heading hint를 만드는 것이다. Surya의 layout label, `top_k`, block 위치, PPTX placeholder 정보가 함께 사용되며, 최종 결과는 `04_normalized.json`의 `reading_order`에 반영된다.

특히 `--headings surya` 모드에서는 placeholder/font fallback이 아니라 Surya layout 결과 기반 heading만 허용한다. 이 모드는 XML만으로 추정한 heading과 Surya가 감지한 heading을 의도적으로 구분하기 위한 것이다. heading depth, source, fallback 정책, unmatched marker가 Markdown 출력에 반영되는 방식은 [heading_reading_order.md](./heading_reading_order.md)에 정리되어 있다.

## 현재 원칙

Surya pipeline은 Surya 결과를 그대로 신뢰하지 않고, 항상 PPTX XML 객체와 매칭해서 사용한다. Surya는 화면 구조를 잘 볼 수 있지만 원본 객체 identity를 모르고, XML은 원본 객체 정보를 갖고 있지만 시각적 순서가 약하기 때문이다.

매칭 실패도 숨기지 않는다. Surya block과 매칭되지 않은 XML 객체는 `xml_append`로 남기고, `main_converter` 출력에서는 필요한 경우 `[unmatched]` marker를 붙여 문제를 드러낸다. 이는 변환 품질 문제를 조용히 누락시키기보다, 디버깅 가능한 형태로 노출하기 위한 선택이다.

장식 요소는 본문 변환 대상에서 제외할 수 있다. 다만 decorative 판정은 변환 결과에 영향을 주므로, 현재 기준과 한계는 계속 검증해야 한다. 남아 있는 문제와 개선 방향은 [known_issues.md](./known_issues.md)에 따로 모아둔다.
