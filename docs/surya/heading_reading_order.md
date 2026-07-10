# Surya 기반 Reading Order와 Heading 규칙

이 문서는 `--reading-order surya`와 `--headings surya`가 어떻게 동작하는지 설명한다.

## CLI 사용법

Surya reading order 사용:

```bash
cd main_converter
python run_pptx_to_markdown.py --reading-order surya <pptx>
```

Surya layout 결과만으로 heading 판정:

```bash
cd main_converter
python run_pptx_to_markdown.py --reading-order surya --headings surya <pptx>
```

`--headings surya`는 반드시 `--reading-order surya`와 함께 사용해야 한다.

## Reading Order

Surya layout 모델은 각 page의 layout block에 `position` 값을 제공한다.

현재 pipeline은 이 값을 다음처럼 다룬다.

```text
Surya raw position
  0-based model reading order

model_position
  1-based로 변환한 내부 reading order 값
```

즉 `01_surya_bboxes.json`의 block 순서는 기본적으로 Surya가 판단한 읽기 순서이다.

하지만 최종 Markdown은 Surya block 자체를 출력하지 않고, Surya block과 매칭된 PPTX XML object를 출력한다. 따라서 reading order는 다음 과정을 거친다.

```text
1. Surya block을 model_position 순서로 순회한다.
2. 각 block과 가장 잘 맞는 XML object를 찾는다.
3. 매칭된 XML object를 해당 Surya 순서에 배치한다.
4. 매칭되지 못한 XML object는 spatial fallback으로 뒤에 추가한다.
5. build_structure_ready 단계에서 slideN.reordered.xml을 생성한다.
```

## Heading 모드

`--headings` 옵션은 두 가지 모드가 있다.

```text
auto
  기존 혼합 전략이다.
  placeholder, Surya hint, font fallback, numbering fallback 등이 함께 영향을 줄 수 있다.

surya
  Surya layout 결과 기반 heading만 허용한다.
  placeholder나 font size만으로는 heading을 만들지 않는다.
```

기본값은 `auto`이다.

## Surya Heading 조건

`--headings surya` 모드에서 어떤 텍스트 객체가 Markdown heading이 되려면 다음 조건을 모두 만족해야 한다.

```text
1. Surya block이 heading 성격의 label 또는 top_k heading 신호를 가진다.
2. 해당 Surya block이 PPTX XML object와 bbox match된다.
3. reading_order_source가 surya_match 또는 surya_region이다.
4. sidecar JSON에 surya_heading_depth_hint가 존재한다.
```

이 모드에서는 다음 신호를 단독 heading 근거로 사용하지 않는다.

```text
placeholder type
font size
numbering pattern
본문/문장 형태 fallback
```

## Heading으로 보는 Surya Label

현재 heading 계열로 보는 label:

```text
SectionHeader
Header
Title
```

구현은 label 문자열을 정규화한 뒤 heading 관련 단어가 포함되는지 본다.

예:

```text
SectionHeader -> heading
PageHeader    -> heading
Title         -> heading
Text          -> heading 아님
ListItem      -> 기본적으로 heading 아님
```

## top_k 사용 방식

일부 Surya output은 `label` 외에 `top_k`를 제공한다.

예:

```json
{
  "label": "Text",
  "top_k": {
    "SectionHeader": 0.997,
    "Text": 0.001
  }
}
```

이 경우 최종 label은 `Text`이지만, SectionHeader 후보 점수가 매우 높다. 현재 구현은 이런 block을 heading 후보로 보정한다.

현재 threshold:

```text
SURYA_HEADING_TOP_K_THRESHOLD = 0.25
```

주의할 점:

```text
1. top_k 1순위 label로 전체 label을 대체하지 않는다.
2. top_k는 heading 판정 보조 신호로만 쓴다.
3. top_k heading 점수가 낮으면 사람이 보기에는 heading이어도 surya mode에서는 heading이 되지 않는다.
```

이 원칙은 Surya-only heading 모드의 근거를 layout model 결과로 추적 가능하게 만들기 위한 것이다.

## Sidecar JSON Heading 필드

`04_normalized.json`과 `slideN.structure_analysis.json`에는 heading 관련 필드가 기록된다.

예:

```json
{
  "heading_sources": ["placeholder", "surya_label"],
  "heading_source": "placeholder",
  "surya_heading_depth_hint": 1,
  "surya_heading_primary_score": 0.884
}
```

필드 의미:

```text
heading_sources
  해당 객체에 연결된 모든 heading 신호이다.
  surya_label이 포함되어 있으면 Surya heading block과 매칭되었다는 뜻이다.

heading_source
  기존 호환성을 위한 대표 heading source이다.

surya_heading_depth_hint
  --headings surya에서 사용할 heading depth이다.

surya_heading_primary_score
  같은 페이지의 heading 후보 중 대표 heading을 고르기 위한 점수이다.
```

Surya-only heading 모드는 `heading_source` 하나만 보지 않고 `heading_sources`에 `surya_label`이 있는지 확인한다.

## Heading Depth 추론

Depth는 Surya `model_position`만으로 정하지 않는다.

`model_position`은 읽기 순서이지 heading 계층 깊이가 아니다. 따라서 같은 페이지의 Surya heading 후보들을 모아 bbox geometry로 대표 heading을 고른다.

현재 대표 heading 점수:

```text
primary_score =
  0.45 * normalized_height
+ 0.30 * normalized_width
+ 0.25 * topness
```

각 항목:

```text
normalized_height
  같은 페이지 heading 후보 중 가장 큰 높이에 대한 상대 높이

normalized_width
  slide 전체 너비에 대한 상대 너비

topness
  slide 위쪽에 가까울수록 높은 값
```

Depth 배정:

```text
primary_score가 가장 높은 Surya heading 후보 -> h1
나머지 Surya heading 후보 -> h2
```

현재는 h1/h2 중심의 단순 규칙이다. 더 깊은 depth가 필요하면 font size cluster, indentation, placeholder level, numbering depth 등을 추가 신호로 확장해야 한다.

## Markdown 렌더링

`main_converter/slide_converter.py`는 sidecar JSON의 heading hint를 읽어 Markdown heading을 만든다.

`--headings surya`에서 필요한 조건:

```text
"surya_label" in heading_sources
reading_order_source in {"surya_match", "surya_region"}
surya_heading_depth_hint가 유효함
heading score가 threshold 이상
```

조건을 만족하면 depth에 따라 다음처럼 렌더링된다.

```text
# h1 heading
## h2 heading
```

조건을 만족하지 않으면 일반 텍스트 또는 리스트로 렌더링된다.

