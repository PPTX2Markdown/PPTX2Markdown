# Surya Pipeline 산출물 흐름

이 문서는 `surya_pipeline/output/{slideName}/` 아래에 생성되는 00~05 산출물의 의미를 설명한다.

## 전체 구조

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
    ...
```

## 00_raw_surya_result.json

생성 단계:

```text
PDF -> Surya layout CLI
```

Surya layout 모델이 PDF 페이지에서 감지한 원본 layout block 결과이다.

주요 필드:

```text
page
  PDF page 번호이다.

bboxes
  해당 page의 layout block 목록이다.

label
  Surya가 최종 부여한 layout label이다.
  예: SectionHeader, Text, ListItem, Table, Figure, PageFooter

confidence
  label에 대한 모델 신뢰도이다.

position
  Surya가 판단한 0-based reading order이다.

top_k
  일부 Surya output에 포함되는 label 후보 점수이다.
  현재 구현에서는 heading 후보 보정에 사용한다.

bbox
  PDF 이미지 좌표계의 bbox이다.
  원본 형식은 [x1, y1, x2, y2]이다.
```

Surya CLI는 원래 `results.json`을 생성한다. 파이프라인은 이 파일을 즉시 `00_raw_surya_result.json`으로 바꿔 저장한다.

## 01_surya_bboxes.json

생성 단계:

```text
00_raw_surya_result.json -> Surya bbox 정규화
```

Surya 원본 block을 PPTX XML 객체와 비교할 수 있도록 정규화한 결과이다.

주요 필드:

```text
block_id
  page와 block 순서를 조합한 식별자이다.
  예: p2_b3

model_position
  Surya 원본 position을 1-based로 변환한 값이다.
  Surya가 판단한 reading order를 나타낸다.

label
  Surya 원본 label이다.

top_k
  Surya 원본 top_k 후보 점수이다.

effective_role
  label/top_k를 해석해 만든 내부 role이다.
  예: heading, text, list, table, figure, footer

surya_heading_top_k_score
  top_k 중 SectionHeader/Header/Title 계열 점수의 최댓값이다.

bbox_px
  PDF 이미지 좌표계 bbox이다.
  저장 형식은 [x, y, w, h]이다.

bbox_emu
  PPTX EMU 좌표계 bbox이다.
  XML object bbox와 직접 비교하는 값이다.
```

이 파일의 순서는 기본적으로 Surya model position, 즉 Surya가 판단한 block 읽기 순서를 따른다.

## 02_xml_bboxes.json

생성 단계:

```text
PPTX slide XML -> XML object bbox 추출
```

PPTX 내부 XML에 존재하는 객체를 bbox화한 결과이다.

대상 객체:

```text
sp
  일반 shape 또는 text box이다.

pic
  이미지 객체이다.

graphicFrame
  table, chart, SmartArt 등에 사용된다.

grpSp
  group shape이다.

cxnSp
  connector line이다.
```

주요 필드:

```text
shape_id
  PPTX XML 객체 id이다.

xml_index
  slide.xml의 spTree 안 원래 순서이다.

tag
  XML 객체 tag이다.

name
  PPTX 객체 이름이다.

ph_type
  placeholder type이다.
  예: title, ctrTitle, subTitle, sldNum

text
  객체에서 추출된 텍스트이다.

font_pt
  XML에서 추출 가능한 최대 font size이다.

bbox
  PPTX EMU 좌표계 bbox이다.
  저장 형식은 [x, y, w, h]이다.

bbox_source
  bbox 출처이다.
  예: slide_xml, placeholder_inheritance, missing

bbox_source_part
  bbox를 읽은 XML 파일 경로이다.

is_title_placeholder
  title/ctrTitle placeholder 여부이다.

is_subtitle_placeholder
  subTitle placeholder 여부이다.

is_decorative
  본문 Markdown으로 출력하지 않아도 되는 장식 요소 여부이다.
```

placeholder가 slide.xml에 직접 bbox를 갖고 있지 않으면 slide layout/master를 따라가 bbox를 상속받는다. 이때 `bbox_source`는 `placeholder_inheritance`가 된다.

## 03_match_candidates.json

생성 단계:

```text
01_surya_bboxes.json + 02_xml_bboxes.json -> 후보 점수 계산
```

각 Surya block이 어떤 XML object와 매칭될 수 있는지 후보를 점수순으로 기록한다.

주요 점수 필드:

```text
score
  최종 match score이다.

overlap_min
  두 bbox의 교집합 면적을 더 작은 bbox 면적으로 나눈 값이다.

layout_coverage
  Surya block 면적 중 XML object와 겹친 비율이다.

object_coverage
  XML object 면적 중 Surya block과 겹친 비율이다.

center_score
  두 bbox 중심점이 가까울수록 높아지는 점수이다.

compatibility
  Surya role과 XML object 특성의 궁합 보정값이다.

reason
  compatibility 보정 이유이다.
  예: heading_title_placeholder, table_graphicFrame, text_shape
```

unmatched 원인을 분석할 때 가장 먼저 봐야 하는 파일이다. 어떤 객체가 `xml_append`로 떨어졌다면 이 파일에서 후보 점수가 threshold를 넘지 못했는지, 또는 애초에 후보가 없었는지 확인한다.

## 04_normalized.json

생성 단계:

```text
03_match_candidates.json -> 최종 reading order와 heading hint 결정
```

Surya block과 XML object의 최종 매칭 결과를 바탕으로 normalized reading order를 만든 결과이다.

주요 필드:

```text
pages
  slide별 normalized 결과이다.

layout_blocks
  해당 slide의 Surya block 목록이다.

xml_objects
  해당 slide의 XML object 목록이다.

match_candidates
  해당 slide의 후보 점수표이다.

reading_order
  최종 reading order row 목록이다.

unmatched_layout_blocks
  어떤 XML object와도 매칭되지 못한 Surya block이다.

decorative_objects
  connector line 등 본문 변환 대상에서 제외한 장식 객체이다.
```

`reading_order` row의 주요 필드:

```text
order_index
  최종 읽기 순서이다.

shape_id
  매칭된 XML object id이다.

reading_order_source
  row가 생성된 경로이다.
  예: surya_match, surya_region, xml_append

label
  연결된 Surya block label이다.

semantic_role
  파이프라인 내부 role이다.

match_score
  최종 선택된 match score이다.

match_quality
  strong, medium, weak, unmatched 중 하나이다.

heading_sources
  heading 판정에 사용된 신호 목록이다.
  예: placeholder, surya_label

surya_heading_depth_hint
  `--headings surya`에서 사용할 heading depth이다.

surya_heading_primary_score
  페이지 내 대표 heading을 고르기 위한 점수이다.
```

`reading_order_source` 의미:

```text
surya_match
  하나의 Surya block이 하나의 XML object와 매칭된 경우이다.

surya_region
  하나의 Surya block이 여러 XML object를 포함하는 region으로 처리된 경우이다.

xml_append
  Surya와 매칭되지 못해 XML spatial fallback으로 추가된 객체이다.
  main_converter 출력에서는 이 객체 앞에 [unmatched] marker를 붙인다.
```

## 05_structure_ready/

생성 단계:

```text
04_normalized.json -> reordered slide XML + sidecar JSON
```

`main_converter --reading-order surya`가 직접 읽는 최종 산출물이다.

파일:

```text
manifest.json
slide1.reordered.xml
slide1.structure_analysis.json
slide2.reordered.xml
slide2.structure_analysis.json
...
```

`slideN.reordered.xml`:

```text
PPTX slide XML의 spTree child 순서를 Surya reading order 기준으로 재정렬한 XML이다.
main_converter는 이 파일을 순회하여 Markdown을 생성한다.
```

`slideN.structure_analysis.json`:

```text
해당 slide의 sidecar JSON이다.
heading hint, reading_order_source, input_xml, layout block 정보 등을 담는다.
main_converter는 이 파일을 읽어 heading hint와 unmatched marker를 적용한다.
```

`manifest.json`:

```text
생성된 reordered XML과 sidecar JSON 목록을 기록한다.
```

