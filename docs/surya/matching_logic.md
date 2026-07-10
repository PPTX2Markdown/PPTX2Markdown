# Surya/XML BBox 매칭 로직

이 문서는 Surya layout block과 PPTX XML object를 어떻게 비교하고 매칭하는지 설명한다.

## 기본 아이디어

Surya는 PDF 렌더링 결과를 보고 layout block을 감지한다. PPTX XML은 slide 안의 실제 shape, image, table 같은 객체를 갖고 있다.

두 결과를 연결하려면 서로 다른 좌표계의 bbox를 같은 기준으로 맞춘 뒤 비교해야 한다.

```text
Surya bbox
  PDF 이미지 좌표계
  [x1, y1, x2, y2]

PPTX XML bbox
  EMU 좌표계
  [x, y, w, h]
```

정규화 후에는 둘 다 PPTX EMU 좌표계의 `[x, y, w, h]` 형식으로 비교한다.

## XML bbox 추출

XML 객체는 `slide.xml`의 `p:spTree` 아래 child를 순회하며 추출한다.

대상 tag:

```text
sp
pic
graphicFrame
grpSp
cxnSp
```

bbox는 주로 다음 XML 경로에서 읽는다.

```text
./p:xfrm
./p:spPr/a:xfrm
./p:grpSpPr/a:xfrm
./a:xfrm
.//p:xfrm
.//a:xfrm
```

특히 `graphicFrame`은 table/chart처럼 `p:xfrm`에 bbox를 갖는 경우가 있으므로 `./p:xfrm`을 우선 확인한다.

## Placeholder bbox inheritance

일부 placeholder는 slide.xml에 직접 bbox가 없고 layout/master에 bbox가 정의되어 있다.

이 경우 다음 순서로 bbox를 찾는다.

```text
slide.xml
-> slide relationship
-> slideLayout.xml
-> slideMaster.xml
```

placeholder는 `ph_type`, `ph_idx`를 기준으로 대응되는 layout/master placeholder를 찾는다.

bbox를 상속받은 객체는 다음처럼 기록된다.

```text
bbox_source: "placeholder_inheritance"
bbox_source_part: bbox를 가져온 XML 파일 경로
```

이 처리가 없으면 title placeholder가 bbox 없는 객체로 남아 Surya heading block과 매칭되지 않을 수 있다.

## 장식 요소 처리

본문 Markdown으로 출력하지 않아도 되는 객체는 decorative로 분류한다.

현재 decorative 기준:

```text
1. 텍스트가 없는 cxnSp connector line
2. bbox width 또는 height가 0 이하인 객체
```

decorative 객체는 matching/fallback 대상에서 제외하고 `decorative_objects`에 기록한다.

이 의미는 객체를 삭제한다는 뜻이 아니라, Markdown 본문으로 읽을 콘텐츠가 아니므로 reading order에 넣지 않는다는 뜻이다.

## Surya role 해석

Surya label은 내부 `effective_role`로 변환된다.

예:

```text
SectionHeader -> heading
Header        -> heading
Title         -> heading
Text          -> text
ListItem      -> list
Table         -> table
Figure        -> figure
PageFooter    -> footer
```

일부 Surya output에는 `top_k`가 포함된다. 현재 구현은 `top_k` 전체를 label 대체값으로 쓰지는 않고, heading 보정에 한정해서 사용한다.

```text
top_k의 SectionHeader/Header/Title 계열 점수 >= 0.25
-> heading 성격의 block으로 보정 가능
```

## Candidate score

각 Surya block은 각 XML object와 후보 점수를 계산한다.

주요 점수 요소:

```text
overlap_min
  두 bbox의 교집합 면적 / 더 작은 bbox 면적

layout_coverage
  교집합 면적 / Surya block 면적

object_coverage
  교집합 면적 / XML object 면적

center_score
  두 bbox 중심점이 가까울수록 높은 점수

compatibility
  Surya role과 XML object 특성의 보정 점수
```

최종 `score`는 geometry overlap과 compatibility를 합산한 값이다.

## Compatibility 보정

순수 bbox overlap만으로는 오매칭이 생길 수 있으므로 role/객체 특성을 함께 본다.

대표 보정:

```text
table_graphicFrame
  Surya role이 table이고 XML tag가 graphicFrame이면 강한 보정

heading_title_placeholder
  Surya role이 heading이고 XML 객체가 title placeholder이면 보정

heading_text
  Surya role이 heading이고 XML 객체가 텍스트 shape이면 보정

font_size
  heading 후보에서 XML font size가 큰 경우 보정

numbered_heading
  heading 후보에서 텍스트가 "1. 제목" 형태인 경우 보정

text_shape
  Surya role이 text/list이고 XML 객체가 textual shape이면 보정
```

## 최종 매칭

기본적으로 matcher는 Surya block을 model_position 순서로 순회한다.

일반 텍스트/heading/list/table block은 가장 높은 점수의 XML object 하나를 선택한다. 이미 다른 Surya block에 소비된 XML object는 다시 선택하지 않는다.

일부 figure/picture 성격의 block은 하나의 Surya region이 여러 XML object를 포함할 수 있도록 region 매칭을 허용한다.

매칭 threshold를 넘지 못한 XML object는 숨기지 않고 `xml_append`로 fallback된다.

```text
surya_match
  Surya block과 XML object가 직접 매칭됨

surya_region
  Surya block region 내부의 XML object로 포함됨

xml_append
  Surya block과 매칭되지 못해 XML spatial fallback으로 추가됨
```

## Unmatched marker

`xml_append`로 들어온 객체는 Surya matching에 실패했다는 의미이다.

`main_converter`는 Markdown 출력 시 해당 객체 앞에 다음 marker를 붙인다.

```text
[unmatched]
```

이는 최종 사용자용 표현이라기보다 현재 Surya matching 품질을 확인하기 위한 표식이다. 나중에 matcher가 안정화되면 옵션화하거나 제거할 수 있다.

