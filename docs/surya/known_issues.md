# Surya Pipeline 알려진 문제와 개선 과제

이 문서는 현재 Surya pipeline에서 남아있는 문제와 앞으로 개선해야 할 방향을 정리한다.

## 1. Surya Block과 XML Object의 Grouping 차이

가장 중요한 문제는 Surya layout block과 PPTX XML object가 항상 1:1로 대응되지 않는다는 점이다.

예를 들어 `Heading_Test` Page 2에서 다음 구조가 있었다.

```text
XML shape 4
  text: 2. Text Box Heading - but font size look like Heading
  역할: heading처럼 보이는 텍스트 박스

XML shape 5
  text: If this text box detected as a Heading = # then implementation ok ...
  역할: shape 4 아래의 본문 텍스트 박스
```

하지만 Surya는 shape 4와 shape 5를 완전히 별도 block으로 분리하지 않고, 하나의 큰 block이 두 객체를 함께 덮는 형태로 감지할 수 있다.

예:

```text
Surya block p2_b2
  label: ListItem
  top_k.SectionHeader: 0.997
  bbox: XML shape 4와 XML shape 5를 함께 덮음

candidate scores:
  p2_b2 -> shape 4: 0.971
  p2_b2 -> shape 5: 0.737
```

현재 matcher는 일반 block에 대해 하나의 Surya block을 하나의 XML object에 매칭한다. 따라서 p2_b2는 점수가 가장 높은 shape 4에 배정되고 shape 5는 소비되지 않는다.

이후 shape 5와 매칭될 수 있는 작은 Surya block이 있더라도 점수가 threshold를 넘지 못하면 shape 5는 `xml_append`가 된다.

```text
Surya block p2_b3 -> XML shape 5
  score: 0.282
  result: threshold 미달로 match 실패

최종 결과:
  shape 5 -> xml_append
  Markdown 출력 시 [unmatched] 표시
```

이 문제는 XML bbox가 비정상적으로 큰 것이 아니라, Surya layout granularity와 PPTX XML object granularity가 다른 데서 생긴다.

개선 방향:

```text
1. 하나의 Surya block이 여러 XML object를 포함할 수 있도록 region/member matching을 확장한다.
2. heading block 내부에 포함된 body text를 surya_region_member 등으로 함께 소비한다.
3. 이미 높은 candidate score를 받은 XML object가 block 내부 member로 보이면 xml_append 대신 region member로 분류한다.
4. Markdown에서는 heading과 body text를 순서대로 출력하되 unmatched marker를 붙이지 않는다.
```

## 2. 사람이 보기에는 Heading이지만 Surya가 Heading 신호를 주지 않는 경우

`--headings surya`는 Surya layout 결과 기반 heading만 허용한다.

따라서 사람이 보기에는 heading이어도 Surya `label`과 `top_k`가 heading 신호를 주지 않으면 heading으로 만들지 않는다.

예상되는 원인:

```text
1. Layout model이 해당 텍스트를 ListItem/Text/Code 등으로 분류함
2. top_k에도 SectionHeader 계열 점수가 충분히 높지 않음
3. bbox match는 성공했지만 heading source가 surya_label로 기록되지 않음
```

현재 원칙상 이 케이스를 placeholder/font fallback으로 보정하지 않는다. 그렇게 하면 `--headings surya`가 Surya-only 모드라는 의미를 잃기 때문이다.

개선 방향:

```text
1. auto 모드에서는 기존 fallback을 계속 허용한다.
2. surya 모드에서는 모델 결과를 우선 존중한다.
3. 필요하다면 별도 옵션으로 "surya+fallback" 같은 중간 모드를 추가한다.
4. Surya 모델 결과 자체를 검토할 수 있도록 00~03 JSON을 유지한다.
```

## 3. top_k와 label이 서로 어긋나는 경우

Surya raw result에서 `label`은 Text인데 `top_k.SectionHeader`가 매우 높은 경우가 있었다.

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

현재 해석:

```text
label
  Surya가 최종적으로 선택한 label이다.

top_k
  후보 label들의 점수로 보고 heading 보정에 사용한다.
```

현재 구현은 `top_k`의 가장 높은 label로 전체 label을 바꾸지는 않는다. 대신 heading 판정에 한정하여 SectionHeader/Header/Title 계열 점수가 threshold 이상이면 heading 후보로 본다.

개선 방향:

```text
1. Surya 버전별 raw output schema를 확인한다.
2. top_k가 실제 logits/probabilities인지, 후처리 후보인지 계속 검증한다.
3. heading 외 role에도 top_k를 적용할지 여부는 별도 결정한다.
```

## 4. XML bbox와 실제 렌더링 bbox의 차이

PPTX XML bbox는 object box이고, Surya bbox는 실제 렌더링된 시각 block이다.

따라서 두 bbox는 완전히 같지 않을 수 있다.

차이가 생기는 이유:

```text
1. text box 내부 margin
2. line spacing
3. text wrapping
4. placeholder inheritance
5. group shape transform
6. PDF 렌더링 과정에서의 좌표 변환 오차
```

다만 최근 분석한 `Heading_Test` Page 2의 unmatched 문제는 단순히 XML bbox가 너무 커서 생긴 문제가 아니었다. 그 케이스는 Surya가 인접한 XML 객체들을 하나의 block으로 묶은 grouping 차이가 더 핵심 원인이었다.

개선 방향:

```text
1. candidate score에 center_score뿐 아니라 line-level/containment 신호를 추가한다.
2. text content similarity를 보조 신호로 도입한다.
3. group shape와 nested transform 처리 정확도를 높인다.
4. bbox 차이가 큰 케이스를 known issue fixture로 보존한다.
```

## 5. `xml_append`가 많은 경우의 확인 순서

unmatched가 많이 생기면 다음 순서로 확인한다.

```text
1. 00_raw_surya_result.json
   Surya가 해당 영역을 block으로 감지했는지 확인한다.

2. 01_surya_bboxes.json
   bbox_emu 변환이 slide 좌표계와 맞는지 확인한다.

3. 02_xml_bboxes.json
   XML object bbox가 missing인지, placeholder_inheritance인지 확인한다.

4. 03_match_candidates.json
   후보 score가 threshold를 넘지 못했는지 확인한다.

5. 04_normalized.json
   최종 reading_order_source가 surya_match/surya_region/xml_append 중 무엇인지 확인한다.
```

대부분의 문제는 다음 셋 중 하나이다.

```text
1. Surya가 block을 잘못 감지함
2. XML bbox 추출이 부정확함
3. Surya block과 XML object의 granularity가 다름
```

