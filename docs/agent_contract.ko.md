# 변환기–에이전트 계약

[English](agent_contract.md)

상태: `rewrite/`의 차세대 엔진에 적용 중 (review 스키마 버전 1).

## 역할

pptx2markdown은 확신이 없는 곳을 스스로 알려 주는 결정적 PPTX 변환기입니다.
파싱 중에는 모델을 부르지 않습니다. 판단하는 쪽(사람, UI, 로컬 모델, Claude,
다른 곳에서 만든 멀티포맷 에이전트)은 그 위에 놓이고, 세 가지 순수 데이터
인터페이스로만 대화합니다.

| 인터페이스 | 방향 | 내용 |
| --- | --- | --- |
| 문서 | 변환기 → 에이전트 | 블록 id가 붙은 Markdown(`<!-- s3.b2 table -->`)과 정규 JSON |
| 결정 지점(`review()`) | 변환기 → 에이전트 | 변환기가 확신 없이 고른 곳과, 바로 쓸 수 있는 대안 |
| 레시피 | 에이전트 → 변환기 | 답을 규칙으로 표현한 것. 결정적으로 재현되고 원본은 건드리지 않음 |

`audit()`는 네 번째 읽기 전용 신호로, 슬라이드에 입력된 텍스트 중 문서에 없는
것을 알려 줍니다. 레시피로 되살릴 수 없으므로 에이전트는 답하지 말고 보고합니다.

이렇게 나누면 양쪽을 각각 교체할 수 있습니다. 변환기는 에이전트 코드 변경 없이
새 결정 주제를 늘릴 수 있고, 모든 변환기가 같은 형태로 결정을 보고하면 하나의
에이전트가 여러 포맷을 다룰 수 있습니다.

## 결정 지점

```json
{
  "id": "table-header-rows@s3.b2",
  "topic": "table-header-rows",
  "target": "s3.b2",
  "page": 3,
  "question": "The header has merged groups; is the next row part of the header?",
  "reasons": ["grouped-header"],
  "look": "text",
  "impact": "high",
  "options": [
    {"id": "keep", "label": "One header row", "rule": null},
    {"id": "two-rows", "label": "Two header rows",
     "rule": {"target": "s3.b2", "header_rows": 2, "reason": "grouped two-row header"}}
  ],
  "suggested": "two-rows",
  "free_text": null,
  "asset_id": null,
  "evidence": {"header": ["Item", "2024", "2024"], "next_row": ["Item", "H1", "H2"], "header_spans_rows": true}
}
```

| 필드 | 의미 |
| --- | --- |
| `id` | `<topic>@<target>`. 같은 덱·같은 변환기 버전에서 고정 |
| `topic` | 결정의 종류(아래 표) |
| `target` | 옵션 규칙이 적용될 레시피 선택자 |
| `page` | 확인할 슬라이드 |
| `reasons` | 왜 묻는지에 대한 기계 판독용 코드 |
| `look` | 판단에 필요한 것: `text`(슬라이드 Markdown으로 충분), `slide`(렌더링), `image`(그림 자체, `asset_id`) |
| `impact` | `high`: 틀리면 검색용 내용이 망가짐, `low`: 서식 |
| `options` | 고정 id를 가진 답. 첫 번째는 항상 `keep`(규칙 `null`, 변환기 결정 유지) |
| `suggested` | 판단 근거가 없을 때 택할 옵션 |
| `free_text` | 설정된 경우, 자유 텍스트 답이 이 규칙 필드를 채움 |
| `evidence` | 변환기가 사용한 주제별 근거 |

항목을 감싸는 보고서(`--review`, MCP `review_presentation`)에는 `schema_version`과
전체 항목에 대한 주제·look·impact별 `counts`가 있습니다. MCP 도구는 `topics`, `looks`,
`impact`, `pages`로 목록만 좁히고 counts는 그대로 두며, `convert_presentation`도
`pages`를 받으므로 큰 덱에서도 에이전트는 지금 판단할 것만 가져옵니다.

`look`과 `impact`로 에이전트가 비용을 계획할 수 있습니다. `text` 항목은 Markdown만
보고 한 번에 답하고, `slide` 항목의 페이지만 렌더링하고, `image` 항목의 그림만
가져오며, 시간·비용이 부족하면 `low` 항목은 건너뜁니다.

### 주제

| 주제 | 대상 | 사유 | look | impact | 옵션 |
| --- | --- | --- | --- | --- | --- |
| `reading-order` | `slide[N]` | `two-way-cut`, `no-cut`, `overlapping-boxes`, `fragile-gap`, `container-across-content` | slide | high | `keep`, `rows-first`, `columns-allowing-overlap`, `as-stored-in-the-file` (다른 순서만) |
| `table-header-rows` | 블록 | `grouped-header`, `numeric-first-row` | text | high | `keep`, `two-rows` / `no-header` |
| `image-description` | 블록 | `no-alt-text`, `default-picture-name` | image | high | `keep`, `decorative`; 자유 텍스트 → `description` |
| `table-caption` | 블록 | `merged-first-row` | text | low | `keep`, `header` |
| `heading` | 블록 | `font-size-only` | text | low | `keep`, `text` |
| `subheading` | 블록 | `bold-line-introduces-block` | text | low | `keep`, `heading` |

읽기 순서 evidence에는 옵션 id별로 그 순서대로 놓인 각 블록의 첫 단어들이 들어
있습니다.

## 답하기

`recipe_from_answers(items, {항목 id: 옵션 id 또는 자유 텍스트})`가 레시피를
만들고, MCP 도구 `answers_to_recipe`도 같은 일을 합니다. 답하지 않은 항목은
변환기 결정을 유지하거나, `unanswered="suggested"`이면 `suggested`를 따릅니다.
모르는 항목 id나 옵션은 모든 문제를 나열한 하나의 오류로 알려 줍니다.

모델 없는 판단 루프가 요청을 만들고 답을 모읍니다. 어떤 함수든 넘길 수 있습니다.

- `describe_images_with(presentation, describe)`: 그림 바이트, alt 텍스트,
  슬라이드 Markdown을 넘기고 설명 또는 "장식"을 받습니다.
- `choose_reading_orders_with(presentation, judge)`: 플래그된 슬라이드마다 한 쪽
  PDF(LibreOffice), 사유, 후보를 넘기고 후보 이름을 받습니다.

`pptx2markdown.snapshot.slide_pdf` / `slide_png`는 다른 판단자를 위해 슬라이드 한 장을
렌더링하고, MCP 도구 `presentation_slides`는 `look: "slide"` 항목의 슬라이드를 호스트
모델에게 보여 주므로 별도 API 키 없이도 답할 수 있습니다.

`pptx2markdown.agent`(선택적 `[agent]` extra)는 이 루프에 Claude를 붙인 것입니다.
모델 SDK를 import하는 유일한 모듈이며, 코어와 MCP 서버는 이를 불러오지 않습니다
(테스트로 강제).

## 버전

`--review` 출력과 MCP 도구 결과의 `schema_version`은 필드의 의미가 바뀌거나
사라질 때 올립니다. 새 주제·사유·옵션·필드는 버전을 유지하므로, 에이전트는 모르는
것을 무시해야 하며 언제든 `suggested`로 물러설 수 있습니다. 레시피는 자체 JSON
Schema(`--recipe-schema`)로 버전이 관리됩니다.
