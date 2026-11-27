# PresentationDocument 2.0

[English](output_schema.md) | 한국어

`--output-format json`과 `ParsedPresentation.render_json()`은 `PresentationDocument`
하나를 씁니다. Markdown도 같은 모델에서 렌더링합니다. JSON Schema는 패키지에서
얻을 수 있습니다.

```python
from pptx2markdown import PresentationDocument

schema = PresentationDocument.model_json_schema()
```

## 최상위

| 필드 | 의미 |
| --- | --- |
| `schema_version` | `"2.0"`. 필드의 의미가 바뀌거나 사라질 때 올림 |
| `source` | 덱의 파일 이름, 형식, SHA-256 (전체 경로는 넣지 않음) |
| `assets` | 이미지와 첨부: id, 종류, media type, 원래 이름, 크기, SHA-256, 원본 part |
| `slides` | 프레젠테이션 순서대로 놓인 슬라이드 |
| `warnings` | 변환하지 못한 콘텐츠에 대한 구조화된 경고 |

## 슬라이드

| 필드 | 의미 |
| --- | --- |
| `page` | 프레젠테이션 슬라이드 목록에서의 위치(1부터) |
| `hidden` | PowerPoint에서 숨긴 슬라이드인지 |
| `section` | 슬라이드가 속한 PowerPoint 구역(있는 경우) |
| `blocks` | 읽기 순서대로 놓인 콘텐츠 블록 |
| `notes` | 발표자 노트 |
| `reading_order` | 여러 방식으로 읽힐 수 있는 배치일 때 그 사유와 후보 순서 |

모든 블록에는 `id`(`s<page>.b<n>`)와 `kind`가 있고, 도형이 있으면 EMU 단위 `bbox`와
`shape_id`도 있습니다. 종류: `heading`(`level`, `source`), `text`, `list`(수준과
번호 여부가 있는 항목), `table`(렌더링된 `markdown`과 병합 범위를 포함한 전체 `rows`
격자), `image`와 `attachment`(`asset_id`), `chart`, `smartart`, `math`(LaTeX),
`unsupported`(변환하지 못한 것. 조용히 버리지 않음).

문서 옆에 쓰이는 번들에는 원본 해시와 저장한 모든 에셋의 경로, 크기, SHA-256을 담은
`bundle-manifest.json`이 있습니다.
