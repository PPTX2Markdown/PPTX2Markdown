# PresentationDocument 1.0

[English](output_schema.md) | 한국어

`PresentationDocument`는 표준 출력 계약입니다. Markdown과 JSON은 동일한 순서
문서를 서로 다르게 렌더링한 결과이며 출력 형식에 따라 파싱 경로가 달라지지
않습니다.

기계 판독용 스키마는 Python 패키지에
[`presentation_document.schema.json`](../src/pptx2markdown/presentation_document.schema.json)으로 포함됩니다.

## 최상위 구조

- `schema_version`: 항상 `"1.0"`
- `source.name`: 입력 파일의 basename만 기록하며 절대경로나 작업공간 의존 경로는 제외
- `source.format`: `"pptx"` 또는 `"ppt"`
- `slides`: 페이지 번호가 중복 없이 오름차순으로 정렬된 슬라이드
- `slides[].hidden`: 일반 슬라이드 쇼에서 숨겨진 슬라이드인지 표시. 숨김
  슬라이드도 문서에 남고 Markdown에서는 페이지 marker 바로 뒤에
  `<!-- hidden: true -->`를 렌더링
- `slides[].notes`: 발표자 노트 본문 placeholder에서 추출한 선택 필드. 화면에
  보이는 `blocks`와 분리하며 Markdown에서는 슬라이드 콘텐츠 뒤
  `[Speaker_Notes]` marker 아래에 렌더링. 노트 페이지 header, footer, 날짜,
  슬라이드 번호는 제외

PowerPoint 검토 댓글은 의도적으로 제외합니다. 댓글은 공동 작업 metadata이며
슬라이드 쇼에 표시되지 않으므로 슬라이드 콘텐츠나 발표자 노트와 섞지 않습니다.

## 콘텐츠 블록

`blocks`의 순서가 읽기 순서입니다. `kind`는 `text`, `heading`, `list`, `math`,
`image`, `chart`, `smartart`, `table`, `attachment`, `unsupported` 중 하나입니다.

`attachment`는 내장 패키지 또는 미디어 파일을 출력의 `attachments/` 디렉터리에
보존합니다. 콘텐츠는 `[attachment: report.pdf](attachments/report.pdf)` 같은
Markdown 링크입니다. 가능하면 OLE Packager, Package, CONTENTS stream을
압축 해제하고, 읽을 수 없는 container는 버리지 않고 `.bin`으로 보존합니다.

- `content`: 비어 있지 않은 Markdown 호환 콘텐츠
- `shape_id`: 존재하는 경우 원본 도형 식별자
- `heading_level`: `heading`에는 필수, 다른 kind에는 금지, 1~6 범위
- `bbox`: 선택적 도형 좌표. `x`, `y`, `width`, `height`, 고정 단위 `"emu"`로 구성.
  width와 height는 양수이며 x와 y는 음수일 수 있음
- `source_part`: `slide`, `layout`, `master` 중 하나

모든 단계에서 알 수 없는 필드를 거부합니다. 호환되지 않는 변경은 새로운
`schema_version`을 사용해야 하며 1.0 필드의 의미를 조용히 바꾸지 않습니다.

## Markdown 렌더링 계약

Markdown은 같은 `PresentationDocument`의 결정론적 렌더링이며 두 번째 파서의
출력이 아닙니다.

- 슬라이드가 없는 문서는 0 byte 파일로 렌더링
- 각 슬라이드는 `[Page_<page>]`로 시작
- 숨김 슬라이드는 페이지 marker 뒤에 `<!-- hidden: true -->` 추가
- 블록은 `slides[].blocks` 순서대로 출력. `heading`만 Markdown 헤딩 문법
  (`#`~`######`)을 자동으로 추가하고 다른 블록은 완성된 Markdown 호환 콘텐츠를
  이미 포함
- 발표자 노트는 화면 콘텐츠 뒤 `[Speaker_Notes]` 아래에 출력
- 미디어와 첨부파일 링크는 문서별 출력 디렉터리 기준 상대경로

이 섹션들 사이의 공백도 골든 계약에 포함하며 JSON 및 Markdown 스냅샷에서 바이트
단위로 검사합니다.

## 검증 및 버전 정책

패키지에 포함된 JSON Schema가 직렬화 구조를 검증합니다. Runtime Pydantic
validator는 JSON Schema로 간결하게 표현하기 어려운 의미 불변조건도 검사합니다.
슬라이드 페이지는 중복 없이 오름차순이어야 하고, heading에는 level이 필요하며,
heading이 아닌 블록에는 heading level을 둘 수 없습니다.

버전 `1.0`은 현재 고정된 개발 계약입니다. 아직 출시 전이므로 이전 내부 구조를
위한 호환 adapter는 유지하지 않습니다. 의도적인 비호환 변경은 같은 변경 안에서
`schema_version`, 이 문서, 저장소 JSON Schema, 골든 스냅샷을 함께 갱신해야
합니다.

## 결정론

문서에는 timestamp, work 디렉터리, cache 경로, 절대 입력 경로를 포함하지
않습니다. 동일한 입력과 파서 옵션의 JSON 출력은 바이트 단위로 같아야 합니다.
