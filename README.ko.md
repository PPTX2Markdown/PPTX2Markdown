# pptx2markdown

[![Quality](https://github.com/PPTX2Markdown/PPTX2Markdown/actions/workflows/quality.yml/badge.svg)](https://github.com/PPTX2Markdown/PPTX2Markdown/actions/workflows/quality.yml)
[![PyPI](https://img.shields.io/pypi/v/pptx2markdown)](https://pypi.org/project/pptx2markdown/)
[![Python](https://img.shields.io/pypi/pyversions/pptx2markdown)](https://pypi.org/project/pptx2markdown/)
[![License: Apache 2.0](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/LICENSE)

[English](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/README.md) | 한국어

`pptx2markdown`은 PowerPoint 프레젠테이션을 결정적인 Markdown, JSON, 검색용 청크로
변환합니다. 네이티브 OOXML 구조와 도형 좌표를 직접 읽고, 모델을 호출하지 않으며,
확신이 없는 곳을 스스로 알려 줍니다. 다시 볼 만한 결정에는 바로 쓸 수 있는 대안이
함께 나오므로, 사람·UI·모델이 레시피로 답하면 그 결과가 결정적으로 재현됩니다.

## 왜 pptx2markdown인가요?

- **슬라이드 구조를 따르는 읽기 순서** — 카드, 라벨 상자, 그룹 도형은 한 단위로
  읽고 배경은 무시합니다. 여러 방식으로 읽힐 수 있는 슬라이드는 후보 순서와 함께
  표시합니다.
- **의미를 지키는 표** — 병합 범위를 포함한 전체 셀 격자, 모든 행이 단독으로
  의미를 갖도록 반복되는 병합 값, 요청 시 두 줄 그룹 머리글.
- **조용한 추측 대신 결정 지점** — `review()`가 확신 없는 선택을 고정 id, 사유,
  봐야 할 것, 중요도와 함께 나열하고, 답은 레시피가 됩니다.
- **조용한 누락 없음** — `audit()`가 슬라이드 XML을 다시 읽어 출력에 없는 입력
  텍스트를 보고합니다.
- **검색 파이프라인 친화적** — 출처 메타데이터가 붙은 슬라이드 단위 청크, bytes·
  stream 입력, 에이전트용 MCP 서버.
- **풍부한 네이티브 콘텐츠** — 텍스트, 제목, 목록, 표, 차트, SmartArt, 수식,
  이미지, 발표자 노트, 구역, 내장 첨부파일.

[기능 벤치마크](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/benchmarks/README.ko.md)에서
각 변환기가 덱에서 지켜 낸 것:

| | pptx2markdown | anydoc 0.2.5 | MarkItDown 0.1.9 |
| --- | :---: | :---: | :---: |
| 유지한 기능 (20개 중) | **20** | 13 | 8 |

0.1.x에서 올라오시나요? 출력, CLI 옵션, Python API가 바뀌었습니다.
[마이그레이션 가이드](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/docs/MIGRATION.ko.md)를 참고하세요.

## 설치

### AI 에이전트에게 설치 맡기기

로컬 코딩 에이전트가 환경을 확인하고, 격리된 설치 방식을 고르고, 필요하면
LibreOffice를 설정한 뒤 CLI를 검증할 수 있습니다. 아래 프롬프트를 주고 모든 권한
요청을 확인하세요.

```text
https://raw.githubusercontent.com/PPTX2Markdown/PPTX2Markdown/main/.github/agent-install.ko.md 를 읽고 이 컴퓨터에 pptx2markdown을 설치해 줘. 설치에 필요한 명령은 설명한 뒤 실행해도 돼. 관리자 권한, sudo, 비밀번호, 시스템 패키지 관리자, 셸 프로필, PATH, LibreOffice 변경 전에는 먼저 물어봐.
```

전체 [에이전트 설치 절차](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/.github/agent-install.ko.md)를 참고하세요.

### 직접 설치

```bash
uv tool install pptx2markdown
# 또는
pipx install pptx2markdown
# 또는 활성화된 가상 환경 안에서
python -m pip install pptx2markdown
```

추가 기능: `pptx2markdown[mcp]`(MCP 서버), `pptx2markdown[images]`(PNG로 그린
EMF/WMF 그림의 여백 정리).

LibreOffice는 선택 사항입니다. `.ppt` 입력, EMF/WMF 그림의 PNG 변환, 읽기 순서가
애매한 슬라이드의 렌더링처럼 요청한 경우에만 씁니다.

```bash
# macOS
brew install --cask libreoffice

# Ubuntu / Debian
sudo apt-get install libreoffice-impress
```

## 빠른 시작

```bash
# 파일 하나 -> ./output/deck/deck.md
pptx2markdown deck.pptx

# 여러 파일, 출력 위치 지정, CPU마다 작업 프로세스 하나
pptx2markdown decks/*.pptx --output-dir converted --jobs 0

# 정규 JSON, 또는 슬라이드 청크마다 JSON 한 줄
pptx2markdown deck.pptx --output-format json
pptx2markdown deck.pptx --output-format chunks

# 변환기가 확신하지 못한 곳, 보존하지 못한 텍스트
pptx2markdown deck.pptx --review
pptx2markdown deck.pptx --audit
```

## 요구 사항

- Python 3.12 이상
- Pydantic 2.10 이상
- omml2latex 0.1.0 이상
- chart2md 0.1.0 이상
- smartart2md 0.1.1 이상
- olefile 0.47 이상
- 로컬 `.pptx` 입력. 레거시 `.ppt`는 요청할 때만 LibreOffice로 변환

파싱 경로에는 OCR, VLM, API key, cloud inference, PowerPoint 자동화가 없습니다.
`.ppt` 파일은 `ppt_converter="libreoffice"`(`--ppt-converter libreoffice`)를 줄 때만
변환하며, LibreOffice를 headless로 실행하고 작업 후 지워지는 임시 디렉터리만
사용합니다(`SOFFICE_PATH`로 실행 파일 지정). `vector_images="png"`
(`--vector-images png`)도 마찬가지입니다. 브라우저·Markdown 뷰어·비전 모델이 읽지 못하는
EMF/WMF 그림을 LibreOffice 한 번 실행으로 PNG로 그려 같은 asset id로 원본을 대신하며,
`images` extra(`pip install "pptx2markdown[images]"`)가 있으면 LibreOffice가 붙이는 페이지
여백을 잘라냅니다.

## Python API

```python
import pptx2markdown

with pptx2markdown.parse("deck.pptx") as presentation:
    for slide in presentation.slides:
        print(slide.page, slide.blocks)

    result = presentation.save("output/deck")
    print(result.document_path)
```

`parse()`는 파일의 bytes나 바이너리 스트림도 받으므로, 업로드된 파일을 디스크에
쓰지 않고 바로 변환할 수 있습니다(`name`은 원본 이름으로 기록됨).

```python
with pptx2markdown.parse(request.files["deck"].stream, name="deck.pptx") as presentation:
    markdown = presentation.render_markdown(assets="omit")
```

메모리 입력은 `PackageLimits.max_package_bytes`(1 GiB)까지 받습니다.

`parse()`는 영구 출력 파일을 남기지 않습니다. 내장 payload는 관리되는 spooled
store에 보관하며 canonical 문서에는 임시경로가 들어가지 않습니다.

저장하지 않고 Markdown을 렌더링할 때는 asset 정책을 명시해야 합니다.

```python
with pptx2markdown.parse("deck.pptx") as presentation:
    text_only = presentation.render_markdown(assets="omit")

    web_markdown = presentation.render_markdown(
        asset_url=lambda asset_id: f"/assets/{asset_id}",
    )
```

### 검색용 청크

`slide_chunks()`는 슬라이드마다 청크 하나를 돌려주고, 인용에 필요한 metadata(원본,
페이지, 슬라이드·덱 제목, 블록 id, 블록 종류)를 붙입니다. `max_chars`보다 긴
슬라이드는 블록 경계에서 나누며 표 중간에서 자르지 않습니다. 각 청크는 덱과
슬라이드를 밝히는 주석으로 시작하므로, 검색된 청크 하나만 봐도 맥락을 알 수
있습니다.

```python
with pptx2markdown.parse("deck.pptx") as presentation:
    for chunk in pptx2markdown.slide_chunks(presentation.document, max_chars=2000):
        index.add(chunk.text, metadata=chunk.metadata)
```

`--output-format chunks`는 같은 청크를 JSON Lines(`<덱>.chunks.jsonl`, 줄마다
`{"text": ..., "metadata": ...}`)로 쓰며, 이미지 링크는 번들 안을 가리킵니다.
`--max-chunk-chars`로 나누는 크기를 정합니다.

### 레시피

레시피는 파서를 건드리지 않고 파싱된 문서를 조정합니다. 문서 전체 설정과,
블록이나 슬라이드를 골라 무엇을 할지 정하는 규칙으로 이루어집니다. 단순한
데이터라서 덱 옆에 저장하거나, 검토하거나, 팀에서 공유할 수 있고, "부록은 빼고
매출 표는 위 두 줄을 머리글로 봐줘" 같은 요청으로부터 언어 모델이 작성할 수도
있습니다. 적용 결과는 항상 같습니다.

```python
recipe = pptx2markdown.Recipe(
    merged_cells="blank",
    rules=[
        {"target": "slide[12-]", "exclude": True, "reason": "부록"},
        {"target": "slide[3]/table[0]", "header_rows": 2},
        {"target": "s5.b2", "description": "2021-2024 분기별 매출"},
        {"target": "s1.b3", "heading_level": 2},
    ],
)
with pptx2markdown.parse("deck.pptx", recipe=recipe) as presentation:
    ...
```

모든 블록에는 `s<페이지>.b<번호>` id가 있습니다. `render_markdown(block_ids=True)`는
이를 `<!-- s3.b2 table -->` 주석으로 보여 주므로 검토자나 모델이 정확한 블록을
가리킬 수 있고, `explain_recipe()`는 적용 전에 각 규칙이 무엇을 고르는지
알려 줍니다. selector는 다음과 같습니다.

| Selector | 고르는 대상 |
| --- | --- |
| `s3.b2` | id로 지정한 블록 하나 |
| `slide[3]`, `slide[2-5]`, `slide[12-]`, `slide[*]`, `slide[hidden]` | 슬라이드 전체 |
| `table`, `image`, `heading` 등 또는 `*` | 해당 종류의 모든 블록 |
| `table[0]`, `table[-1]` | 슬라이드마다 해당 종류의 첫 번째·마지막 블록 |
| `slide[3]/table[0]` | 3번 슬라이드의 첫 번째 표 |
| `section[Appendix]`, `section[Appendix]/image` | PowerPoint 구역(section)의 슬라이드(또는 블록) |

설정: `merged_cells`, `include_notes`, `include_hidden`, 그리고 `headings`
(`placeholder`는 제목 placeholder에서 나온 제목만 남기며, legacy의
`--headings strict`에 해당합니다).

규칙 동작: `exclude`, `merged_cells`, `header_rows`, `caption_rows`(표), `heading_level`(제목 수준
변경 또는 텍스트를 제목으로), `as_text`(제목·목록), `description`(이미지 대체
텍스트), `notes`와 `order`(슬라이드 selector: 발표자 노트 유지·제거, 블록 id 순서대로
읽기), 그리고 검토용으로만 남는 `reason`. 뒤의
규칙이 앞의 규칙보다 우선하며, 모든 selector는 파싱된 그대로의 문서를 봅니다.
`recipe_json_schema()`(또는 `pptx2markdown --recipe-schema`)는 입력 폼이나 모델의
구조화 출력에 쓸 JSON Schema를 돌려줍니다.

### 검토(review)

변환기는 항상 스스로 결정합니다. `review()`는 그중 다시 볼 만한 결정을 골라,
각 대안을 바로 쓸 수 있는 레시피 규칙으로 함께 내놓습니다. 병합된 그룹이 있는
머리글 아래 행이 라벨처럼 보이는 경우(머리글 두 줄?), 첫 행이 모두 숫자인
경우(머리글 없음?), 병합된 첫 행을 표 제목으로 본 경우, 설명이 없는 이미지,
글꼴 크기만으로 판정한 제목(제목이 맞나?), 목록이나 표를 이끄는 짧은 굵은
줄(소제목?), 그리고 여러 순서로 읽힐 수 있는 배치의 슬라이드가 그 대상입니다.
사람이나 UI, 언어 모델은 선택지만 고르면 되고, 고른 결과가 곧 레시피가 됩니다.

읽기 순서는 슬라이드의 구조를 따릅니다. 그룹과 다른 도형을 감싸는 도형(카드
테두리, 라벨이 붙은 상자, 콜아웃이 올라간 그림)은 한 단위로 읽고, 슬라이드를 덮는
배경은 무시합니다. 기하 정보만으로 순서를 정할 수 없는 곳(행으로도 열로도 읽히는
격자, 틈 없이 나란한 텍스트)은 `reading-order` 검토 항목이 되고, 그 선택지는 후보
순서를 담은 `order` 규칙입니다. 그래서 사람이나 모델은 그런 슬라이드만 보면 됩니다.
`choose_reading_orders_with(presentation, judge)`는 그 슬라이드만 LibreOffice로 로컬에서
한 장짜리 PDF로 렌더링해, 사용자가 준 `judge` 함수에 후보 중 하나를 고르게 합니다.
[설계 문서](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/docs/reading_order_design.ko.md)를 참고하세요.

```python
with pptx2markdown.parse("deck.pptx") as presentation:
    for item in pptx2markdown.review(presentation.document):
        print(item.target, item.question, [option.label for option in item.options])
```

`pptx2markdown deck.pptx --review`는 번들을 쓰지 않고 같은 내용을 JSON으로
출력합니다.

### 이미지 설명(모델 무관)

이미지 설명은 추가 설치 없이 로컬이든 아니든 어떤 모델로도 만들 수 있습니다.
`describe_images_with`는 설명이 필요한 그림마다 슬라이드 Markdown을 맥락으로 붙여
사용자가 준 함수에 넘기고, 답을 레시피로 돌려줍니다. 그 함수가 보내지 않는 한
어떤 데이터도 밖으로 나가지 않습니다(사내망·온프레미스 환경).

```python
from pptx2markdown import ImageDescription, describe_images_with

def describe(request):  # request.data, .media_type, .alt_text, .slide_markdown
    text = my_local_vision_model(request.data, context=request.slide_markdown)
    return ImageDescription(description=text)

with pptx2markdown.parse("deck.pptx") as presentation:
    presentation.apply(describe_images_with(presentation, describe))
```

### 감사(audit)

규칙 기반 변환기는 조용히 실패합니다. 이해하지 못한 도형은 출력에서 그냥 빠집니다.
`audit()`은 파서의 규칙을 거치지 않고 슬라이드 XML을 다시 읽어, 문서에 없는 입력
텍스트를 슬라이드 번호와 도형 id와 함께 모두 보고하고, 문서에서 빠진 슬라이드도
알려 줍니다. 텍스트는 글자와 숫자만 비교하므로 이스케이프, 강조, 목록 기호, 줄바꿈은
손실로 치지 않습니다. 바닥글·날짜·슬라이드 번호 placeholder와 병합으로 가려진 셀은
제외하며, 차트·SmartArt·수식은 감사하지 않습니다.

```python
with pptx2markdown.parse("deck.pptx") as presentation:
    report = presentation.audit()
    print(report.ok, report.coverage, report.losses)
```

`pptx2markdown deck.pptx --audit`는 번들을 쓰지 않고 보고서를 JSON으로 출력하며,
손실이 있으면 종료 코드 1을 돌려줍니다. 레시피가 지운 내용도 손실로 보이므로 CLI와
MCP 도구는 레시피 없이 변환 결과를 감사합니다.

### 모델로 판단하기

패키지 안에서는 어떤 모델도 호출하지 않습니다. 판단이 필요한 곳에서는 그 판단에
필요한 것만 사용자가 넘긴 함수에 전달하고, 답을 레시피 규칙으로 바꿉니다. 검토
항목은 `recipe_from_answers`, 이미지 설명은 `describe_images_with`, 읽기 순서는
`choose_reading_orders_with`로 처리합니다. 함수는 사람이어도, 로컬 모델이어도,
호스팅 모델이어도 됩니다. 아래 MCP 서버를 쓰면 호스트 모델(예: Claude Code)이
별도 API 키 없이 같은 일을 합니다. 여러 포맷을 다루는 에이전트는
[계약](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/docs/agent_contract.ko.md)을 통해 그 위에 올릴 수 있습니다.

### MCP 서버

`pip install "pptx2markdown[mcp]"`를 설치하면 MCP(Model Context Protocol) 서버
`pptx2markdown-mcp`(stdio)가 추가되어, Claude 같은 에이전트가 레시피 작업 흐름을 직접
수행할 수 있습니다. 도구는 `convert_presentation`(블록 id가 표시된 Markdown, 레시피
적용 가능), `review_presentation`, `audit_presentation`(변환에서 빠진 텍스트),
`presentation_images`(설명 없는 그림을 이미지로 돌려주어 호스트 모델이 레시피로 설명을
달 수 있게 함), `check_recipe`(규칙마다 고르는 대상), `presentation_chunks`,
`recipe_schema`입니다. 두 프롬프트가 호스트 모델을 작업 흐름으로 안내합니다.
`prepare_for_retrieval`(감사, 검토 답하기, 그림 설명, 레시피 확인, 청크)과
`apply_request`(자연어 요청을 확인된 레시피로)입니다. 로컬 `.pptx` 파일을 읽기만 하고 아무것도
쓰지 않습니다.

```bash
claude mcp add pptx2markdown -- pptx2markdown-mcp
```

또는 MCP 클라이언트 설정에서:

```json
{"mcpServers": {"pptx2markdown": {"command": "pptx2markdown-mcp"}}}
```

파싱과 저장을 한 번에 수행하려면 `convert()`를 사용합니다.

```python
result = pptx2markdown.convert(
    "deck.pptx",
    output_dir="output",
)
```

## CLI

```bash
pptx2markdown deck.pptx
pptx2markdown first.pptx second.pptx --output-dir converted
pptx2markdown decks/*.pptx --jobs 0   # CPU마다 작업 프로세스 하나
pptx2markdown deck.pptx --output-format json
pptx2markdown deck.pptx --output-format chunks --max-chunk-chars 2000
pptx2markdown deck.pptx --merged-cells blank
pptx2markdown deck.pptx --recipe recipe.json
pptx2markdown --recipe-schema
pptx2markdown deck.pptx --review
pptx2markdown deck.pptx --audit
pptx2markdown old-deck.ppt --ppt-converter libreoffice --vector-images png
```

PowerPoint 잠금 파일(`~$deck.pptx`, 덱을 열어 둔 동안 생김)은 경고만 하고 건너뛰므로
`decks/*.pptx` 같은 글롭이 이 파일 때문에 실패하지 않습니다.

기본 출력:

```text
output/
└── deck/
    ├── deck.md
    ├── media/
    ├── attachments/
    └── bundle-manifest.json
```

Markdown 링크는 `deck.md` 기준 상대경로이므로 `deck/` 디렉터리 전체를 옮겨도
내장 asset 연결이 유지됩니다.

## 지원하는 정적 콘텐츠

- 슬라이드 순서와 숨김 슬라이드 metadata
- 재귀 XY-cut으로 좌표만 사용하는 읽기 순서: 세로 공백으로 나뉜 열은 한 열씩
  읽고, 전체 폭 제목이나 구분선은 그 아래 열보다 먼저 읽으며, 조금씩 어긋나게
  쌓인 상자는 위에서 아래로 읽습니다. 1px까지 겹쳐도 공백으로 보고, 배경 틀과
  바닥글은 열 경계를 막지 않습니다.
- title/subtitle placeholder, 절 번호, 글꼴 크기에 따른 제목 판정. 각 제목은 판정
  근거를 `source`(`placeholder`, `position`, `numbering`, `text`, `font`)로 남기므로
  가장 약한 판정만 골라 검토하거나 뺄 수 있습니다.
- slide layout/master에서 상속한 placeholder 종류, 위치, 글꼴 크기, 글머리 기호
  (바닥글 placeholder는 제외)
- slide layout/master에 그려진 텍스트 shape와 그림(블록에 `source_part`가 남고, 가는
  선·띠는 장식으로 봅니다. 반복되는 로고는 레시피 설정 `template_pictures: false`로
  뺄 수 있습니다)
- 텍스트, hyperlink, 목록. 한 shape 안에서 이어지는 일반 문단은 텍스트 블록
  하나(문단마다 한 줄)로, 이어지는 글머리 기호 문단은 목록 하나로 묶습니다.
  같은 shape 안의 텍스트·목록·수식은 각자 종류에 맞는 블록으로 둡니다.
  직접 입력한 `▶`로 시작하는 문단은 목록 항목이 되고, 항목 바로 뒤에 공백으로
  들여 쓴 문단은 그 항목에 이어 붙습니다.
- 굵게·기울임 run은 텍스트 일부를 강조할 때 `**굵게**`/`*기울임*`으로 씁니다. 표
  셀 전체나 여러 줄짜리 텍스트 상자 전체에 걸린 굵게는 서식이라 빼고, 렌더러가
  표시할 수 없는 강조도 뺍니다. 텍스트·제목·목록의 `text`는
  이스케이프된 inline Markdown이라, 직접 입력한 `# x`, `- x`, `<tag>`, `[x](y)`가
  제목·목록·HTML·링크로 바뀌지 않고 글자 그대로 보입니다.
- 수식(OMML)을 omml2latex로 LaTeX 변환
- 내장 이미지(그림으로 채운 도형 포함)
- 내장 audio/video payload
- OLE, 3D model, 오디오·비디오 첨부. 라벨은 파일 이름이고, OLE 복합 파일은 그 안의
  실제 파일(Packager의 원본 파일이나 프로그램의 문서 stream)로 풀어서 저장하며
  ProgID는 `program`에 남깁니다.
- 네이티브 표: JSON에는 렌더링된 `markdown`과 함께 전체 격자(`rows`)를 남깁니다.
  병합 시작 칸에는 텍스트와 `row_span`/`col_span`이, 병합으로 덮인 칸에는
  `merged_from`이 들어갑니다. 첫 행을 머리글로 쓰고 빈 머리글은 `col_N`으로
  채웁니다. 덮인 칸에 무엇을 쓸지는 `merged_cells`로 정합니다. 기본값 `fill`은
  병합된 값을 모든 칸에 반복해 각 행이 혼자서도 뜻이 통하게 하고(검색·RAG에
  유리), `header`는 머리글 행에서만 반복하며, `blank`는 반복하지
  않습니다. 표 전체 폭으로 병합된 첫 행은 표 제목으로 보고 표 위에 한 번만
  씁니다(`caption_rows`). 셀 안의 글머리 기호·번호·링크를 유지하고, 셀 위에 놓인
  그림이나 셀 배경 그림은 그 셀 안에 넣습니다.
- chart2md를 이용한 차트 변환: 일반 차트는 cache 값으로, Office 2016 chartEx
  차트(깔때기, 폭포 등)는 내장 workbook으로 변환합니다. 내장 workbook도 패키지와
  같은 기준으로 검사하며, 거부되면 cache 값을 사용합니다.
- smartart2md를 이용한 SmartArt 변환: 계층은 들여 쓴 목록, 프로세스·순환은 번호
  단계, 행렬은 표, 그림은 이미지로 씁니다. 텍스트 없는 노드만 있는 다이어그램은
  `[smartart: N unlabeled nodes]`로 개수를 남깁니다.
- 발표자 노트
- 그룹 도형과 Markup Compatibility `AlternateContent`

지원하지 못하거나 관계가 없는 콘텐츠는 구조화 warning 또는
`UnsupportedBlock`으로 남기며, 성공한 것처럼 조용히 버리지 않습니다.

OCR과 layout/master에 그려진 표는 지원하지 않으며, 이미지 설명은
사용자가 넘긴 설명 함수가 만듭니다.

## 개발

```bash
python -m pip install -e ".[mcp]"
python -m unittest discover -v
python scripts/update_goldens.py --check
ruff check src tests scripts benchmarks
ruff format --check src tests scripts benchmarks
```

변환 결과가 의도적으로 바뀌면 golden 스냅샷의 diff를 검토한 뒤
`python scripts/update_goldens.py --accept`로 받아들입니다.

0.1.x에서 옮겨 오는 경우 [마이그레이션 가이드](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/docs/MIGRATION.ko.md)를 참고하세요.

소유권, 수명주기, 보안 설계는 [ARCHITECTURE.ko.md](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/docs/ARCHITECTURE.ko.md)를
참고하세요.

## 라이선스

Copyright 2026 Moonyoung Lee, Inseong Lee, Seung-Woo Jeong

이 프로젝트는 [Apache License 2.0](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/LICENSE)으로 배포됩니다. 서드파티 golden
fixture의 출처 표기는 [`THIRD_PARTY_NOTICES.ko.md`](https://github.com/PPTX2Markdown/PPTX2Markdown/blob/main/tests/fixtures/golden/THIRD_PARTY_NOTICES.ko.md)에 있습니다.
