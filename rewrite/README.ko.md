# PPTX2Markdown 새 구현

[English](README.md)

이 디렉터리는 기존 변환기와 파일을 공유하지 않는 독립적인 새 구현입니다.
저장소 루트의 기존 변환기를 import하거나 수정하지 않습니다.

새 구현은 책임을 네 단계로 분리합니다.

1. `parse()`가 OOXML을 경로 독립적인 `PresentationDocument`로 읽습니다.
2. `AssetStore`가 반환된 `ParsedPresentation`의 수명 동안 이미지와 첨부파일
   payload를 소유합니다.
3. renderer가 파일을 쓰지 않고 표준 문서를 Markdown 또는 JSON으로 변환합니다.
4. `save()`가 staging 디렉터리를 거쳐 검증된 이동 가능한 bundle을 저장합니다.

## 요구 사항

- Python 3.12 이상
- Pydantic 2.10 이상
- omml2latex 0.1.0 이상
- chart2md 0.1.0 이상
- smartart2md 0.1.1 이상
- olefile 0.47 이상
- 로컬 `.pptx` 입력

파싱 경로에는 OCR, VLM, API key, cloud inference, PowerPoint 자동화,
LibreOffice 의존성이 없습니다.

## Python API

```python
import pptx2markdown

with pptx2markdown.parse("deck.pptx") as presentation:
    for slide in presentation.slides:
        print(slide.page, slide.blocks)

    result = presentation.save("output/deck")
    print(result.document_path)
```

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

규칙 동작: `exclude`, `merged_cells`, `header_rows`, `caption_rows`(표), `heading_level`(제목 수준
변경 또는 텍스트를 제목으로), `as_text`(제목·목록), `description`(이미지 대체
텍스트), `notes`(슬라이드 selector), 그리고 검토용으로만 남는 `reason`. 뒤의
규칙이 앞의 규칙보다 우선하며, 모든 selector는 파싱된 그대로의 문서를 봅니다.
`recipe_json_schema()`(또는 `pptx2markdown --recipe-schema`)는 입력 폼이나 모델의
구조화 출력에 쓸 JSON Schema를 돌려줍니다.

### 검토(review)

변환기는 항상 스스로 결정합니다. `review()`는 그중 다시 볼 만한 결정을 골라,
각 대안을 바로 쓸 수 있는 레시피 규칙으로 함께 내놓습니다. 병합된 그룹이 있는
머리글 아래 행이 라벨처럼 보이는 경우(머리글 두 줄?), 첫 행이 모두 숫자인
경우(머리글 없음?), 병합된 첫 행을 표 제목으로 본 경우, 설명이 없는 이미지,
목록이나 표를 이끄는 짧은 굵은 줄(소제목?)이 그 대상입니다. 사람이나 UI, 언어
모델은 선택지만 고르면 되고, 고른 결과가 곧 레시피가 됩니다.

```python
with pptx2markdown.parse("deck.pptx") as presentation:
    for item in pptx2markdown.review(presentation.document):
        print(item.target, item.question, [option.label for option in item.options])
```

`pptx2markdown deck.pptx --review`는 번들을 쓰지 않고 같은 내용을 JSON으로
출력합니다.

### 에이전트(선택)

`pip install "pptx2markdown[agent]"`를 설치하면 레시피와 검토 위에 언어 모델
계층이 추가됩니다. 모델은 파서나 본문을 직접 고치지 않습니다. 블록 id가 표시된
덱을 읽고 레시피를 쓰며, 그 레시피는 검증을 거치고 모든 규칙이 무언가를
고르는지 확인한 뒤(문제가 있으면 한 번 되돌려 고치게 함) 손으로 쓴 레시피와
똑같이 적용됩니다. Anthropic SDK로 Claude(기본 `claude-opus-5-5`, API의 거절
시 fallback 사용)를 호출하며, 자격 증명은 `ANTHROPIC_API_KEY`나 `ant auth login`을
씁니다. 파싱 경로에서는 모델을 호출하지 않습니다.

```python
from pptx2markdown import agent, merge_recipes

with pptx2markdown.parse("deck.pptx") as presentation:
    draft = agent.draft_recipe(
        presentation.document,
        "검색 인덱스에 넣을 거야. 20번 슬라이드 뒤 부록은 빼고, 그룹으로 묶인 표 "
        "머리글은 두 줄로 봐줘.",
    )
    print(draft.summary, draft.unmatched)
    answers = agent.answer_review(presentation.document)  # 검토 선택지 고르기
    images = agent.describe_images(presentation, language="Korean")  # 이미지 설명
    presentation.apply(merge_recipes(draft.recipe, answers, images))
```

`pptx2markdown deck.pptx --ask "부록은 빼줘"`는 작성된 레시피를 JSON으로 출력하므로,
검토한 뒤 `--recipe`로 적용하면 됩니다.

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
pptx2markdown deck.pptx --output-format json
pptx2markdown deck.pptx --merged-cells blank
pptx2markdown deck.pptx --recipe recipe.json
pptx2markdown --recipe-schema
pptx2markdown deck.pptx --review
pptx2markdown deck.pptx --ask "부록은 빼줘" > recipe.json
```

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
- 좌표만 사용하는 읽기 순서
- title/subtitle placeholder, 절 번호, 글꼴 크기에 따른 제목 판정
- slide layout/master에서 상속한 placeholder 종류, 위치, 글꼴 크기, 글머리 기호
  (바닥글 placeholder는 제외)
- slide layout/master에 그려진 텍스트 shape
- 텍스트, hyperlink, 목록. 한 shape 안에서 이어지는 일반 문단은 텍스트 블록
  하나(문단마다 한 줄)로, 이어지는 글머리 기호 문단은 목록 하나로 묶습니다.
  같은 shape 안의 텍스트·목록·수식은 각자 종류에 맞는 블록으로 둡니다.
  직접 입력한 `▶`로 시작하는 문단은 목록 항목이 되고, 항목 바로 뒤에 공백으로
  들여 쓴 문단은 그 항목에 이어 붙습니다.
- 굵게·기울임 run은 `**굵게**`/`*기울임*`으로 씁니다. 텍스트·제목·목록의 `text`는
  이스케이프된 inline Markdown이라, 직접 입력한 `# x`, `- x`, `<tag>`, `[x](y)`가
  제목·목록·HTML·링크로 바뀌지 않고 글자 그대로 보입니다.
- 수식(OMML)을 omml2latex로 LaTeX 변환
- 내장 이미지
- 내장 audio/video payload
- OLE, 3D model, 오디오·비디오 첨부. 라벨은 파일 이름이고, OLE 복합 파일은 그 안의
  실제 파일(Packager의 원본 파일이나 프로그램의 문서 stream)로 풀어서 저장하며
  ProgID는 `program`에 남깁니다.
- 네이티브 표: JSON에는 렌더링된 `markdown`과 함께 전체 격자(`rows`)를 남깁니다.
  병합 시작 칸에는 텍스트와 `row_span`/`col_span`이, 병합으로 덮인 칸에는
  `merged_from`이 들어갑니다. 첫 행을 머리글로 쓰고 빈 머리글은 `col_N`으로
  채웁니다. 덮인 칸에 무엇을 쓸지는 `merged_cells`로 정합니다. 기본값 `fill`은
  병합된 값을 모든 칸에 반복해 각 행이 혼자서도 뜻이 통하게 하고(검색·RAG에
  유리), `header`는 legacy처럼 머리글 행에서만 반복하며, `blank`는 반복하지
  않습니다. 표 전체 폭으로 병합된 첫 행은 표 제목으로 보고 표 위에 한 번만
  쓰며(`caption_rows`), 본문에서 전체 폭으로 병합된 행은 구역 이름이라 열마다
  반복하지 않습니다. 셀 안의 글머리 기호·번호·링크를 유지하고, 셀 위에 놓인
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

구형 `.ppt`, layout/master에 그려진 그림과 표, OCR, 이미지 설명,
vector 이미지 rasterization은 첫 독립 구현 범위에 포함하지 않았습니다.

## 개발

저장소 루트에서 실행합니다.

```bash
PYTHONPATH=rewrite/src python -m unittest discover -s rewrite/tests -t rewrite -v
ruff check rewrite
ruff format --check rewrite
```

소유권, 수명주기, 보안 설계는 [ARCHITECTURE.ko.md](ARCHITECTURE.ko.md)를
참고하세요.

