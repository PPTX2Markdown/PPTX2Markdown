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
- 네이티브 표: 첫 행을 머리글로 쓰고, 여러 열에 걸친 머리글은 각 열에 같은
  텍스트를 반복하며, 빈 머리글은 `col_N`으로 채웁니다. 본문의 병합 셀은 비워
  둡니다. 셀 안의 글머리 기호·번호·링크를 유지하고, 셀 위에 놓인 그림이나 셀
  배경 그림은 그 셀 안에 넣습니다.
- chart2md를 이용한 차트 변환: 일반 차트는 cache 값으로, Office 2016 chartEx
  차트(깔때기, 폭포 등)는 내장 workbook으로 변환합니다. 내장 workbook도 패키지와
  같은 기준으로 검사하며, 거부되면 cache 값을 사용합니다.
- SmartArt 텍스트 데이터
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

