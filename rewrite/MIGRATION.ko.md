# 배포된 변환기에서 rewrite로 옮기기

[English](MIGRATION.md)

rewrite는 같은 패키지 이름 `pptx2markdown`으로 배포되므로, 업그레이드하면 배포된
변환기를 대체합니다. 0.1.x 사용자에게 무엇이 바뀌는지 정리합니다.

## 명령줄

| 0.1.x | rewrite |
| --- | --- |
| `--output-dir`, `-v` | 동일 |
| `--output-format markdown\|json` | 동일, 검색용 JSON Lines인 `chunks` 추가 |
| `--headings auto\|strict` | 동일. `strict`는 레시피 설정 `headings: "placeholder"`와 같음 |
| `--merged-cells fill\|blank` | 동일, 머리글 행에서만 반복하는 `header` 추가 |
| `--work-dir` | 필요 없음: 번들 밖에는 아무것도 쓰지 않음 |
| `--placeholder-inheritance` | 항상 켜짐(`style`): placeholder가 위치·크기·글머리를 상속 |
| `--inherited-shapes` | 항상 `visible`: 레이아웃·마스터에 그려진 텍스트와 그림을 그룹 안까지 유지. 슬라이드를 덮는 템플릿 그림은 배경으로 보고 제외(레시피 `template_pictures: false`는 템플릿 그림 전체 제외) |
| `.ppt` 입력, `--ppt-converter auto\|powerpoint\|libreoffice` | 명시적으로 켤 때만: `--ppt-converter libreoffice`(기본값 `none`). PowerPoint 자동화는 지원하지 않음 |
| `--convert-vector-images` | 지원하지 않음: EMF/WMF 이미지는 그대로 유지 |
| — | `--no-assets`, `--overwrite`, `--tolerant`, `--recipe`, `--review`, `--audit`, `--ask`, `--max-chunk-chars`, `--recipe-schema` |

입력은 직접 지정해야 하며, 입력 없이 실행해도 현재 디렉터리의 모든 프레젠테이션을
변환하지 않습니다.

## Python API

| 0.1.x | rewrite |
| --- | --- |
| `convert(inputs, output_dir=...) -> int`(종료 코드) | `convert(path, output_dir=...) -> SaveResult`, `convert_many(paths) -> BatchConversionResult`(`.exit_code` 제공) |
| — | `parse(경로 \| bytes \| 스트림) -> ParsedPresentation`: 메모리의 문서, 파일을 쓰지 않음 |
| — | `slide_chunks()`, `Recipe`, `review()`, `pptx2markdown.agent`, `pptx2markdown-mcp` |

`convert()`는 더 이상 logging 설정을 바꾸지 않습니다.

## 출력

번들 배치는 같습니다: `<출력>/<덱>/<덱>.md`(또는 `.json`)와 `media/`,
`attachments/`. 실행 전체에 대한 `convert_manifest.json` 하나 대신, 번들마다 자산
해시를 담은 `bundle-manifest.json`이 있습니다.

JSON은 `PresentationDocument` 2.0입니다. 블록이 Markdown `content` 문자열 대신
종류별 필드를 가집니다.

| 종류 | 1.0 필드 | 2.0 필드 |
| --- | --- | --- |
| text, heading | `content`(+ `heading_level`) | `text`(+ `level`, `source`) |
| list | `content`(Markdown) | `text`와 `level`을 가진 `items[]`, `ordered`, `start` |
| table | `content`(Markdown) | `markdown`과 셀 격자 `rows`(범위·병합), `header_rows`, `caption_rows` |
| math | `content` | `latex` |
| image | `content`(Markdown 링크) | `asset_id`, `alt_text`, `description` |
| attachment | `content`(Markdown 링크) | `asset_id`, `label`(파일 이름), `program` |
| chart, smartart | `content` | `markdown`(+ `asset_ids`) |

모든 블록에는 `id`(`s3.b2`)도 있습니다. 문서는 해시와 media type이 담긴 `assets`
목록, 원본의 SHA-256, 구조화된 `warnings`를 가집니다.

## 동작

Markdown에서 보이는 차이로, 모두 의도된 것입니다.

- Markdown이 해석할 원문 문자는 이스케이프하고(`\# x`, `\<tag>`), 텍스트 일부를
  강조하는 굵게·기울임은 유지합니다.
- 번호 매긴 항목 여러 개는 긴 제목 하나가 아니라 목록입니다.
- 표 전체 폭으로 병합된 첫 행은 표 위의 제목입니다.
- 한 도형 안의 텍스트·목록·수식은 각각 다른 블록입니다.
- 글머리 기호 없이 `lvl`로만 들여 쓴 문단은 목록 항목이 아닙니다.
- OLE 첨부는 안의 파일로 풀고 그 이름을 씁니다.
