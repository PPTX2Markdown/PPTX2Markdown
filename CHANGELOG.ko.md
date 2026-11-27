# 변경 기록

[English](CHANGELOG.md)

PyPI `pptx2markdown` 패키지의 주요 변경 사항입니다.

## 0.2.0 — 2026-11-27

새 엔진이 0.1 변환기를 대체합니다. 기존 변환기 옆에서 개발하며 모든 변경마다 비교
검증했고(fixture와 966개 덱 코퍼스에 대한 parity 하네스), 그
이상으로 나아갔습니다. 바뀐 CLI 옵션, API, 출력은
[마이그레이션 가이드](docs/MIGRATION.ko.md)를 참고하세요.

### 추가

- **Python API**: 경로·bytes·stream을 받는 `parse()`가 메모리 안의 문서를 돌려주고,
  `convert()` / `convert_many(jobs=N)`는 staging 디렉터리를 거쳐 이동 가능한 번들을
  씁니다. 에셋은 파싱된 프레젠테이션이 소유합니다.
- **슬라이드 구조를 따르는 읽기 순서**: 그룹과 감싸는 프레임은 한 단위로 읽고, 배경은
  무시하며, 격자는 행 단위로 읽습니다. 여러 방식으로 읽힐 수 있는 슬라이드는 후보
  순서와 함께 표시합니다.
- **결정 지점**: `review()`가 확신 없는 결정(읽기 순서, 표 머리글 행과 캡션, 제목,
  소제목, 이미지 설명)을 고정 id, 사유 코드, 봐야 할 것, 중요도, 선택지, 추천과 함께
  나열합니다(review 스키마 버전 1). `recipe_from_answers()`는 답을 레시피로 바꿉니다.
  [변환기–에이전트 계약](docs/agent_contract.ko.md)을 참고하세요.
- **레시피**: 블록 id(`s3.b2`)와 선택자(`slide[12-]`, `section[Appendix]/image`,
  `table[0]`)를 쓰는 JSON 규칙으로 내용 제외, 표 머리글·캡션 행, 제목 수준, 이미지
  설명, 노트, 읽기 순서를 정합니다. `--recipe`, `--recipe-schema`.
- **감사**: `audit()` / `--audit`가 슬라이드 XML을 다시 읽어 출력에 없는 입력
  텍스트를 보고합니다.
- **검색용 청크**: `slide_chunks()`와 `--output-format chunks`(출처 메타데이터가 붙은
  JSON Lines, `--max-chunk-chars`).
- **MCP 서버**(`pptx2markdown[mcp]`, `pptx2markdown-mcp`): 변환, 검토, 답변, 감사,
  청크, 레시피 확인, 그리고 그림과 표시된 슬라이드를 호스트 모델에게 보여 주기.
- 모델 무관 판단 루프: `describe_images_with()`, `choose_reading_orders_with()`,
  `snapshot.slide_pdf()` / `slide_png()`.
- 표는 JSON에 병합 범위를 포함한 전체 셀 격자를 유지하고, 표 전체 너비로 병합된
  본문 행은 한 번만 쓰며, 그룹 머리글은 두 행으로 지정할 수 있습니다.
- 인라인 굵게·기울임, Markdown 이스케이프, PowerPoint 구역, ISO Strict OOXML, 내장
  통합 문서를 읽는 chartEx 차트, ProgID와 함께 풀어내는 OLE 첨부, 세로쓰기 합치기,
  레이아웃·마스터 그림.
- `--jobs`로 여러 덱을 작업 프로세스에서 변환합니다.

### 변경

- CLI 옵션, Python API, JSON 스키마(`schema_version` 2.0)가 새로 바뀌었습니다.
  마이그레이션 가이드를 참고하세요.
- `.ppt` 입력과 EMF/WMF→PNG는 요청할 때만(`--ppt-converter libreoffice`,
  `--vector-images png`) LibreOffice로 실행합니다.
- 작업 디렉터리가 없습니다. 출력 번들 밖에는 아무것도 쓰지 않습니다.

### 제거

- 0.1 변환기와 그 옵션 `--work-dir`, `--placeholder-inheritance`,
  `--inherited-shapes`, `--convert-vector-images`, 그리고 PowerPoint COM을 통한
  `.ppt` 변환. 0.1 코드는 `legacy-0.1` 브랜치에 남아 있습니다.

## 0.1 계열 (미출시 수정, `legacy-0.1` 브랜치)

0.1.2 이후 0.1 변환기에 들어갔지만 배포하지 않은 수정입니다. 0.2.0 엔진에는 모두
들어 있습니다.

### 변경

- 병합된 표 셀은 머리글과 본문 구분 없이 병합 범위의 모든 칸에 값을 반복하므로,
  행 하나만 봐도 뜻이 통합니다(`--merged-cells fill`, 기본값).
  `--merged-cells blank`(API: `merged_cells="blank"`)는 병합이 시작되는 칸에만 값을
  둡니다. 이전에는 머리글 병합만 반복했습니다.
- 슬라이드는 프레젠테이션의 슬라이드 목록(`p:sldIdLst`) 순서를 따르고 위치대로
  번호를 매깁니다. 이전에는 PowerPoint에서 순서를 바꾼 덱이 파일 이름 순서로
  나오고, 프레젠테이션에 없는 슬라이드 part도 변환했습니다.

### 수정

- 터무니없는 점 개수(`ptCount="999999999999"`)를 선언한 차트 때문에 차트 변환기가
  메모리를 모두 쓰던 문제. 이제 해당 차트는 변환 실패로 보고합니다.
- 라틴 문자가 아닌 Windows 코드 페이지로 이름이 저장된 내장 파일(`보고서.hwp`,
  `新建文本文档.txt`) 이름이 깨지던 문제. Packager가 함께 저장한 유니코드 이름을 씁니다.
- 가로·세로 병합이 모두 걸린 2D 병합 블록 안의 칸이 왼쪽 이웃을 가리켜 비어 있던
  문제.
- 행이 없는 표 하나 때문에 슬라이드 전체가 실패하던 문제. 이제 경고와 함께
  건너뛰고 `table_skipped_blocks`에 집계합니다.
- 슬라이드 텍스트의 원시 HTML(예: `<script>`)이 Markdown에 실제 HTML로 들어가던
  문제. 이제 이스케이프합니다.
- `...pptx` 같은 입력 이름이 `--work-dir` 전체를 지울 수 있던 문제, 같은 이름의
  덱이 서로 덮어쓰던 문제, 발표자 노트 관계로 패키지 밖 파일을 읽을 수 있던 문제,
  대소문자나 Unicode 정규화만 다른 zip 항목과 대소문자만 다른 첨부 이름이 충돌하던
  문제, 이전 실행의 출력이 남아 있던 문제.
- 직접 지정한 Office 잠금 파일(`~$deck.pptx`)을 아무 말 없이 건너뛰던 문제. 이제
  경고를 남깁니다.

### 성능

- 덱마다 원본 패키지를 한 번만 열고, 레이아웃·마스터·관계 part를 한 번만
  파싱합니다. 1000장 덱이 12.3초에서 약 5.5초로 줄었습니다.

## 0.1.2 — 2026-09-16

- 프로젝트 저자와 유지관리자 명시, NOTICE 저작권 갱신.

## 0.1.1 — 2026-07-29

- 이미지 렌더링과 라이선스 메타데이터 수정.
- 원본 벡터(EMF/WMF) 이미지 자산 보존. `--convert-vector-images`로 PNG 변환을
  선택할 수 있으며 덱 단위로 일괄 처리합니다.

## 0.1.0 — 2026-07-13

- PyPI 첫 배포: 정적 OOXML 파싱, XY-cut 읽기 순서, 하나의
  `PresentationDocument`에서 만드는 Markdown과 JSON, 골든 회귀 테스트.
