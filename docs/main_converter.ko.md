# 메인 변환기

[English](main_converter.md) | 한국어

메인 변환기는 하나 이상의 PowerPoint 파일을 공통 `PresentationDocument`
표현으로 바꾼 뒤 Markdown 또는 JSON으로 렌더링합니다. 설치 후 진입점은
`pptx2markdown`이며 구현은 `src/pptx2markdown/main_converter/`에 있습니다.

## 파이프라인

1. `.pptx`와 구형 `.ppt` 입력을 탐색합니다.
2. 선택한 외부 변환기로 `.ppt`를 `.pptx`로 변환합니다.
3. work 디렉터리 아래에서 OOXML 패키지를 검증하고 압축 해제합니다.
4. 슬라이드, 레이아웃, 마스터 상속을 해석합니다.
5. 기하학적 재귀 XY-cut으로 네이티브 도형 순서를 정합니다.
6. 도형을 순서가 있는 `ContentBlock` 객체로 파싱합니다.
7. 하나의 `PresentationDocument`를 만들고 검증합니다.
8. 결정론적 Markdown 또는 표준 JSON을 렌더링합니다.
9. 일괄 변환 매니페스트를 기록합니다.

Markdown과 JSON이 별도의 파서를 사용하지 않습니다.

## 입력 정책

인자를 지정하면 선택한 각 `.pptx` 또는 `.ppt`를 처리합니다.

```bash
pptx2markdown first.pptx second.ppt
```

인자를 생략하면 현재 디렉터리와 work 디렉터리의 `target_pptx/` staging
디렉터리에서 프레젠테이션을 찾습니다.

```bash
pptx2markdown
```

`~$*.pptx` 형식의 임시 Office 잠금 파일은 조용히 건너뜁니다. 일반적인 누락
입력은 계속 오류로 처리합니다.

## 출력 정책

기본 최종 출력 루트는 `./output`, 기본 work 루트는
`./.pptx2markdown`입니다.

```text
output/
├── convert_manifest.json
└── <deck>/
    ├── <deck>.md           # 또는 <deck>.json
    ├── media/
    └── attachments/

.pptx2markdown/
├── target_pptx/
├── target_slides/
├── structure_analysis/
├── table_pipeline/
└── .cache/
```

최종 문서와 복사된 asset은 output 루트 아래에만 기록합니다. 압축 해제 패키지,
분석 sidecar, 캐시는 work 루트 아래에만 기록합니다.

## 읽기 순서

모든 슬라이드는 네이티브 도형 bbox 기반 재귀 XY-cut을 사용합니다. 공개 읽기 순서
옵션은 없습니다. 텍스트, 번호, 헤딩, placeholder, 장식 분류는 순서를 바꾸지
않습니다. XML index는 좌표가 같은 경우의 결정론적 tie-break와 유효한 좌표가 없는
객체에만 사용합니다.

## 공개 CLI 옵션

```text
--output-dir <path>
--work-dir <path>
--output-format markdown|json
--headings auto|strict
--placeholder-inheritance none|geometry|style
--inherited-shapes none|visible|all
--ppt-converter auto|powerpoint|libreoffice
--verbose
```

정확한 옵션은 `pptx2markdown --help`를 기준으로 합니다.

## Python API

```python
import pptx2markdown

exit_code = pptx2markdown.convert(
    ["first.pptx", "second.pptx"],
    output_dir="converted",
    work_dir="work",
    output_format="json",
)
```

선택한 모든 슬라이드 변환에 성공하면 `0`, 하나라도 실패하면 `1`을 반환합니다.

## 관련 문서

- [출력 스키마](output_schema.ko.md)
- [구조 분석기](structure_analyzer.ko.md)
- [저장소 골든셋](../tests/fixtures/golden/README.ko.md)
