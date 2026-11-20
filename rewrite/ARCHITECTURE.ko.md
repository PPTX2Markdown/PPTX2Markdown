# 새 구현 아키텍처

[English](ARCHITECTURE.md)

## 의존성 방향

```text
api / cli
    ↓
parser → ooxml / inheritance / reading_order / headings / omml / charts / tables
    ↓
model + AssetStore
    ↓
ParsedPresentation
    ↓
renderer
    ↓
transactional bundle writer
```

표준 문서는 로컬 경로나 binary payload를 소유하지 않습니다. Runtime 소유권은
`ParsedPresentation`에 있으며, 이 객체가 `AssetStore`를 닫습니다.

## Asset 식별

각 OOXML 논리 part의 ID는 정규화된 package part 이름으로부터 생성합니다.
실제 byte의 SHA-256은 별도로 기록합니다. 따라서 논리 식별자는 안정적으로
유지하면서 동일 payload는 내부에서 중복 저장하지 않을 수 있습니다.

작은 payload는 메모리에 남고 spool 임계값을 넘으면 운영체제가 관리하는 임시
파일로 이동합니다. 이 임시 파일 경로는 공개되지 않습니다.
`ParsedPresentation.close()`가 모든 payload를 정리합니다.

## 렌더링과 저장

이미지와 첨부파일 블록은 경로가 포함된 Markdown이 아니라 `asset_id`를
가집니다. Markdown renderer는 URL resolver 또는 명시적인 asset 제외 정책을
요구합니다.

Bundle writer는 asset을 먼저 저장하고 크기와 SHA-256을 검증합니다. 그 뒤
상대경로를 렌더링하고, 결정론적 manifest를 쓰고, 모든 경로가 staging root
안에 있는지 검사한 후 staging 디렉터리를 최종 위치로 rename합니다.
`overwrite="replace"`를 명시하지 않으면 기존 출력은 변경하지 않습니다.

## 결정론

표준 문서에는 timestamp, 절대 입력 경로, work 경로, 임시 이름을 넣지 않습니다.
Asset과 warning은 안정적으로 정렬합니다. Bundle 파일 이름이 충돌하면 content
hash suffix를 사용합니다.

## 보안 경계

어떤 part도 읽기 전에 ZIP entry 수, 암호화 entry, symlink, 경로 탈출, 정규화
후 중복 이름, 압축 해제 크기, 압축률을 검사합니다. Relationship target은 source
part 기준으로 정규화하며 ZIP 밖으로 이동할 수 없습니다. DTD 또는 entity 선언이
있는 XML은 거부합니다.

chart2md는 자체 `zipfile`이 아니라 검증된 패키지 위의 어댑터로만 차트 part를
읽습니다. 차트에 내장된 workbook은 그 자체로 ZIP이므로, chart2md가 열기 전에
entry 수·크기·압축률 검사와 XML part의 DTD 검사를 똑같이 거칩니다. 거부된
workbook은 warning을 남기고, 차트는 cache 값으로 변환합니다.

