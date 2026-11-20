# 새 구현 아키텍처

[English](ARCHITECTURE.md)

## 의존성 방향

```text
api / cli                                agent (선택, anthropic)
    ↓                                        ↓ 레시피만 작성
parser → ooxml / inheritance / reading_order / headings / inline
       → omml / charts / smartart / tables / attachments
    ↓
model + AssetStore
    ↓
ParsedPresentation ←── recipe (apply_recipe: 문서 → 문서)
    ↓                  review (문서 → 레시피 규칙이 달린 질문)
renderer / chunks
    ↓
transactional bundle writer
```

표준 문서는 로컬 경로나 binary payload를 소유하지 않습니다. Runtime 소유권은
`ParsedPresentation`에 있으며, 이 객체가 `AssetStore`를 닫습니다.


## 결정은 데이터로

읽기 순서, 제목, 목록 구조, 표 머리글과 제목은 항상 parser가 결정합니다. 파싱
이후의 모든 단계는 표준 문서만 다룹니다.

- `recipe.apply_recipe`는 문서와 레시피를 받아 문서를 돌려주는 순수 함수입니다.
  selector는 읽기 순서대로 매긴 블록 id(`s<페이지>.b<번호>`)나 슬라이드·종류로
  블록을 고르며, 항상 파싱된 그대로의 문서를 보므로 규칙의 대상이 앞 규칙에 따라
  달라지지 않습니다.
- `review.review`는 문서를 읽고 다시 볼 만한 결정을 골라, 대안을 바로 쓸 수 있는
  레시피 규칙으로 함께 돌려줍니다.
- `agent`(선택)는 레시피만 만듭니다. 요청·검토 질문·이미지를 받아 검증된 레시피를
  내놓고, `explain_recipe`로 모든 규칙이 무언가를 고르는지 확인합니다. 모델은
  본문이나 parser를 고치지 않습니다.

표는 전체 격자(`rows`, 병합 범위, `merged_from`)를, 제목은 제목이 된 근거
(`source`)를 남기므로, 레시피는 패키지를 다시 파싱하지 않고 블록을 다시
렌더링하거나 분류를 바꿀 수 있습니다.

## Inline Markdown

텍스트·제목·목록·표 셀의 텍스트는 이스케이프된 inline Markdown입니다. parser가
텍스트 규칙(제목, 직접 입력한 글머리)을 적용하는 동안 생성한 마크업(링크, 문장 속
수식, 강조)은 private-use 표지 문자로 감싸 두고, 블록을 만들 때
`inline.finalize`가 원문 문자만 이스케이프한 뒤 표지를 지웁니다. 발표자 노트는
모델에서는 평문으로 두고 렌더링할 때 이스케이프합니다.

## 읽기 순서

`reading_order`는 재귀 XY-cut입니다. 영역을 관통하는 세로 공백이 있으면 열로
나눠 한 열씩 읽고, 없으면 맨 위 띠를 먼저 읽은 뒤 나머지를 다시 나눕니다. 촘촘히
쌓인 단일 상자들은 위에서 아래로 읽고, 1px까지 겹쳐도 공백으로 보며, 바닥글과
장식 도형은 공백을 찾는 데 쓰지 않습니다.

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


내장 OLE 복합 파일은 asset 한도로 크기가 제한된 메모리 버퍼에서 olefile로
풉니다. olefile은 짧은 bytes를 파일 경로로 읽기 때문에 bytes를 그대로 넘기지
않습니다.
