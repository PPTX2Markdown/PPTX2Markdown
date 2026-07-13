# 골든 출력 검토

[English](REVIEW.md) | 한국어

이 검토 문서는 변환 성공을 곧바로 정답으로 취급하지 않고 저장소에 포함된 파서
결과를 분류합니다. 스냅샷을 승인할 때마다 다시 검토해야 합니다.

## 예상 결과로 확인한 동작

- 복잡한 이미지·텍스트 레이아웃이 기하학적 순서와 로컬 이미지 링크를 유지
- 내용이 없는 유효한 프레젠테이션이 빈 슬라이드 목록과 0 byte Markdown 생성
- SVG fallback, 오디오·비디오 poster, 3D preview, source payload가 안정적인
  상대경로로 출력
- 복구 가능한 경우 OLE와 Package payload를 원래 Office, JSON, 오디오, 비디오,
  GLB 형식으로 압축 해제
- 네이티브 차트, 표, SmartArt, hyperlink, 목록, group 도형, 회전, CJK 텍스트,
  일반 숫자 label이 검색 가능한 콘텐츠로 유지
- 날짜와 슬라이드 번호 field를 제외하되 인접 텍스트는 보존
- 발표자 노트는 화면 블록과 분리하고 명시적인 `[Speaker_Notes]` 아래에만 렌더링

## 골든셋으로 고정한 정책

- 읽기 순서는 항상 기하학적 재귀 XY-cut을 사용하며 대체 읽기 순서 옵션이나
  의미 label override가 없음
- 열 경계 1px 겹침은 열 분할로 허용하지만 24px 겹침 fixture는 row-major로
  fallback
- 0.5px 양수 gutter는 열을 나누기에 충분
- 1~3px 정렬 흔들림과 서로 다른 텍스트 상자 높이에 별도 row-clustering이
  필요하지 않음
- 전체 너비 영역은 내부 grid 전후 순서를 유지. 촘촘히 정렬된 grid fixture에서는
  수평 분할로 row-major 순서를 생성
- 병합 header cell은 Markdown header 열마다 값을 반복하고 병합 body cell은 병합
  시작점에만 콘텐츠 유지
- Unicode는 작성된 형태 그대로 보존. 사용자 출력에서 NFC/NFD 및 full-width
  variant를 다시 쓰지 않음
- label 없는 SmartArt는 버리거나 가짜 label을 만들지 않고 안정적인 node-count
  placeholder로 표현

## 골든셋 구축 중 발견하고 해결한 결함

- `/ppt/notesSlides/notesSlide6.xml` 같은 패키지 절대경로 노트 relationship을
  filesystem root로 잘못 해석하던 문제를 OOXML 패키지 root 기준으로 수정
- `~$*.pptx` 형식의 Office 잠금 파일을 변환 및 골든 탐색에서 모두 제외

## 열린 품질 backlog

현재 저장된 결과 중 알려진 파서 결함으로 분류된 것은 없습니다. 새로운 실패가
관찰되면 먼저 최소 PPTX와 검토된 스냅샷을 추가한 뒤 파서를 수정해야 합니다.
새로운 정책 경계나 실패 유형 없이 corpus 크기만 늘리는 작업은 이 명시적 기대를
유지하는 것보다 우선순위가 낮습니다.
