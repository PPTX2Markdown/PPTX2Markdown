# 개발 현황과 인수인계

[English](development_status.md) | 한국어

마지막 확인: 2026-11-27

현재 엔지니어링 상태를 기록합니다. 작업 트리를 바꾸기 전에
`git status --short --branch`와 함께 읽으세요.

## 저장소 상태

- 개발 브랜치: `dev` (PR은 `dev`로 병합, 출시는 `dev` -> `main` 후 GitHub
  release가 PyPI에 배포)
- 배포 버전: `0.2.0`, 0.1 변환기를 대체한 엔진. 변경 사항은
  [CHANGELOG.ko.md](../CHANGELOG.ko.md)
- 0.1 변환기는 `legacy-0.1` 브랜치에 보존(버그 수정만, 출시 계획 없음). 제거한
  Claude 에이전트 계층은 별도 멀티포맷 에이전트 프로젝트를 위해 `archive/agent`에 보존
- 패키지 구조: `src/pptx2markdown/`, Hatchling, Python 3.12 이상

## 제품 계약

- 변환은 정적이며 모델을 쓰지 않습니다. 패키지는 어떤 모델 SDK도 import하지
  않습니다(테스트로 강제). 판단은 레시피, `review()` 답, 판단 루프로 연결합니다.
  [변환기–에이전트 계약](agent_contract.ko.md)을 참고하세요.
- 같은 입력과 옵션은 지원하는 모든 OS에서 같은 바이트를 냅니다.
- `.pptx` 변환은 로컬에서 일어납니다. LibreOffice는 `.ppt` 입력, EMF/WMF의 PNG
  변환, 표시된 슬라이드 렌더링처럼 요청할 때만 실행합니다.
- 출력 번들 밖에는 아무것도 쓰지 않습니다.

소유권, 수명주기, 보안은 [ARCHITECTURE.ko.md](ARCHITECTURE.ko.md), 읽기 순서는
[읽기 순서 설계](reading_order_design.ko.md)를 참고하세요.

## 회귀 자산

- `tests/fixtures/golden/`: PPTX 17개(프로젝트 소유 합성 덱 2개, 고정된 MIT
  라이선스 상류 덱 15개)와 검토된 Markdown·JSON 스냅샷, 에셋 해시.
  `scripts/update_goldens.py --check`와 `tests/test_goldens.py`가 바이트 단위로
  비교합니다. [golden 모음](../tests/fixtures/golden/README.ko.md)을 참고하세요.
- `tests/fixtures/xycut_*`: 기대 읽기 순서.
- `benchmarks/`: MarkItDown·anydoc과의 기능 비교(20개 중 20개 유지, anydoc 13,
  MarkItDown 8)와 합성 읽기 순서 케이스(11개 중 9개, 나머지 둘은 올바른 후보와 함께
  표시).
- 넓은 검증용 로컬 corpus(`.pptx2markdown/corpus/`, 덱 966개). CI 의존성은 아닙니다.

0.2.0 이전에는 엔진을 0.1 변환기와 도형 단위로 비교했습니다(`scripts/parity.py`,
현재 `legacy-0.1`에 있음). corpus에서 legacy 단어의 99.8%를 유지했고, `audit()`는
어떤 corpus 덱에서도 빠진 텍스트를 찾지 않았습니다.

## 알려진 한계

1. 이미지 안의 텍스트는 OCR하지 않고, 오디오·비디오는 전사하지 않습니다.
2. 연결선 의미, 애니메이션, 전환은 복원하지 않습니다.
3. 슬라이드 레이아웃·마스터에 그려진 표는 읽지 않습니다.
4. PowerPoint 검토 메모는 의도적으로 제외합니다.

## 다음 단계

1. 결정 지점 계약 위에 멀티포맷 에이전트를 별도 프로젝트로 만듭니다
   (`archive/agent`에서 시작).
2. 한국어 비즈니스 덱에서 읽기 순서와 검토 항목을 검증합니다. 로컬 corpus에는
   거의 없습니다.
3. omml2latex가 Python 3.10/3.11에서 돌도록 고쳐 지원 버전을 넓힙니다.

## 새 맥락 체크리스트

```bash
git status --short --branch
python -m pip install -e ".[mcp]"
python -m unittest discover -v
python scripts/update_goldens.py --check
ruff check src tests scripts benchmarks
ruff format --check src tests scripts benchmarks
```
