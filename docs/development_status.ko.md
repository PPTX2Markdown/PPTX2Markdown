# 개발 상태 및 인수인계

[English](development_status.md) | 한국어

마지막 검증: 2026-07-13

이 문서는 현재 개발 상태를 기록합니다. 작업 트리를 변경하기 전에
`git status --short --branch` 결과와 함께 읽으세요.

## 저장소 상태

- 활성 개발 브랜치: `dev`
- 출시 목표: `0.1.0`
- 출시 pull request: `dev` -> `main`, PR #17
- 공개 패키지 구조: `src/pptx2markdown/`
- 패키징 backend: Hatchling
- 지원 Python: 3.12 이상
- 로컬 sample, output, work, archive 디렉터리는 출시 source가 아님

PyPI 구조 변경 전 source는 `archive-pre-pypi-restructure` 브랜치와
`pre-pypi-restructure` annotated tag로 보존되어 있습니다.

## 제품 계약

- 변환은 정적입니다. Surya, OCR, VLM, model provider, API key, cloud inference
  경로가 남아 있지 않습니다.
- 네이티브 도형 좌표 기반 XY-cut이 유일한 읽기 순서 알고리즘입니다.
- Markdown과 JSON은 하나의 `PresentationDocument 1.0`에서 렌더링합니다.
- 출시 전 내부 구조에 대한 호환 계층은 유지하지 않습니다.
- 일반 `.pptx` 변환은 로컬에서 동작합니다. 구형 `.ppt` 변환은 Windows의
  PowerPoint 또는 LibreOffice를 실행할 수 있습니다.
- 최종 출력과 중간 작업은 서로 분리된 설정 가능 루트를 사용합니다.
- `~$*.pptx` 형식의 Office 잠금 파일은 조용히 건너뜁니다.

## 출력 계약

패키지 JSON Schema는
`src/pptx2markdown/presentation_document.schema.json`입니다. 다음을
규정합니다.

- basename만 사용하는 입력 식별자
- 양수이며 오름차순인 슬라이드 페이지
- 숨김 슬라이드와 발표자 노트
- 닫힌 콘텐츠 블록 kind
- heading level 불변조건
- 선택적 EMU 좌표
- slide/layout/master 출처
- 결정론적 상대 asset 링크

[출력 스키마](output_schema.ko.md)를 참고하세요.

## 회귀 테스트 자료

항상 실행되는 저장소 모음에는 PPTX 17개가 있습니다.

- 프로젝트 소유 synthetic 프레젠테이션 2개
- revision이 고정된 MIT 라이선스 외부 프레젠테이션 15개
- 검토 완료된 전체 JSON 스냅샷
- 검토 완료된 전체 Markdown 스냅샷
- 입력, 출력, 생성 asset의 SHA-256 hash
- 고정된 출처와 외부 저작권 표시

추가로 ignore된 corpus를 넓은 로컬 검증에 사용하지만 CI 의존성은 아닙니다.
[골든셋 문서](../tests/fixtures/golden/README.ko.md)를 참고하세요.

## 최근 검증 기준선

2026-07-13에 다음 검증을 모두 통과했습니다.

- Ruff lint 및 format 검사
- 단위·회귀 테스트 87개
- 저장소 골든 PPTX 17개를 JSON 및 Markdown 모드로 변환
- 실제 사례에서 고른 PPTX 51개, 슬라이드 633개
- wheel 및 sdist 빌드와 `twine check`
- 격리된 wheel 설치와 설치된 CLI 변환
- Ubuntu, macOS, Windows 및 Python 3.12, 3.13 GitHub Actions

PR #17 전용 workflow도 통과했습니다.

## 보안 경계

공통 OOXML 압축 해제 단계는 경로 탈출, symlink, 암호화 entry, relationship
탈출, 과도한 entry 수, 과도한 압축 해제 크기, 의심스러운 압축률을 거부합니다.
추출된 첨부파일은 여전히 신뢰할 수 없는 입력 데이터이므로 자동으로 열면 안
됩니다.

## 알려진 제한사항

1. 이미지 내부 텍스트를 OCR하지 않습니다.
2. 오디오와 비디오를 전사하지 않습니다.
3. connector 의미, 애니메이션, transition을 재구성하지 않습니다.
4. 출력은 문서 구조를 보존하며 픽셀 단위 슬라이드 외형을 재현하지 않습니다.
5. PowerPoint 검토 댓글은 의도적으로 제외합니다.
6. 넓은 로컬 corpus 테스트는 ignore된 source 파일이 없으면 건너뜁니다.

## 남은 출시 작업

1. PR #17 CI를 통과 상태로 유지하고 문서 검토를 완료합니다.
2. `publish.yml`, environment `pypi`용 PyPI Trusted Publisher를 설정합니다.
3. pull request를 Ready로 전환하고 `main`에 병합합니다.
4. `v0.1.0` GitHub release를 발행합니다.
5. 깨끗한 환경에서 공개된 PyPI 패키지를 검증합니다.

## 중요 파일

- `src/pptx2markdown/api.py`: 공개 Python API
- `src/pptx2markdown/main_converter/run_pptx_to_markdown.py`: orchestration과 CLI
- `src/pptx2markdown/main_converter/slide_converter.py`: slide-to-document 파싱
- `src/pptx2markdown/main_converter/converter_models.py`: 문서 model과 renderer
- `src/pptx2markdown/structure_analyzer/structure.py`: XY-cut과 헤딩 분석
- `src/pptx2markdown/ooxml_security.py`: 공통 패키지 보안 정책
- `scripts/update_goldens.py`: 골든 검증과 통제된 갱신
- `tests/test_repository_goldens.py`: 저장소 골든 회귀 테스트

## 새 작업 맥락 점검표

```bash
git status --short --branch
python -m unittest discover -v
python scripts/update_goldens.py --check
ruff check src tests scripts
ruff format --check src tests scripts
```

수동 변환 검증에는 새로운 임시 output 및 work 디렉터리를 사용하세요. 관련 없는
로컬 변경을 보존하세요. 파서 또는 스키마 변경으로 출력이 바뀌면 스냅샷을
승인하기 전에 JSON과 Markdown diff를 모두 검토하세요.
