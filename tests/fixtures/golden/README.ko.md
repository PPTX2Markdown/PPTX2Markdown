# 저장소 골든 PPTX 모음

[English](README.md) | 한국어

이 디렉터리에는 테스트 환경에서 항상 사용할 수 있어야 하는 작고 검토 가능한
PPTX corpus가 있습니다. `.pptx2markdown/corpus/`의 더 큰 선택적 corpus를
보완합니다.

- `synthetic_reading_order.pptx`는 프로젝트 소유 fixture이며 XY-cut의 좌표
  허용치를 격리해 검증합니다. 행/열 흔들림, 맞닿거나 겹친 열 경계, subpixel
  gutter, 서로 다른 높이, 전체 너비 영역, 빈 도형을 포함합니다.
- `synthetic_text_objects.pptx`는 프로젝트 소유 fixture이며 텍스트 run, 목록,
  링크, 숫자 label, Unicode variant, 병합 표, 네이티브 차트, 패키지 절대경로
  발표자 노트 relationship, 회전, 혼합 font 크기를 격리해 검증합니다.
- 나머지 15개 파일은 MIT 라이선스 외부 프로젝트에서 선택한 작고 가치가 높은
  fixture입니다. 정확한 source, revision, license, 입력 hash, 출력 hash, 추출
  asset hash는 `manifest.json`에 기록합니다. 필요한 저작권 표시와 라이선스
  조건은 `THIRD_PARTY_NOTICES.ko.md`에서 확인할 수 있습니다.

이 모음은 원본 PPTX와 검토된 파서 스냅샷을 함께 저장합니다.

- `expected/json/`: PPTX마다 완전한 presentation-document JSON 결과 하나
- `expected/markdown/`: 그에 대응하는 Markdown 결과
- 렌더 screenshot: 임시 QA 산출물이며 저장소에 포함하지 않음

`tests/test_repository_goldens.py`는 모든 파일을 두 형식으로 변환하고, 생성된
byte를 저장된 스냅샷과 비교하며, 매니페스트 hash와 생성된 모든 media/attachment
hash를 확인합니다.

이 모음만 실행:

```bash
python -m unittest tests.test_repository_goldens -v
```

현재 스냅샷 검사:

```bash
python scripts/update_goldens.py --check
```

의도적인 파서 또는 스키마 변경으로 출력이 바뀌면 `expected/` diff를 직접
검토한 뒤 스냅샷과 `manifest.json`을 함께 갱신하세요. 모든 JSON, Markdown,
asset 변경을 검토한 경우에만 다음을 실행합니다.

```bash
python scripts/update_goldens.py --accept
```

변환 프로세스가 단순히 성공 종료했다는 사실만으로 새 골든 결과를 승인하면 안
됩니다.

`expected/markdown/` 아래 파일은 사람이 작성한 문서가 아니라 파서의 회귀
출력이므로 영문·한글 문서 쌍 정책에서 의도적으로 제외합니다.
