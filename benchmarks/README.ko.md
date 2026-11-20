# PPTX 변환기 벤치마크

[English](README.md)

PowerPoint 덱에서 각 변환기가 무엇을 보존하는지 기능별로 확인합니다.

`decks.py`는 python-pptx로 기능마다 작은 덱을 하나씩 만들고(바이너리를 저장하지
않으므로 언제든 다시 만들어 PowerPoint로 열어 볼 수 있음), 변환기가 돌려준
Markdown을 검사하는 함수와 짝지습니다. `run.py`는 설치된 모든 변환기를 모든 덱에
실행하고 `results/results.md`와 `results/results.json`을 씁니다.

```bash
python3.12 -m venv .bench && . .bench/bin/activate
pip install -r benchmarks/requirements.txt
python benchmarks/run.py --large 300 --keep /tmp/bench   # 덱과 출력 보관
```

## 결과

최신 실행 결과는 [results/results.md](results/results.md)에 있습니다(버전, Python,
플랫폼은 `results.json`에 기록됨).

## 결과를 읽는 법

- 기능 항목은 발표 자료를 이해하는 검색(RAG)에 필요한 것으로 골랐습니다. 인용을
  위한 슬라이드 경계, 제목, 읽기 순서, 표와 병합 셀, 수식, 노트, 숨김 슬라이드,
  이미지 설명입니다. pptx2markdown이 중점을 두는 기능이라 항목 선정이 이 도구에
  유리합니다. 범용 변환기는 다른 균형을 택합니다.
- anydoc은 14개 포맷을 다루는 Rust 라이브러리이며 약 20배 빠릅니다. 슬라이드
  구조보다 여러 포맷의 처리 속도가 중요하다면 실제로 큰 장점입니다.
- 각 검사는 작은 합성 덱의 성질 하나만 봅니다. 실제 덱은 모든 요소가 섞여
  있으므로, 이 결과는 기능 지원 여부이지 실제 코퍼스에서의 출력 품질이 아닙니다.
- 검사는 올바른 결과의 여러 표기를 인정합니다(`**x**`와 `__x__`, `mc^{2}`와
  `m{c}^{2}`, 슬라이드 표시 방식 등). 변환기가 정보를 보존했는데 검사가 놓친다면
  검사의 버그이니 알려 주세요.
