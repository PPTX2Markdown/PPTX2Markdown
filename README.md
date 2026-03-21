# pptx2markdown

PPTX를 Markdown으로 변환하는 파이프라인입니다.

## 핵심 모듈

- `main_converter`: 메인 엔드투엔드 실행기
- `surya_pipeline`: Surya reading-order 재정렬 파이프라인
- `structure_analyzer`: XML 기반 reading-order 분석기
- `table_extractor`, `table_parser`: 테이블 추출/파싱/렌더링
- `image_pipeline`: 이미지 테이블 분류 + Surya 기반 표 인식 보조

## 현재 실행 정책 (중요)
- 핵심 실행 스크립트는 `convert_slides_to_md.py`입니다.
- 입력은 .pptx 파일이며, 출력은 `output/<reading-order>/` 하위에 생성됩니다.
  - <reding-order> : xml | surya
- `--reading-order surya`를 사용하면 Surya 기반 순서 재정렬 결과를 반영합니다

간단한 명령어 사용 규칙은 다음과 같습니다.
```bash
python3 convert_slides [--per-slide] [--reading-order {xml(default) | surya}] [--strict] [INPUT_PPTX ...]
``` 
- `[INPUT_PPTX]`
  - *.pptx 파일을 0개 이상 전달할 수 있습니다.
  - 예: `sample3.pptx sample4.pptx`
  - 생략하면 `target_pptx/*.pptx` 전체를 자동 처리합니다.
- `--per-slide`
  - 패키지 통합 결과(result.md)와 함께 슬라이드별 Markdown(`per_slide/slideN.md`)을 추가 생성합니다.
- `--reading-order`
  - 읽기 순서 전략을 선택합니다.
  - xml(기본): XML 기반 순서 사용
  - surya: Surya 파이프라인 결과(`slideN.reordered.xml`) 사용
- `--strict`
  - XML 모드에서 stricter heading 규칙을 적용합니다.
  - surya 모드에서는 무시됩니다.

## 빠른 사용법

저장소 루트 기준:

기본 실행(default):

```bash
python3 main_converter/convert_slides_to_md.py
```
- 이 동작은 `main_converter/target_pptx/*` 의 모든 `.pptx` 파일을 대상으로 `--reading-order xml` 방식으로 수행
- 즉, 다음 동작과 같음

```bash
python3 main_converter/convert_slides_to_md.py --reading-order xml [none_target == target/pptx/* ]
```


특정 PPTX만 변환 수행:
```bash
python3 main_converter/convert_slides_to_md.py [*targets]
```

예시(Example) 
```bash
python3 main_converter/convert_slides_to_md.py sample3.pptx sample4.pptx
```

Surya reading order:

```bash
python3 main_converter/convert_slides_to_md.py --reading-order surya sample3.pptx sample4.pptx
```

슬라이드별 파일도 생성:

```bash
python3 main_converter/convert_slides_to_md.py --per-slide
```

## 출력 경로

- `main_converter/output/<reading-order>/<package>/result.md`
- `main_converter/output/<reading-order>/convert_manifest.json`
- `main_converter/output/<reading-order>/<package>/media/*`
- (옵션) `main_converter/output/<reading-order>/<package>/per_slide/slideN.md`
<reding-order> : xml | surya

## Surya 연동 요약

`--reading-order surya` 실행 시:

1. `main_converter/target_slides`를 Surya 슬라이드 루트로 사용
2. `surya_pipeline/run_surya_pipeline.py` 실행
3. `surya_pipeline/output/structure_ready/<package>/slideN.reordered.xml` 사용

자세한 내용은 [surya_pipeline/README.md](surya_pipeline/README.md),  
메인 옵션은 [main_converter/README.md](main_converter/README.md)를 참고하세요.
