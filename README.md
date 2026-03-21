# pptx2markdown

PPTX를 Markdown으로 변환하는 파이프라인입니다.

## 핵심 모듈

- `main_converter`: 메인 엔드투엔드 실행기
- `surya_pipeline`: Surya reading-order 재정렬 파이프라인
- `structure_analyzer`: XML 기반 reading-order 분석기
- `table_extractor`, `table_parser`: 테이블 추출/파싱/렌더링
- `image_pipeline`: 이미지 테이블 분류 + Surya 기반 표 인식 보조

## 현재 실행 정책 (중요)

`main_converter`는 `.pptx` 입력 자동 추출을 지원합니다.

- 인자에 `.pptx`를 주면 자동 추출 후 변환
- 인자가 없으면 `target_pptx/*.pptx` 전체를 자동 추출/변환

## 빠른 사용법

저장소 루트 기준:

기본 실행:

```bash
python3 main_converter/convert_slides_to_md.py
```

특정 PPTX:

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

## Surya 연동 요약

`--reading-order surya` 실행 시:

1. `main_converter/target_slides`를 Surya 슬라이드 루트로 사용
2. `surya_pipeline/run_surya_pipeline.py` 실행
3. `surya_pipeline/output/structure_ready/<package>/slideN.reordered.xml` 사용

자세한 내용은 [surya_pipeline/README.md](surya_pipeline/README.md),  
메인 옵션은 [main_converter/README.md](main_converter/README.md)를 참고하세요.
