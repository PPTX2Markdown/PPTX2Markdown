"""Shared constants for the image-to-markdown pipeline."""

from __future__ import annotations


DEFAULT_PROVIDER = "local"
DEFAULT_GEMINI_MODEL = "gemini-2.5-flash"
DEFAULT_GEMINI_API_KEY_ENV = "GEMINI_API_KEY"
DEFAULT_OPENAI_MODEL = "gpt-4.1-mini"
DEFAULT_OPENAI_API_KEY_ENV = "OPENAI_API_KEY"
DEFAULT_GEMINI_MAX_RETRIES = 2
DEFAULT_GEMINI_BASE_BACKOFF_SEC = 2.0
DEFAULT_GEMINI_MAX_BACKOFF_SEC = 30.0
DEFAULT_GEMINI_MIN_REQUEST_INTERVAL_SEC = 1.5

QWEN_VL_MODELS = {
    "3b": "Qwen/Qwen2.5-VL-3B-Instruct",
    "7b": "Qwen/Qwen2.5-VL-7B-Instruct",
}

DEFAULT_PROMPT = (
    "이미지에서 문서 검색, 검색 증강, 문서 이해에 유용한 정보만 간결한 한국어 Markdown으로 추출하세요.\n\n"
    "규칙:\n"
    "1. 답변은 반드시 한국어로 작성하세요.\n"
    "2. 검색, 검색 증강, 문서 이해에 유용한 정보만 남기세요.\n"
    "3. 읽을 수 없거나 불확실하거나 순수하게 장식적인 내용은 생략하세요.\n"
    "4. 보이지 않는 텍스트나 세부사항을 추측해서 만들지 마세요.\n"
    "5. 답변을 삼중 백틱 코드 블록으로 감싸지 마세요.\n\n"
    "이미지 유형별 출력 정책:\n"
    "- 이미지가 주로 표인 경우:\n"
    "  - Markdown 표로 재구성하세요.\n"
    "  - 읽을 수 있는 헤더, 행 라벨, 단위, 셀 값을 보존하세요.\n"
    "  - 읽을 수 없는 셀은 추측하지 말고 생략하세요.\n"
    "- 이미지에 수식이나 방정식이 포함된 경우:\n"
    "  - 충실하게 추출하세요.\n"
    "  - 가능하면 LaTeX 스타일 수학 표기법을 사용하세요.\n"
    "  - 변수명, 아래첨자, 위첨자, 연산자를 보존하세요.\n"
    "- 이미지가 주로 일반 텍스트인 경우:\n"
    "  - 읽을 수 있는 텍스트를 깔끔한 Markdown으로 전사하세요.\n"
    "  - 보이는 제목, 불릿, 짧은 문단 구조를 보존하세요.\n"
    "- 이미지가 의미 있는 그림, 차트, 다이어그램, 스크린샷, 일러스트인 경우:\n"
    "  - 먼저 이미지가 전달하는 정보를 한국어 한 문장으로 간결하게 요약하세요.\n"
    "  - 이어서 읽을 수 있는 라벨, 범례, 축 이름, 주요 값, 포함된 텍스트를 Markdown 불릿으로 나열하세요.\n"
    "- 이미지가 장식용이거나 중복되거나 문서 검색에 유용하지 않은 경우:\n"
    "  - 정확히 다음과 같이 답하세요: 불필요한 정보\n\n"
    "추가 제약:\n"
    "- 자연스러운 재작성보다 충실한 추출을 우선하세요.\n"
    "- 간결하게 작성하되 중요한 개체, 숫자, 라벨, 관계는 누락하지 마세요.\n"
    "- 이미지 안의 기술 용어, 고유명사, 제품명, 변수명은 원문 표기를 보존하세요.\n"
    "- 일부만 읽을 수 있다면 읽을 수 있는 부분만 추출하세요."
)

DEFAULT_MAX_NEW_TOKENS = 1024
DEFAULT_TRANSPARENT_BG_GRAY = 192
