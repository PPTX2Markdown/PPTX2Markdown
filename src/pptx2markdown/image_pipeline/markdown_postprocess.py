"""Markdown cleanup helpers shared across providers."""

from __future__ import annotations

import re

_PROMPT_ECHO_PHRASES = {
    "return markdown only.",
    "extract only document-worthy information from this image as concise markdown for retrieval.",
    "rules:",
    (
        "1. keep only information that is useful for search, retrieval, "
        "or understanding the document."
    ),
    "2. omit unreadable, uncertain, or purely decorative content.",
    "3. do not invent missing text or details.",
    "4. do not wrap the answer in triple backticks.",
    "output policy by image type:",
    "if the image is mainly a table, recreate it as a markdown table.",
    "preserve readable headers, row labels, units, and cell values.",
    "omit cells that are unreadable instead of guessing.",
    "if the image contains formulas or equations:",
    "extract them faithfully.",
    "use latex-style math notation when possible.",
    "keep variable names, subscripts, superscripts, and operators.",
    "if the image is mainly plaintext:",
    "transcribe the readable text in clean markdown.",
    "preserve headings, bullet points, and short paragraph structure when visible.",
    "if the image is an informative figure, chart, diagram, screenshot, or illustration:",
    "first write exactly one concise sentence summarizing what information the image conveys.",
    (
        "then list any clearly readable labels, legends, axis names, key values, "
        "or embedded text in markdown bullets."
    ),
    "if the image is decorative, redundant, or not useful for document retrieval:",
    "answer exactly: 불필요한 정보",
    "additional constraints:",
    "prefer faithful extraction over fluent rewriting.",
    (
        "keep the output concise, but do not drop important entities, numbers, "
        "labels, or relationships."
    ),
    "preserve technical terms as written in the image.",
    "if only part of the image is readable, extract only that readable part.",
    "do not wrap the answer in triple backticks.",
    (
        "이미지에서 문서 검색, 검색 증강, 문서 이해에 유용한 정보만 "
        "간결한 한국어 markdown으로 추출하세요."
    ),
    "규칙:",
    "1. 답변은 반드시 한국어로 작성하세요.",
    "2. 검색, 검색 증강, 문서 이해에 유용한 정보만 남기세요.",
    "3. 읽을 수 없거나 불확실하거나 순수하게 장식적인 내용은 생략하세요.",
    "4. 보이지 않는 텍스트나 세부사항을 추측해서 만들지 마세요.",
    "5. 답변을 삼중 백틱 코드 블록으로 감싸지 마세요.",
    "이미지 유형별 출력 정책:",
    "이미지가 주로 표인 경우:",
    "markdown 표로 재구성하세요.",
    "읽을 수 있는 헤더, 행 라벨, 단위, 셀 값을 보존하세요.",
    "읽을 수 없는 셀은 추측하지 말고 생략하세요.",
    "이미지에 수식이나 방정식이 포함된 경우:",
    "충실하게 추출하세요.",
    "가능하면 latex 스타일 수학 표기법을 사용하세요.",
    "변수명, 아래첨자, 위첨자, 연산자를 보존하세요.",
    "이미지가 주로 일반 텍스트인 경우:",
    "읽을 수 있는 텍스트를 깔끔한 markdown으로 전사하세요.",
    "보이는 제목, 불릿, 짧은 문단 구조를 보존하세요.",
    "이미지가 의미 있는 그림, 차트, 다이어그램, 스크린샷, 일러스트인 경우:",
    "먼저 이미지가 전달하는 정보를 한국어 한 문장으로 간결하게 요약하세요.",
    (
        "이어서 읽을 수 있는 라벨, 범례, 축 이름, 주요 값, 포함된 텍스트를 "
        "markdown 불릿으로 나열하세요."
    ),
    "이미지가 장식용이거나 중복되거나 문서 검색에 유용하지 않은 경우:",
    "정확히 다음과 같이 답하세요: 불필요한 정보",
    "추가 제약:",
    "자연스러운 재작성보다 충실한 추출을 우선하세요.",
    "간결하게 작성하되 중요한 개체, 숫자, 라벨, 관계는 누락하지 마세요.",
    "이미지 안의 기술 용어, 고유명사, 제품명, 변수명은 원문 표기를 보존하세요.",
    "일부만 읽을 수 있다면 읽을 수 있는 부분만 추출하세요.",
}


def _unwrap_latex_text_line(line: str) -> str:
    match = re.fullmatch(r"\$\\text\s*\{\s*(.*?)\s*\}\$", line.strip())
    if match:
        return match.group(1).strip()
    return line


def _is_prompt_echo_line(line: str) -> bool:
    normalized = _unwrap_latex_text_line(line).strip()
    if not normalized:
        return False
    normalized = re.sub(r"^[\-\*\u2022]\s*", "", normalized)
    normalized = re.sub(r"\s+", " ", normalized).strip().lower()
    return normalized in _PROMPT_ECHO_PHRASES


def normalize_markdown(text: str) -> str:
    normalized = text.strip()
    fence_match = re.fullmatch(
        r"```(?:markdown|md)?\s*(.*?)```", normalized, flags=re.DOTALL | re.IGNORECASE
    )
    if fence_match:
        normalized = fence_match.group(1).strip()
    kept_lines = [
        line.rstrip() for line in normalized.splitlines() if not _is_prompt_echo_line(line)
    ]
    normalized = "\n".join(line for line in kept_lines).strip()
    return normalized.rstrip() + "\n" if normalized else ""
