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
    "Extract only document-worthy information from this image as concise Markdown for retrieval.\n\n"
    "Rules:\n"
    "1. Keep only information that is useful for search, retrieval, or understanding the document.\n"
    "2. Omit unreadable, uncertain, or purely decorative content.\n"
    "3. Do not invent missing text or details.\n"
    "4. Do not wrap the answer in triple backticks.\n\n"
    "Output policy by image type:\n"
    "- If the image is mainly a table:\n"
    "  - Recreate the table as a Markdown table.\n"
    "  - Preserve readable headers, row labels, units, and cell values.\n"
    "  - Omit cells that are unreadable instead of guessing.\n"
    "- If the image contains formulas or equations:\n"
    "  - Extract them faithfully.\n"
    "  - Use LaTeX-style math notation when possible.\n"
    "  - Keep variable names, subscripts, superscripts, and operators.\n"
    "- If the image is mainly plaintext:\n"
    "  - Transcribe the readable text in clean Markdown.\n"
    "  - Preserve headings, bullet points, and short paragraph structure when visible.\n"
    "- If the image is an informative figure, chart, diagram, screenshot, or illustration:\n"
    "  - First write exactly one concise sentence summarizing what information the image conveys.\n"
    "  - Then list any clearly readable labels, legends, axis names, key values, or embedded text in Markdown bullets.\n"
    "- If the image is decorative, redundant, or not useful for document retrieval:\n"
    "  - Answer exactly: 불필요한 정보\n\n"
    "Additional constraints:\n"
    "- Prefer faithful extraction over fluent rewriting.\n"
    "- Keep the output concise, but do not drop important entities, numbers, labels, or relationships.\n"
    "- Preserve technical terms as written in the image.\n"
    "- If only part of the image is readable, extract only that readable part."
)

DEFAULT_MAX_NEW_TOKENS = 1024
DEFAULT_TRANSPARENT_BG_GRAY = 192
