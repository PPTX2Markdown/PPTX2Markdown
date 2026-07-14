from __future__ import annotations

import re
import unittest
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
AUTHORED_DOCUMENT_DIRS = (
    REPO_ROOT,
    REPO_ROOT / ".github",
    REPO_ROOT / "docs",
    REPO_ROOT / "tests" / "fixtures" / "golden",
)
MARKDOWN_LINK = re.compile(r"\[[^\]]+\]\(([^)]+)\)")
FENCED_CODE = re.compile(r"```.*?```", re.DOTALL)
INLINE_CODE = re.compile(r"`[^`\n]*`")
PUBLIC_FLAGS = (
    "--output-dir",
    "--work-dir",
    "--output-format",
    "--headings",
    "--placeholder-inheritance",
    "--inherited-shapes",
    "--ppt-converter",
    "--verbose",
)


def authored_documents() -> list[Path]:
    documents: list[Path] = []
    for directory in AUTHORED_DOCUMENT_DIRS:
        documents.extend(path for path in directory.glob("*.md") if path.is_file())
    return sorted(set(documents))


def english_path(path: Path) -> Path:
    if path.name.endswith(".ko.md"):
        return path.with_name(path.name.removesuffix(".ko.md") + ".md")
    return path


def korean_path(path: Path) -> Path:
    return path.with_name(path.stem + ".ko.md")


class DocumentationPolicyTests(unittest.TestCase):
    def test_copyright_owner_is_present_in_notice_and_package_metadata(self) -> None:
        notice = (REPO_ROOT / "NOTICE").read_text(encoding="utf-8")
        self.assertIn("HANKOOK TIRE & TECHNOLOGY CO., LTD.", notice)

        pyproject = (REPO_ROOT / "pyproject.toml").read_text(encoding="utf-8")
        self.assertIn('authors = [{ name = "HANKOOK TIRE & TECHNOLOGY CO., LTD." }]', pyproject)

    def test_every_authored_document_has_an_english_and_korean_pair(self) -> None:
        documents = authored_documents()
        self.assertTrue(documents)
        for path in documents:
            with self.subTest(path=path.relative_to(REPO_ROOT)):
                english = english_path(path)
                korean = korean_path(english)
                self.assertTrue(english.is_file(), f"missing English document: {english}")
                self.assertTrue(korean.is_file(), f"missing Korean document: {korean}")
                self.assertIn("[한국어](", english.read_text(encoding="utf-8"))
                self.assertIn("[English](", korean.read_text(encoding="utf-8"))

    def test_local_markdown_links_resolve(self) -> None:
        for path in authored_documents():
            text = path.read_text(encoding="utf-8")
            prose = INLINE_CODE.sub("", FENCED_CODE.sub("", text))
            for raw_target in MARKDOWN_LINK.findall(prose):
                target = raw_target.strip().split("#", 1)[0]
                if not target or target.startswith(("http://", "https://", "mailto:")):
                    continue
                resolved = (path.parent / target).resolve()
                with self.subTest(path=path.relative_to(REPO_ROOT), target=target):
                    self.assertTrue(resolved.exists(), f"broken local link: {target}")

    def test_root_readmes_document_the_public_cli(self) -> None:
        for relative_path in ("README.md", "README.ko.md"):
            text = (REPO_ROOT / relative_path).read_text(encoding="utf-8")
            for flag in PUBLIC_FLAGS:
                with self.subTest(path=relative_path, flag=flag):
                    self.assertIn(flag, text)


if __name__ == "__main__":
    unittest.main()
