"""pptx2markdown — convert PPTX/PPT presentations to Markdown."""

from importlib.metadata import PackageNotFoundError
from importlib.metadata import version as _pkg_version

try:
    __version__ = _pkg_version("pptx2markdown")
except PackageNotFoundError:  # 개발 트리에서 미설치 상태로 import된 경우
    __version__ = "0.0.0"

from pptx2markdown.api import convert

__all__ = ["convert", "__version__"]
