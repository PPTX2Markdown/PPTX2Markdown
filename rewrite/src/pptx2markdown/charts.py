"""Chart parts converted to Markdown through chart2md."""

from __future__ import annotations

import xml.etree.ElementTree as ET
from dataclasses import dataclass, field

from chart2md import convert_chart

from .errors import PackageSecurityError
from .ooxml import OOXMLPackage, validate_embedded_package

_WORKBOOK_SUFFIXES = (".xlsx", ".xlsm")


@dataclass
class _Rejections:
    messages: list[str] = field(default_factory=list)


class _PackageContext:
    """The part of chart2md's ``ZipContext`` interface it uses, read through the package.

    Every read goes through :class:`OOXMLPackage`, so chart2md sees only parts that passed
    the package limits, relationships that stay inside the package and XML without DTDs.
    Embedded workbooks are ZIP files of their own and are checked before chart2md opens them.
    """

    def __init__(self, package: OOXMLPackage, path: str, rejections: _Rejections) -> None:
        self._package = package
        self.path = path
        self._rejections = rejections

    def resolve(self, rid: str) -> _PackageContext | None:
        target = self._package.related_part(self.path, rid)
        if target is None:
            return None
        return _PackageContext(self._package, target, self._rejections)

    def resolve_all(self) -> dict[str, _PackageContext]:
        resolved: dict[str, _PackageContext] = {}
        for rid in self._package.relationships(self.path):
            context = self.resolve(rid)
            if context is not None:
                resolved[rid] = context
        return resolved

    def read_bytes(self) -> bytes:
        data = self._package.read(self.path)
        if self.path.casefold().endswith(_WORKBOOK_SUFFIXES):
            try:
                validate_embedded_package(data, self._package.limits, self.path)
            except PackageSecurityError as exc:
                # chart2md falls back to the chart's cached XML values.
                self._rejections.messages.append(str(exc))
                raise
        return data

    def parse_xml(self) -> ET.Element:
        return self._package.xml(self.path)


@dataclass(frozen=True)
class ChartConversion:
    markdown: str
    rejected_workbooks: tuple[str, ...] = ()


def chart_markdown(package: OOXMLPackage, part_name: str) -> ChartConversion:
    """Convert a DrawingML (``c:``) or chartEx (``cx:``) chart part to Markdown.

    Like the legacy converter, DrawingML charts read the cached XML values and chartEx
    charts prefer their embedded workbook.
    """

    rejections = _Rejections()
    context = _PackageContext(package, part_name, rejections)
    markdown = convert_chart(context.parse_xml(), context)  # type: ignore[arg-type]
    return ChartConversion(
        markdown=markdown.strip(),
        rejected_workbooks=tuple(rejections.messages),
    )
