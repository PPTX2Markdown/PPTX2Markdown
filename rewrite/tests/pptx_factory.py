"""Minimal synthetic PPTX packages used only by rewrite tests."""

from __future__ import annotations

import zipfile
from pathlib import Path

CONTENT_TYPES = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="mp3" ContentType="audio/mpeg"/>
  <Override PartName="/ppt/presentation.xml"
    ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml"
    ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/notesSlides/notesSlide1.xml"
    ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>
</Types>
"""

ROOT_RELS = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="ppt/presentation.xml"/>
</Relationships>
"""

PRESENTATION = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId1"/>
  </p:sldIdLst>
  <p:sldSz cx="9144000" cy="6858000"/>
</p:presentation>
"""

PRESENTATION_RELS = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
    Target="slides/slide1.xml"/>
</Relationships>
"""

SLIDE = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name="Group"/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr/>

      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Title"/>
          <p:cNvSpPr/>
          <p:nvPr><p:ph type="title"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="0" y="0"/><a:ext cx="9000000" cy="800000"/></a:xfrm>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/><a:lstStyle/>
          <a:p><a:r><a:t>Quarterly Report</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>

      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="3" name="Body"/>
          <p:cNvSpPr/><p:nvPr><p:ph type="body"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="0" y="1000000"/><a:ext cx="4000000" cy="1800000"/></a:xfrm>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/><a:lstStyle/>
          <a:p>
            <a:r><a:rPr><a:hlinkClick r:id="rIdLink"/></a:rPr><a:t>OpenAI</a:t></a:r>
            <a:r><a:t> reference</a:t></a:r>
          </a:p>
          <a:p><a:pPr><a:buChar char="•"/></a:pPr><a:r><a:t>Revenue grew</a:t></a:r></a:p>
          <a:p><a:pPr lvl="1"><a:buChar char="•"/></a:pPr><a:r><a:t>APAC led</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>

      <p:pic>
        <p:nvPicPr>
          <p:cNvPr id="4" name="Picture 1" descr="Sales chart"/>
          <p:cNvPicPr/><p:nvPr/>
        </p:nvPicPr>
        <p:blipFill>
          <a:blip r:embed="rIdImage"/>
          <a:stretch><a:fillRect/></a:stretch>
        </p:blipFill>
        <p:spPr>
          <a:xfrm><a:off x="5000000" y="1000000"/><a:ext cx="3000000" cy="1800000"/></a:xfrm>
        </p:spPr>
        <a:audioFile r:link="rIdAudio"/>
      </p:pic>

      <p:graphicFrame>
        <p:nvGraphicFramePr>
          <p:cNvPr id="5" name="Table 1"/>
          <p:cNvGraphicFramePr/><p:nvPr/>
        </p:nvGraphicFramePr>
        <p:xfrm><a:off x="0" y="3200000"/><a:ext cx="8000000" cy="1800000"/></p:xfrm>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
            <a:tbl>
              <a:tblPr/><a:tblGrid><a:gridCol w="4000000"/><a:gridCol w="4000000"/></a:tblGrid>
              <a:tr h="600000">
                <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Region</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Revenue</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
              </a:tr>
              <a:tr h="600000">
                <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>APAC</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>12.4M</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
              </a:tr>
            </a:tbl>
          </a:graphicData>
        </a:graphic>
      </p:graphicFrame>
    </p:spTree>
  </p:cSld>
</p:sld>
"""

SLIDE_RELS = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdImage"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    Target="../media/image1.png"/>
  <Relationship Id="rIdAudio"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio"
    Target="../media/media1.mp3"/>
  <Relationship Id="rIdNotes"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"
    Target="../notesSlides/notesSlide1.xml"/>
  <Relationship Id="rIdLink"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
    Target="https://openai.com" TargetMode="External"/>
</Relationships>
"""

NOTES = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld><p:spTree>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="Notes"/><p:cNvSpPr/><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr>
      <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Presenter note</a:t></a:r></a:p></p:txBody>
    </p:sp>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="3" name="Slide number"/><p:cNvSpPr/><p:nvPr><p:ph type="sldNum"/></p:nvPr></p:nvSpPr>
      <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>1</a:t></a:r></a:p></p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
</p:notes>
"""


TEMPLATED_CONTENT_TYPES = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml"
    ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml"
    ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml"
    ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml"
    ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
</Types>
"""

_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _relationships(*entries: str) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + "".join(entries)
        + "</Relationships>"
    )


def build_templated_pptx(
    path: Path,
    *,
    slide_xml: str,
    layout_xml: str,
    master_xml: str,
    layout_relationships: str = "",
    master_relationships: str = "",
    extra_parts: dict[str, str | bytes] | None = None,
) -> Path:
    """One slide inheriting from one layout and one master."""

    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", TEMPLATED_CONTENT_TYPES)
        archive.writestr("_rels/.rels", ROOT_RELS)
        archive.writestr("ppt/presentation.xml", PRESENTATION)
        archive.writestr("ppt/_rels/presentation.xml.rels", PRESENTATION_RELS)
        archive.writestr("ppt/slides/slide1.xml", slide_xml)
        archive.writestr(
            "ppt/slides/_rels/slide1.xml.rels",
            _relationships(
                f'<Relationship Id="rId1" Type="{_REL}/slideLayout"'
                ' Target="../slideLayouts/slideLayout1.xml"/>'
            ),
        )
        archive.writestr("ppt/slideLayouts/slideLayout1.xml", layout_xml)
        archive.writestr(
            "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
            _relationships(
                f'<Relationship Id="rId1" Type="{_REL}/slideMaster"'
                ' Target="../slideMasters/slideMaster1.xml"/>',
                layout_relationships,
            ),
        )
        archive.writestr("ppt/slideMasters/slideMaster1.xml", master_xml)
        if master_relationships:
            archive.writestr(
                "ppt/slideMasters/_rels/slideMaster1.xml.rels",
                _relationships(master_relationships),
            )
        for name, data in (extra_parts or {}).items():
            archive.writestr(name, data)
    return path


def build_test_pptx(
    path: Path,
    slide_xml: str = SLIDE,
    *,
    extra_slide_relationships: str = "",
    extra_parts: dict[str, str | bytes] | None = None,
) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    slide_relationships = SLIDE_RELS.replace(
        "</Relationships>", f"{extra_slide_relationships}</Relationships>"
    )
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", CONTENT_TYPES)
        archive.writestr("_rels/.rels", ROOT_RELS)
        archive.writestr("ppt/presentation.xml", PRESENTATION)
        archive.writestr("ppt/_rels/presentation.xml.rels", PRESENTATION_RELS)
        archive.writestr("ppt/slides/slide1.xml", slide_xml)
        archive.writestr("ppt/slides/_rels/slide1.xml.rels", slide_relationships)
        for name, data in (extra_parts or {}).items():
            archive.writestr(name, data)
        archive.writestr("ppt/notesSlides/notesSlide1.xml", NOTES)
        archive.writestr("ppt/media/image1.png", b"\x89PNG\r\n\x1a\nsynthetic")
        archive.writestr("ppt/media/media1.mp3", b"ID3synthetic-audio")
    return path
