# Golden output review

English | [한국어](REVIEW.ko.md)

This review classifies the checked-in parser results rather than treating every
successful conversion as correct. It should be revisited whenever snapshots are
accepted.

## Confirmed expected results

- Complex image/text layouts retain geometric order and local image links.
- Empty valid presentations produce an empty slide list and zero-byte Markdown.
- SVG fallbacks, audio/video posters, 3D previews, and their source payloads are
  emitted with stable relative paths.
- OLE and Package payloads are unpacked to their original Office, JSON, audio,
  video, or GLB types when recoverable.
- Native charts, tables, SmartArt, hyperlinks, lists, grouped shapes, rotations,
  CJK text, and ordinary numeric labels remain searchable content.
- Date and slide-number fields are excluded without dropping neighboring text.
- Speaker notes remain outside visual blocks and are rendered only under the
  explicit `[Speaker_Notes]` section.

## Intentional policy decisions frozen by the suite

- Reading order is geometric: a recursive XY-cut over shapes, with groups and
  enclosing frames read as units and slide-covering backgrounds ignored. A slide
  whose layout can be read more than one way keeps its candidate orders for
  review; a recipe `order` rule overrides it.
- A 1 px column-boundary overlap is tolerated as a column cut, while the 24 px
  overlap fixture falls back to row-major order.
- A positive 0.5 px gutter is sufficient to separate columns.
- Small 1–3 px alignment jitter and different text-box heights do not require a
  separate row-clustering pass.
- Full-width regions remain before and after their enclosed grid. Within the
  tightly aligned grid fixture, horizontal cuts yield row-major order.
- Merged cells repeat the merged value in every cell they cover, in the header
  and the body, horizontally and vertically (`--merged-cells fill`, the default),
  so each row reads on its own; `--merged-cells blank` keeps the value only in
  the cell the merge starts from. A body row merged across the whole table is
  written once.
- Unicode content is preserved as authored. NFC/NFD and full-width variants are
  not rewritten in the user-visible output.
- Unlabeled SmartArt is represented by a stable node-count placeholder rather
  than discarded or fabricated into labeled content.

## Defects found and resolved while building the suite

- Package-absolute notes relationships such as
  `/ppt/notesSlides/notesSlide6.xml` were previously resolved as filesystem-root
  paths. They now resolve from the extracted OOXML package root.
- Office lock files matching `~$*.pptx` are skipped with a warning.
- The snapshots were regenerated for the 0.2.0 engine; the block counts per kind
  match the 0.1 snapshots they replaced.

## Open quality backlog

No checked-in result is currently classified as a known parser defect. Future
work should begin by adding a minimal PPTX and reviewed snapshot for a newly
observed failure, then changing the parser. Broad corpus growth without a new
policy boundary or failure mode is lower value than maintaining these explicit
expectations.
