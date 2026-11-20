"""Runtime owner for a canonical document and its managed asset payloads."""

from __future__ import annotations

from pathlib import Path
from typing import Literal

from .assets import AssetStore
from .errors import AssetUnavailableError
from .model import PresentationDocument
from .render import AssetResolver, render_presentation_json, render_presentation_markdown


class ParsedPresentation:
    """A parsed document whose asset payloads remain valid until close()."""

    def __init__(
        self,
        *,
        document: PresentationDocument,
        assets: AssetStore,
        source_path: Path,
    ) -> None:
        descriptor_ids = tuple(asset.id for asset in assets.descriptors())
        document_ids = tuple(asset.id for asset in document.assets)
        if descriptor_ids != document_ids:
            assets.close()
            raise ValueError("document asset descriptors do not match the owned asset store")
        self.document = document
        self.assets = assets
        self.source_path = source_path
        self._closed = False

    @property
    def closed(self) -> bool:
        return self._closed

    @property
    def slides(self):
        """Convenience view of document.slides."""

        return self.document.slides

    def _ensure_open(self) -> None:
        if self._closed:
            raise AssetUnavailableError("parsed presentation is closed")

    def render_markdown(
        self,
        *,
        asset_url: AssetResolver | None = None,
        assets: Literal["error", "omit"] = "error",
    ) -> str:
        """Render Markdown with explicit asset URL behavior."""

        self._ensure_open()
        return render_presentation_markdown(
            self.document,
            asset_resolver=asset_url,
            asset_mode=assets,
        )

    def render_json(self) -> str:
        """Render canonical JSON; asset bytes are intentionally excluded."""

        return render_presentation_json(self.document)

    def export_asset(self, asset_id: str, directory: str | Path) -> Path:
        """Materialize one asset into an explicit durable directory."""

        self._ensure_open()
        descriptor = self.assets.descriptor(asset_id)
        destination = Path(directory).expanduser().resolve() / descriptor.original_name
        if destination.exists():
            raise FileExistsError(destination)
        self.assets.materialize(asset_id, destination)
        return destination

    def save(
        self,
        bundle_dir: str | Path,
        *,
        output_format: Literal["markdown", "json"] = "markdown",
        assets: Literal["copy", "omit"] = "copy",
        overwrite: Literal["error", "replace"] = "error",
    ):
        """Write a portable bundle using the atomic writer."""

        self._ensure_open()
        from .writer import SaveOptions, save_bundle

        return save_bundle(
            self,
            bundle_dir,
            options=SaveOptions(
                output_format=output_format,
                assets=assets,
                overwrite=overwrite,
            ),
        )

    def close(self) -> None:
        if self._closed:
            return
        self.assets.close()
        self._closed = True

    def __enter__(self) -> "ParsedPresentation":
        self._ensure_open()
        return self

    def __exit__(self, *_: object) -> None:
        self.close()

    def __del__(self) -> None:
        try:
            self.close()
        except Exception:
            pass
