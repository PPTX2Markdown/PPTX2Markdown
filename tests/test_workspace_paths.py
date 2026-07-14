from __future__ import annotations

import argparse
import tempfile
import unittest
import zipfile
from pathlib import Path

from pptx2markdown.main_converter.package_inputs import (
    default_ppt_conversion_cache_dir,
    default_pptx_input_dir,
    default_target_dir,
)
from pptx2markdown.main_converter.run_pptx_to_markdown import _build_config
from pptx2markdown.table_pipeline.run import (
    _collect_default_pptx_inputs,
    run_pipeline,
)
from pptx2markdown.workspace_paths import WorkspacePaths


class WorkspacePathTests(unittest.TestCase):
    def test_default_layout_has_one_output_root_and_one_work_root(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            project = Path(temp_dir)
            paths = WorkspacePaths.from_base(project)

            project = project.resolve()
            self.assertEqual(paths.output_dir, project / "output")
            self.assertEqual(paths.work_dir, project / ".pptx2markdown")
            self.assertEqual(paths.target_slides, paths.work_dir / "target_slides")
            self.assertEqual(paths.target_pptx, paths.work_dir / "target_pptx")
            self.assertEqual(
                paths.structure_analysis,
                paths.work_dir / "structure_analysis",
            )
            self.assertEqual(paths.table_pipeline, paths.work_dir / "table_pipeline")
            self.assertFalse(paths.output_dir.exists())
            self.assertFalse(paths.work_dir.exists())

    def test_main_converter_and_pipeline_helpers_share_the_same_paths(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            project = Path(temp_dir)
            work_dir = project / "custom-work"
            output_dir = project / "custom-output"
            config = _build_config(
                argparse.Namespace(
                    inputs=[],
                    output_dir=str(output_dir),
                    work_dir=str(work_dir),
                    output_format="markdown",
                    headings="auto",
                    pptx_inheritance="style",
                    inherited_shapes="visible",
                    ppt_converter="auto",
                )
            )
            paths = WorkspacePaths.from_base(
                project,
                work_dir=work_dir,
                output_dir=output_dir,
            )

            self.assertEqual(config.cwd, paths.work_dir)
            self.assertEqual(config.output_dir, paths.output_dir)
            self.assertFalse(config.convert_vector_images)
            self.assertEqual(default_target_dir(config.cwd), paths.target_slides)
            self.assertEqual(default_pptx_input_dir(config.cwd), paths.target_pptx)
            self.assertEqual(
                default_ppt_conversion_cache_dir(config.cwd),
                paths.ppt_conversion_cache,
            )

    def test_input_lookup_does_not_create_empty_directories(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            work_dir = Path(temp_dir) / "work"

            self.assertEqual(_collect_default_pptx_inputs(work_dir), [])
            self.assertFalse(work_dir.exists())

    def test_table_pipeline_writes_only_below_shared_work_directory(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            project = Path(temp_dir)
            work_dir = project / "work"
            source = project / "sample.pptx"
            with zipfile.ZipFile(source, "w") as archive:
                archive.writestr(
                    "ppt/slides/slide1.xml",
                    '<p:sld xmlns:p="http://schemas.openxmlformats.org/'
                    'presentationml/2006/main"/>',
                )

            self.assertEqual(run_pipeline([source], work_dir=work_dir), 0)

            paths = WorkspacePaths.from_base(project, work_dir=work_dir)
            self.assertTrue(paths.target_slides.is_dir())
            self.assertTrue(paths.table_extract_results.is_dir())
            self.assertTrue(paths.table_parsing_results.is_dir())
            self.assertTrue(paths.table_markdown_results.is_dir())
            self.assertFalse((project / "target_slides").exists())
            self.assertFalse((project / "artifacts").exists())


if __name__ == "__main__":
    unittest.main()
