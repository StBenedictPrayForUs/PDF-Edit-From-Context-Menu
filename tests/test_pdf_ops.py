import tempfile
import unittest
from pathlib import Path

import fitz
from PySide6.QtGui import QColor, QImage

from app import launcher, pdf_ops


def _create_pdf(path: Path, labels: list[str]) -> None:
    doc = fitz.open()
    try:
        for label in labels:
            page = doc.new_page(width=300, height=400)
            page.insert_text((72, 72), label, fontsize=18)
        doc.save(str(path))
    finally:
        doc.close()


def _create_image(path: Path, width: int, height: int, color_name: str) -> None:
    image = QImage(width, height, QImage.Format_RGB32)
    image.fill(QColor(color_name))
    if not image.save(str(path)):
        raise RuntimeError(f"Failed to save image fixture: {path}")


class PdfOpsTests(unittest.TestCase):
    def test_compute_sections_returns_empty_for_zero_page_pdf(self) -> None:
        self.assertEqual(pdf_ops.compute_sections(0, {1}), [])

    def test_sanitize_filename_rewrites_reserved_windows_name(self) -> None:
        self.assertEqual(pdf_ops.sanitize_filename("CON", "fallback"), "CON_file")

    def test_sanitize_filename_keeps_extension_like_name_safe(self) -> None:
        self.assertEqual(pdf_ops.sanitize_filename("aux.pdf", "fallback"), "aux.pdf_file")

    def test_default_combined_output_path_uses_source_directory(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source_a = root / "one.pdf"
            source_b = root / "two.png"
            _create_pdf(source_a, ["A"])
            _create_image(source_b, 100, 100, "red")

            destination = pdf_ops.default_combined_output_path([source_a, source_b])

            self.assertEqual(destination.parent.resolve(), root.resolve())
            self.assertEqual(destination.name, "combined.pdf")

    def test_default_combined_output_path_falls_back_to_first_parent(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            first_dir = root / "first"
            second_dir = root / "second"
            first_dir.mkdir()
            second_dir.mkdir()
            source_a = first_dir / "one.pdf"
            source_b = second_dir / "two.png"
            _create_pdf(source_a, ["A"])
            _create_image(source_b, 100, 100, "blue")

            destination = pdf_ops.default_combined_output_path([source_a, source_b])

            self.assertEqual(destination.parent.resolve(), first_dir.resolve())

    def test_combine_documents_to_pdf_preserves_file_order(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            pdf_one = root / "one.pdf"
            image_path = root / "middle.png"
            pdf_two = root / "two.pdf"
            output = root / "combined.pdf"

            _create_pdf(pdf_one, ["PDF ONE"])
            _create_image(image_path, 120, 180, "green")
            _create_pdf(pdf_two, ["PDF TWO A", "PDF TWO B"])

            pdf_ops.combine_documents_to_pdf([pdf_one, image_path, pdf_two], output)

            doc = fitz.open(str(output))
            try:
                self.assertEqual(doc.page_count, 4)
                self.assertIn("PDF ONE", doc.load_page(0).get_text())
                self.assertEqual(doc.load_page(1).get_text().strip(), "")
                self.assertIn("PDF TWO A", doc.load_page(2).get_text())
                self.assertIn("PDF TWO B", doc.load_page(3).get_text())
            finally:
                doc.close()

    def test_combine_documents_to_pdf_supports_image_only_selection(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            image_a = root / "a.jpg"
            image_b = root / "b.png"
            output = root / "images.pdf"

            _create_image(image_a, 200, 120, "yellow")
            _create_image(image_b, 140, 220, "magenta")

            pdf_ops.combine_documents_to_pdf([image_a, image_b], output)

            doc = fitz.open(str(output))
            try:
                self.assertEqual(doc.page_count, 2)
            finally:
                doc.close()

    def test_combine_documents_to_pdf_deletes_sources_after_success(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            pdf_path = root / "one.pdf"
            image_path = root / "two.png"
            output = root / "combined.pdf"

            _create_pdf(pdf_path, ["PDF ONE"])
            _create_image(image_path, 120, 180, "green")

            pdf_ops.combine_documents_to_pdf(
                [pdf_path, image_path],
                output,
                delete_sources=True,
            )

            self.assertTrue(output.exists())
            self.assertFalse(pdf_path.exists())
            self.assertFalse(image_path.exists())

    def test_combine_documents_to_pdf_deletes_single_image_source_after_success(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            image_path = root / "photo.png"
            output = root / "photo.pdf"

            _create_image(image_path, 120, 180, "green")

            pdf_ops.combine_documents_to_pdf(
                [image_path],
                output,
                delete_sources=True,
            )

            self.assertTrue(output.exists())
            self.assertFalse(image_path.exists())

    def test_combine_documents_to_pdf_balanced_profile_downscales_large_images(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            image_path = root / "large.png"
            output = root / "large.pdf"

            _create_image(image_path, 3200, 1600, "cyan")

            pdf_ops.combine_documents_to_pdf([image_path], output)

            doc = fitz.open(str(output))
            try:
                first_page = doc.load_page(0)
                self.assertLessEqual(max(first_page.rect.width, first_page.rect.height), 2200)
            finally:
                doc.close()

    def test_validate_combine_sources_rejects_unsupported_extensions(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "notes.txt"
            source.write_text("not a supported document", encoding="utf-8")

            with self.assertRaises(pdf_ops.UnsupportedSourceError):
                pdf_ops.validate_combine_sources([source])

    def test_split_pdf_keeps_default_page_order_when_not_specified(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "source.pdf"
            _create_pdf(source, ["ONE", "TWO", "THREE"])

            outputs = pdf_ops.split_pdf(
                source_path=source,
                password=None,
                split_starts={3},
                section_names={1: "first", 3: "second"},
                page_rotations={},
                output_dir=root,
            )

            self.assertEqual([path.name for path in outputs], ["first.pdf", "second.pdf"])
            first_doc = fitz.open(str(outputs[0]))
            second_doc = fitz.open(str(outputs[1]))
            try:
                self.assertEqual(first_doc.page_count, 2)
                self.assertIn("ONE", first_doc.load_page(0).get_text())
                self.assertIn("TWO", first_doc.load_page(1).get_text())
                self.assertEqual(second_doc.page_count, 1)
                self.assertIn("THREE", second_doc.load_page(0).get_text())
            finally:
                first_doc.close()
                second_doc.close()

    def test_split_pdf_uses_reordered_page_sequence(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "source.pdf"
            _create_pdf(source, ["ONE", "TWO", "THREE"])

            outputs = pdf_ops.split_pdf(
                source_path=source,
                password=None,
                split_starts=set(),
                section_names={1: "reordered"},
                page_rotations={},
                page_order=[2, 1, 3],
                output_dir=root,
            )

            doc = fitz.open(str(outputs[0]))
            try:
                self.assertEqual(doc.page_count, 3)
                self.assertIn("TWO", doc.load_page(0).get_text())
                self.assertIn("ONE", doc.load_page(1).get_text())
                self.assertIn("THREE", doc.load_page(2).get_text())
            finally:
                doc.close()

    def test_split_pdf_applies_rotation_to_reordered_source_page(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "source.pdf"
            _create_pdf(source, ["ONE", "TWO", "THREE"])

            outputs = pdf_ops.split_pdf(
                source_path=source,
                password=None,
                split_starts=set(),
                section_names={1: "rotated"},
                page_rotations={2: 90},
                page_order=[2, 1, 3],
                output_dir=root,
            )

            doc = fitz.open(str(outputs[0]))
            try:
                first_page = doc.load_page(0)
                self.assertIn("TWO", first_page.get_text())
                self.assertAlmostEqual(first_page.rect.width, 400.0, places=1)
                self.assertAlmostEqual(first_page.rect.height, 300.0, places=1)
            finally:
                doc.close()

    def test_split_pdf_split_markers_follow_moved_page_content(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "source.pdf"
            _create_pdf(source, ["ONE", "TWO", "THREE", "FOUR"])

            outputs = pdf_ops.split_pdf(
                source_path=source,
                password=None,
                split_starts={2},
                section_names={1: "first", 2: "second"},
                page_rotations={},
                page_order=[1, 3, 2, 4],
                output_dir=root,
            )

            self.assertEqual([path.name for path in outputs], ["first.pdf", "second.pdf"])
            first_doc = fitz.open(str(outputs[0]))
            second_doc = fitz.open(str(outputs[1]))
            try:
                self.assertEqual(first_doc.page_count, 1)
                self.assertIn("ONE", first_doc.load_page(0).get_text())
                self.assertEqual(second_doc.page_count, 3)
                self.assertIn("THREE", second_doc.load_page(0).get_text())
                self.assertIn("TWO", second_doc.load_page(1).get_text())
                self.assertIn("FOUR", second_doc.load_page(2).get_text())
            finally:
                first_doc.close()
                second_doc.close()

    def test_split_pdf_optionally_deletes_source_after_success(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "source.pdf"
            _create_pdf(source, ["ONE", "TWO"])

            outputs = pdf_ops.split_pdf(
                source_path=source,
                password=None,
                split_starts=set(),
                section_names={1: "out"},
                page_rotations={},
                delete_source=True,
                output_dir=root,
            )

            self.assertEqual([path.name for path in outputs], ["out.pdf"])
            self.assertFalse(source.exists())
            self.assertTrue(outputs[0].exists())


class LauncherParseTests(unittest.TestCase):
    def test_parse_args_keeps_existing_split_behavior(self) -> None:
        mode, paths = launcher.parse_args([r"C:\Docs\file.pdf"])
        self.assertEqual(mode, "split")
        self.assertEqual(paths, [r"C:\Docs\file.pdf"])

    def test_parse_args_accepts_combine_mode_with_multiple_sources(self) -> None:
        mode, paths = launcher.parse_args(
            ["combine", r"C:\Docs\one file.pdf", r"C:\Docs\scan 1.png"]
        )
        self.assertEqual(mode, "combine")
        self.assertEqual(paths, [r"C:\Docs\one file.pdf", r"C:\Docs\scan 1.png"])

    def test_parse_args_keeps_explorer_selection_flag_for_combine_mode(self) -> None:
        mode, paths = launcher.parse_args(
            [
                "combine",
                "--from-explorer-selection",
                r"C:\Docs\one file.pdf",
                r"C:\Docs\scan 1.png",
            ]
        )
        self.assertEqual(mode, "combine")
        self.assertEqual(
            paths,
            [
                "--from-explorer-selection",
                r"C:\Docs\one file.pdf",
                r"C:\Docs\scan 1.png",
            ],
        )

    def test_parse_args_supports_convert_image_mode(self) -> None:
        mode, paths = launcher.parse_args(["convert-image", r"C:\Docs\scan 1.png"])
        self.assertEqual(mode, "convert-image")
        self.assertEqual(paths, [r"C:\Docs\scan 1.png"])

    def test_parse_args_keeps_explorer_selection_flag_for_convert_image_mode(self) -> None:
        mode, paths = launcher.parse_args(
            [
                "convert-image",
                "--from-explorer-selection",
                r"C:\Docs\scan 1.png",
                r"C:\Docs\scan 2.jpg",
            ]
        )
        self.assertEqual(mode, "convert-image")
        self.assertEqual(
            paths,
            [
                "--from-explorer-selection",
                r"C:\Docs\scan 1.png",
                r"C:\Docs\scan 2.jpg",
            ],
        )


if __name__ == "__main__":
    unittest.main()
