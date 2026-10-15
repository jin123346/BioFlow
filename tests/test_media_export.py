import os
from pathlib import Path
import subprocess
import tempfile
import unittest
from unittest.mock import Mock, patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
from PySide6.QtWidgets import QApplication, QCheckBox, QLineEdit, QTableWidget, QTableWidgetItem, QWidget

from services.media_storage_service import copy_database_media, find_media_tool
from ui.tabs.upload_tab import ROW_COLUMNS, SPECIMEN_ROW_COLUMNS, UploadTab


class MediaExportTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def editor(self, source, destination, specimen=False, media_type="이미지"):
        editor = UploadTab.__new__(UploadTab)
        QWidget.__init__(editor)
        editor.autosave_suspended = editor.specimen_autosave_suspended = True
        editor._specimen_taxon_edited_fields = set()
        editor.preserve_database_numbers = QCheckBox()
        editor.preserve_database_numbers.setChecked(True)
        editor.set_save_status = Mock()
        editor.set_specimen_save_status = Mock()
        editor.extract_location_from_media = Mock(return_value=(None, None, ""))
        columns = SPECIMEN_ROW_COLUMNS if specimen else ROW_COLUMNS
        table = QTableWidget(1, len(columns))
        values = {"파일경로": str(source), "데이터베이스번호": "TEST-SP-BI-260100100",
                  "미디어타입": media_type, "종고유ID": "12345", "사업년도": "2026", "순번": "1", "사업별 순번": "100"}
        for col, field in enumerate(columns):
            table.setItem(0, col, QTableWidgetItem(values.get(field, "")))
        setattr(editor, "specimen_data_table" if specimen else "data_table", table)
        setattr(editor, "specimen_media_storage_dir" if specimen else "media_storage_dir", str(destination))
        setattr(editor, "specimen_media_storage_path_input" if specimen else "media_storage_path_input", QLineEdit())
        return editor, table, columns

    def test_image_export_and_db_renumber_for_both_tabs(self):
        for specimen in (False, True):
            with self.subTest(specimen=specimen), tempfile.TemporaryDirectory() as folder:
                root = Path(folder)
                source = root / "원본" / "photo.JPG"
                source.parent.mkdir()
                source.write_bytes(b"original photo")
                destination = root / "저장"
                editor, table, columns = self.editor(source, destination, specimen)
                try:
                    self.assertTrue(editor.prepare_media_for_excel_export(specimen))
                    saved = destination / "TEST-SP-BI-260100100.jpg"
                    self.assertEqual(saved.read_bytes(), source.read_bytes())
                    self.assertEqual(table.item(0, columns.index("파일경로")).text(), str(saved))
                    self.assertEqual(table.item(0, columns.index("원본경로")).text(), str(source))
                    self.assertTrue(editor.prepare_media_for_excel_export(specimen))
                    self.assertEqual(len(list(destination.glob("*.jpg"))), 1)
                    table.item(0, columns.index("데이터베이스번호")).setText("TEST-SP-BI-260100200")
                    self.assertTrue(editor.prepare_media_for_excel_export(specimen))
                    renamed = destination / "TEST-SP-BI-260100200.jpg"
                    self.assertEqual(renamed.read_bytes(), b"original photo")
                    self.assertEqual(table.item(0, columns.index("파일명")).text(), renamed.name)
                    self.assertEqual(source.read_bytes(), b"original photo")
                finally:
                    editor.close()

    def test_existing_identical_file_is_reused_but_conflicting_file_is_not_overwritten(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source = root / "source.jpg"
            source.write_bytes(b"photo one")
            saved = copy_database_media(source, root / "out", "DB100")
            self.assertEqual(copy_database_media(source, root / "out", "DB100"), saved)
            source.write_bytes(b"photo two")
            with self.assertRaises(FileExistsError):
                copy_database_media(source, root / "out", "DB100")
            self.assertEqual(saved.read_bytes(), b"photo one")
            self.assertFalse((root / "out" / "DB100_2.jpg").exists())

    def test_cancel_destination_stops_export_without_copying(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source = root / "image.jpg"
            source.write_bytes(b"photo")
            editor, _, _ = self.editor(source, root / "out")
            try:
                editor.media_storage_dir = ""
                with patch("ui.tabs.upload_tab.QFileDialog.getExistingDirectory", return_value=""):
                    self.assertFalse(editor.prepare_media_for_excel_export(False))
                self.assertFalse((root / "out").exists())
            finally:
                editor.close()

    def test_excel_export_asks_for_destination_each_time_for_both_tabs(self):
        for specimen in (False, True):
            with self.subTest(specimen=specimen), tempfile.TemporaryDirectory() as folder:
                root = Path(folder)
                source = root / "original.jpg"
                source.write_bytes(b"image")
                editor, table, columns = self.editor(source, root / "previous", specimen)
                try:
                    for name in ("first", "second"):
                        destination = root / name
                        with patch("ui.tabs.upload_tab.QFileDialog.getExistingDirectory", return_value=str(destination)) as chooser:
                            self.assertTrue(editor.prepare_media_for_excel_export(specimen, choose_destination=True))
                            chooser.assert_called_once()
                        saved = Path(table.item(0, columns.index("파일경로")).text())
                        self.assertEqual(saved.parent, destination)
                        self.assertEqual(saved.read_bytes(), b"image")
                    self.assertFalse((root / "previous").exists())
                    with patch("ui.tabs.upload_tab.QFileDialog.getExistingDirectory", return_value="") as chooser:
                        self.assertFalse(editor.prepare_media_for_excel_export(specimen, choose_destination=True))
                        chooser.assert_called_once()
                    self.assertEqual(table.rowCount(), 1)
                finally:
                    editor.close()

    def test_no_media_does_not_ask_for_destination(self):
        for specimen in (False, True):
            editor, table, columns = self.editor(Path("unused.jpg"), Path("unused"), specimen)
            try:
                table.item(0, columns.index("파일경로")).setText("")
                with patch("ui.tabs.upload_tab.QFileDialog.getExistingDirectory") as chooser:
                    self.assertTrue(editor.prepare_media_for_excel_export(specimen, choose_destination=True))
                    chooser.assert_not_called()
            finally:
                editor.close()

    def test_export_form_sync_defers_media_copy(self):
        for specimen in (False, True):
            editor, _, _ = self.editor(Path("unused.jpg"), Path("unused"), specimen)
            try:
                editor._defer_media_copy = True
                values = {"파일경로": "original.jpg"}
                prepare = editor.prepare_image_for_specimen_record if specimen else editor.prepare_media_for_record
                with patch("ui.tabs.upload_tab.QFileDialog.getExistingDirectory") as chooser:
                    self.assertIs(prepare(values), values)
                    chooser.assert_not_called()
            finally:
                editor.close()

    def test_missing_source_stops_export(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            editor, _, _ = self.editor(root / "missing.jpg", root / "out")
            try:
                with self.assertRaises(FileNotFoundError):
                    editor.prepare_media_for_excel_export(False)
            finally:
                editor.close()

    def test_capture_failure_is_reported_and_keeps_rows(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source = root / "video.mp4"
            source.write_bytes(b"test video")
            editor, table, _ = self.editor(source, root / "out", media_type="동영상")
            try:
                with patch("ui.tabs.upload_tab.find_media_tool", return_value=None):
                    with self.assertRaisesRegex(RuntimeError, "동영상 캡처"):
                        editor.prepare_media_for_excel_export(False)
                self.assertEqual(table.rowCount(), 1)
            finally:
                editor.close()

    def test_video_without_species_id_stops_before_copy(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source = root / "original.mp4"
            source.write_bytes(b"video")
            editor, table, columns = self.editor(source, root / "out", media_type="동영상")
            try:
                table.item(0, columns.index("종고유ID")).setText("")
                with self.assertRaisesRegex(ValueError, "종고유ID"):
                    editor.prepare_media_for_excel_export(False)
                self.assertFalse((root / "out").exists())
                self.assertEqual(table.rowCount(), 1)
            finally:
                editor.close()

    def test_capture_never_runs_on_original_unsaved_video(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source = root / "original.mp4"
            source.write_bytes(b"video")
            editor, _, _ = self.editor(source, root / "out", media_type="동영상")
            try:
                editor.capture_video_thumbnail = Mock()
                self.assertEqual(editor.capture_pending_video_thumbnails(), (0, 0, 1, 1))
                editor.capture_video_thumbnail.assert_not_called()
                self.assertFalse(source.with_suffix(".jpg").exists())
            finally:
                editor.close()

    def test_short_video_is_saved_and_captured_with_same_db_name(self):
        ffmpeg = find_media_tool("ffmpeg")
        if not ffmpeg:
            self.skipTest("ffmpeg is unavailable")
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source = root / "original.mp4"
            subprocess.run([ffmpeg, "-v", "error", "-f", "lavfi", "-i", "color=c=black:s=32x32:r=10",
                            "-t", "0.2", "-an", "-c:v", "mpeg4", str(source)],
                           capture_output=True, check=True, timeout=20)
            editor, table, columns = self.editor(source, root / "out", media_type="동영상")
            try:
                self.assertTrue(editor.prepare_media_for_excel_export(False))
                saved = Path(table.item(0, columns.index("파일경로")).text())
                capture = editor.video_thumbnail_path(saved)
                self.assertEqual(saved.name, "TEST-SP-BI-260100100.mp4")
                self.assertEqual(capture.name, "TEST-SP-BI-260100100.jpg")
                self.assertEqual(saved.parent, root / "out" / "12345")
                self.assertEqual(capture.parent, saved.parent)
                self.assertGreater(capture.stat().st_size, 0)
                self.assertEqual(table.item(0, columns.index("동영상 파일명")).text(), saved.name)
                self.assertTrue(source.exists())
                self.assertTrue(editor.prepare_media_for_excel_export(False))
                table.item(0, columns.index("데이터베이스번호")).setText("TEST-SP-BI-260100200")
                self.assertTrue(editor.prepare_media_for_excel_export(False))
                new_video = root / "out" / "12345" / "TEST-SP-BI-260100200.mp4"
                self.assertTrue(new_video.exists())
                self.assertGreater(new_video.with_suffix(".jpg").stat().st_size, 0)
                table.item(0, columns.index("종고유ID")).setText("67890")
                self.assertTrue(editor.prepare_media_for_excel_export(False))
                regrouped = root / "out" / "67890" / new_video.name
                self.assertTrue(regrouped.exists())
                self.assertGreater(regrouped.with_suffix(".jpg").stat().st_size, 0)
            finally:
                editor.close()

    def test_capture_button_only_creates_preview_without_destination_or_species_id(self):
        ffmpeg = find_media_tool("ffmpeg")
        if not ffmpeg:
            self.skipTest("ffmpeg is unavailable")
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source = root / "original.mp4"
            subprocess.run([ffmpeg, "-v", "error", "-f", "lavfi", "-i", "color=c=black:s=32x32:r=10",
                            "-t", "0.2", "-an", "-c:v", "mpeg4", str(source)],
                           capture_output=True, check=True, timeout=20)
            editor, table, columns = self.editor(source, root / "previous", media_type="동영상")
            editor.sync_current_form_to_table = Mock()
            table.item(0, columns.index("종고유ID")).setText("")
            table.item(0, columns.index("데이터베이스번호")).setText("")
            # Preview does not require database metadata or copy unrelated photos.
            table.insertRow(1)
            for col, field in enumerate(columns):
                table.setItem(1, col, QTableWidgetItem({"미디어타입": "이미지", "파일경로": str(root / "missing.jpg")}.get(field, "")))
            try:
                with patch("ui.tabs.upload_tab.QFileDialog.getExistingDirectory") as chooser, \
                     patch("ui.tabs.upload_tab.QMessageBox.information") as info, \
                     patch("ui.tabs.upload_tab.QMessageBox.critical") as error:
                    editor.generate_video_thumbnails()
                    chooser.assert_not_called()
                    error.assert_not_called()
                    self.assertIn("1건 생성", info.call_args.args[2])
                preview = root / "썸네일" / "original.jpg"
                self.assertGreater(preview.stat().st_size, 0)
                self.assertEqual(table.item(0, columns.index("파일경로")).text(), str(source))
                self.assertFalse(source.with_suffix(".jpg").exists())
                self.assertFalse((root / "previous").exists())
                self.assertTrue(source.is_file())
                self.assertEqual(table.rowCount(), 2)
                self.assertFalse(editor._defer_media_copy)
            finally:
                editor.close()


if __name__ == "__main__":
    unittest.main()
