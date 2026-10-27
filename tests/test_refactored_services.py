import hashlib
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

from services.excel_service import derived_excel_path, normalize_excel_date
from services.identifier_service import build_identifier, expected_identifier, media_code
from services.media_metadata_service import ProbeResult, duration_text
from services.media_storage_service import copy_database_media


def digest(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


class IdentifierServiceTests(unittest.TestCase):
    def test_existing_observation_and_specimen_formats_are_preserved(self):
        self.assertEqual(
            build_identifier("NSMK", "PL", 2026, 7, 5, media="OI"),
            "NSMK-OI-PL-260700005",
        )
        self.assertEqual(
            build_identifier("NSMK", "PL", 2026, 7, 5, specimen=True),
            "NSMK-PL-260700005",
        )

    def test_existing_media_mapping_and_validation_fallback_are_preserved(self):
        self.assertEqual(media_code("동영상"), "OM")
        self.assertEqual(media_code("", "OI"), "OI")
        result = expected_identifier(
            {
                "기관코드": "NSMK",
                "분류코드": "PL",
                "미디어타입": "이미지",
                "사업년도": "2026",
                "순번": "7",
                "사업별 순번": "5",
            }
        )
        self.assertEqual(result.expected, "NSMK-OI-PL-260700005")
        self.assertTrue(result.pattern.fullmatch(result.expected))

    def test_missing_type_defaults_to_ob_and_types_keep_separate_namespaces(self):
        base = {
            "기관코드": "NSMK", "분류코드": "PL", "사업년도": "2026",
            "순번": "7", "사업별 순번": "5",
        }
        defaulted = expected_identifier(base)
        image = expected_identifier({**base, "미디어타입": "이미지"})
        video = expected_identifier({**base, "미디어타입": "동영상"})
        self.assertEqual(media_code(""), "OB")
        self.assertEqual(defaulted.expected, "NSMK-OB-PL-260700005")
        self.assertEqual(image.expected, "NSMK-OI-PL-260700005")
        self.assertEqual(video.expected, "NSMK-OM-PL-260700005")
        self.assertEqual(len({defaulted.expected, image.expected, video.expected}), 3)


class ExcelServiceTests(unittest.TestCase):
    def test_existing_date_and_output_name_rules_are_preserved(self):
        self.assertEqual(normalize_excel_date("260508"), "2026-05-08")
        self.assertEqual(normalize_excel_date(2, datetime(1899, 12, 30)), "1900-01-02")
        source = Path("new_data.xlsx")
        self.assertEqual(derived_excel_path(source, "_완성본"), Path("new_data_완성본.xlsx"))


class OriginalMediaImmutabilityTests(unittest.TestCase):
    def test_copy_never_changes_original_content_name_or_timestamp(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source = root / "original" / "camera-file.JPG"
            source.parent.mkdir()
            source.write_bytes(b"original-media-bytes")
            before = (source.name, digest(source), source.stat().st_mtime_ns)

            destination = copy_database_media(source, root / "export", "DB-001")

            self.assertEqual((source.name, digest(source), source.stat().st_mtime_ns), before)
            self.assertTrue(source.is_file())
            self.assertEqual(destination.read_bytes(), source.read_bytes())
            self.assertNotEqual(destination.resolve(), source.resolve())


class MediaMetadataTests(unittest.TestCase):
    def test_duration_falls_back_to_video_stream_and_dimensions_are_available(self):
        probe = ProbeResult(
            {"duration": "N/A"},
            ({"codec_type": "video", "duration": "65.4", "width": 1920, "height": 1080},),
        )
        self.assertEqual(duration_text(probe), "1:05")
        self.assertEqual(probe.dimensions, (1920, 1080))


if __name__ == "__main__":
    unittest.main()
