import tempfile
import unittest
from pathlib import Path

from PIL import Image

from services.media_metadata_service import classify_media_file, duration_text, probe_media
from services.media_storage_service import find_media_tool


class MixedMediaMetadataTests(unittest.TestCase):
    def test_mixed_folder_classifies_image_content_and_video_extension_independently(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            disguised_image = root / "photo.bin"
            Image.new("RGB", (8, 6), "red").save(disguised_image, format="PNG")
            video = root / "clip.mp4"
            video.write_bytes(b"fixture")
            text = root / "notes.txt"
            text.write_text("not media", encoding="utf-8")

            self.assertEqual(classify_media_file(disguised_image), "이미지")
            self.assertEqual(classify_media_file(video), "동영상")
            self.assertIsNone(classify_media_file(text))

    def test_real_video_duration_dimensions_and_source_are_read_only(self):
        ffmpeg = find_media_tool("ffmpeg")
        if not ffmpeg:
            self.skipTest("ffmpeg is unavailable")
        import subprocess

        with tempfile.TemporaryDirectory() as folder:
            video = Path(folder) / "source.mp4"
            subprocess.run(
                [ffmpeg, "-v", "error", "-f", "lavfi", "-i", "color=c=black:s=40x30:r=10",
                 "-t", "1.2", "-an", "-c:v", "mpeg4", str(video)],
                capture_output=True, check=True, timeout=20,
            )
            before = (video.read_bytes(), video.stat().st_mtime_ns)
            result = probe_media(video)

            self.assertEqual(result.dimensions, (40, 30))
            self.assertEqual(duration_text(result), "0:01")
            self.assertEqual((video.read_bytes(), video.stat().st_mtime_ns), before)


if __name__ == "__main__":
    unittest.main()
