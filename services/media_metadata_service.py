"""Read-only media identification and metadata extraction helpers."""

from __future__ import annotations

import json
import logging
import subprocess
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from services.media_storage_service import find_media_tool

logger = logging.getLogger("bioflow.media_metadata")

IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".tif", ".tiff", ".heic", ".png"}
VIDEO_EXTENSIONS = {".mp4", ".mov", ".m4v", ".avi", ".mts", ".3gp"}
AUDIO_EXTENSIONS = {".wav", ".mp3", ".m4a", ".aac", ".flac"}


@dataclass(frozen=True)
class ProbeResult:
    format: dict[str, Any]
    streams: tuple[dict[str, Any], ...]
    error: str = ""

    @property
    def duration(self) -> float | None:
        candidates = [self.format.get("duration")]
        candidates.extend(stream.get("duration") for stream in self.streams)
        for value in candidates:
            try:
                if value not in (None, "", "N/A"):
                    return float(value)
            except (TypeError, ValueError):
                continue
        return None

    @property
    def dimensions(self) -> tuple[int | None, int | None]:
        for stream in self.streams:
            if stream.get("codec_type") == "video":
                try:
                    return int(stream.get("width")), int(stream.get("height"))
                except (TypeError, ValueError):
                    break
        return None, None


def classify_media_file(path: Path) -> str | None:
    """Classify supported media, checking readable image content when possible."""
    suffix = path.suffix.lower()
    try:
        from PIL import Image

        with Image.open(path) as image:
            image.verify()
        return "이미지"
    except Exception as exc:
        if suffix in IMAGE_EXTENSIONS:
            logger.warning("Image signature check failed; using extension for %s: %s", path, exc)
    if suffix in IMAGE_EXTENSIONS:
        return "이미지"
    if suffix in VIDEO_EXTENSIONS:
        return "동영상"
    if suffix in AUDIO_EXTENSIONS:
        return "사운드"
    return None


def probe_media(path: Path, timeout: int = 10) -> ProbeResult:
    """Probe a file without modifying it; return/log failures per file."""
    ffprobe = find_media_tool("ffprobe")
    if not ffprobe:
        return ProbeResult({}, (), "ffprobe를 찾을 수 없음")
    try:
        result = subprocess.run(
            [ffprobe, "-v", "error", "-print_format", "json", "-show_format", "-show_streams", str(path)],
            capture_output=True, text=True, encoding="utf-8", errors="replace",
            timeout=timeout, check=False,
        )
        if result.returncode:
            error = result.stderr.strip() or f"ffprobe 종료 코드 {result.returncode}"
            logger.warning("Metadata probe failed for %s: %s", path, error)
            return ProbeResult({}, (), error)
        data = json.loads(result.stdout or "{}")
        return ProbeResult(data.get("format", {}), tuple(data.get("streams", ())))
    except Exception as exc:
        logger.exception("Metadata probe failed for %s", path)
        return ProbeResult({}, (), str(exc))


def duration_text(probe: ProbeResult) -> str:
    duration = probe.duration
    if duration is None:
        return ""
    total_seconds = int(round(duration))
    hours, remainder = divmod(total_seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    return f"{hours}:{minutes:02d}:{seconds:02d}" if hours else f"{minutes}:{seconds:02d}"
