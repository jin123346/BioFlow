"""Read filenames from a selected local directory; never rename or copy files."""

from collections import defaultdict
import os
from pathlib import Path
import re

from services.sql_excel_service import MissingFormulaValue
from utils.text_utils import normalize_text


MEDIA_EXTENSIONS = {
    "image": {".jpg", ".jpeg", ".png", ".tif", ".tiff", ".heic", ".heif", ".webp", ".bmp", ".gif"},
    "video": {".mp4", ".mov", ".m4v", ".avi", ".mts", ".m2ts", ".3gp", ".mkv", ".wmv", ".webm", ".mpg", ".mpeg"},
    "sound": {".wav", ".mp3", ".m4a", ".flac", ".ogg", ".aac", ".wma", ".aiff", ".aif"},
}


class MediaFileIndex:
    def __init__(self, folder, recursive=True, skip_previews=True):
        if not str(folder or "").strip():
            raise ValueError("파일명 매칭에 사용할 폴더를 선택해 주세요.")
        self.root = Path(folder).resolve()
        if not self.root.is_dir():
            raise ValueError("파일명 매칭 폴더를 찾을 수 없습니다.")
        self.by_stem = defaultdict(list)
        def failed(error):
            raise ValueError(f"매칭 폴더를 읽지 못했습니다: {error}") from error
        for root, directories, files in os.walk(self.root, onerror=failed, followlinks=False):
            if skip_previews:
                directories[:] = [name for name in directories if name != "썸네일"]
            for name in files:
                path = Path(root) / name
                if path.is_file() and not path.is_symlink():
                    self.by_stem[path.stem.casefold()].append(path)
            if not recursive:
                break

    def filename(self, database_number, kind, column):
        return self.filepath(database_number, kind, column).name

    def filepath(self, database_number, kind, column):
        if isinstance(database_number, MissingFormulaValue):
            raise ValueError("파일명 매칭용 DB 번호 수식을 엑셀에서 재계산해 주세요.")
        number = normalize_text(database_number)
        if not number:
            raise ValueError("파일명 매칭용 데이터베이스번호 컬럼 또는 값이 없습니다.")
        group = "image" if column["name"] == "LVBNG_MVPCT_RPRSN_IMAGE" or "image" in kind else (
            "video" if kind == "observation_video" else "sound")
        extensions = MEDIA_EXTENSIONS[group]
        chosen = str(column.get("filename_extension") or "").strip().lower()
        if chosen:
            chosen = chosen if chosen.startswith(".") else "." + chosen
            if not re.fullmatch(r"\.[a-z0-9]{1,12}", chosen) or chosen not in extensions:
                raise ValueError(f"{column['name']}: 선택한 미디어 유형에 맞는 확장자를 지정해 주세요.")
            extensions = {chosen}
        candidates = [path for path in self.by_stem.get(number.casefold(), []) if path.suffix.lower() in extensions]
        if not candidates:
            raise ValueError(f"DB 번호 '{number}'와 일치하는 {group} 파일이 없습니다.")
        if len(candidates) != 1:
            names = ", ".join(str(path.relative_to(self.root)) for path in sorted(candidates)[:5])
            raise ValueError(f"DB 번호 '{number}'의 파일 후보가 {len(candidates)}개입니다: {names}. 폴더나 확장자를 좁혀 주세요.")
        return candidates[0]
