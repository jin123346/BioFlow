"""File dialogs that retain the last selected directory between app launches."""

import json
import logging
from pathlib import Path

from PySide6.QtWidgets import QFileDialog

from config.paths import CACHE_DIR


class RecentFileDialog:
    cache_path = CACHE_DIR / "recent_file_locations.json"

    @classmethod
    def _read(cls):
        try:
            data = json.loads(cls.cache_path.read_text(encoding="utf-8"))
            return data if isinstance(data, dict) else {}
        except (OSError, ValueError):
            return {}

    @classmethod
    def _start(cls, mode, caption, directory):
        data = cls._read()
        for key in (f"{mode}:{caption}", f"{mode}:last"):
            cached = data.get(key)
            if isinstance(cached, str) and Path(cached).is_dir():
                if mode == "save" and directory:
                    return str(Path(cached) / Path(directory).name)
                return cached
        return directory

    @classmethod
    def _remember(cls, mode, caption, selection):
        if not selection:
            return
        selected = Path(selection).resolve()
        folder = selected if mode == "folder" else selected.parent
        data = cls._read()
        data[f"{mode}:{caption}"] = str(folder)
        data[f"{mode}:last"] = str(folder)
        try:
            cls.cache_path.parent.mkdir(parents=True, exist_ok=True)
            temporary = cls.cache_path.with_suffix(".tmp")
            temporary.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
            temporary.replace(cls.cache_path)
        except OSError:
            logging.getLogger(__name__).warning("Could not cache file dialog location", exc_info=True)

    @classmethod
    def getOpenFileName(cls, parent=None, caption="", directory="", *args, **kwargs):
        result = QFileDialog.getOpenFileName(parent, caption, cls._start("open", caption, directory), *args, **kwargs)
        cls._remember("open", caption, result[0])
        return result

    @classmethod
    def getOpenFileNames(cls, parent=None, caption="", directory="", *args, **kwargs):
        result = QFileDialog.getOpenFileNames(parent, caption, cls._start("open", caption, directory), *args, **kwargs)
        cls._remember("open", caption, result[0][0] if result[0] else "")
        return result

    @classmethod
    def getSaveFileName(cls, parent=None, caption="", directory="", *args, **kwargs):
        result = QFileDialog.getSaveFileName(parent, caption, cls._start("save", caption, directory), *args, **kwargs)
        cls._remember("save", caption, result[0])
        return result

    @classmethod
    def getExistingDirectory(cls, parent=None, caption="", directory="", *args, **kwargs):
        result = QFileDialog.getExistingDirectory(parent, caption, cls._start("folder", caption, directory), *args, **kwargs)
        cls._remember("folder", caption, result)
        return result
