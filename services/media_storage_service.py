"""Copy media under its database identifier without silently changing the name."""

import filecmp
from pathlib import Path
import shutil
import sys
import re

from utils.text_utils import normalize_text


def media_destination(root, record):
    root = Path(root)
    if normalize_text(record.get("미디어타입")) != "동영상":
        return root
    species_id = normalize_text(record.get("종고유ID"))
    if not species_id or species_id == "0":
        raise ValueError("동영상 저장에 필요한 종고유ID가 없습니다. 국명·학명 조회로 종정보를 먼저 지정해 주세요.")
    if species_id in (".", "..") or species_id.endswith(".") or re.search(r'[\\/:*?"<>|]', species_id):
        raise ValueError("종고유ID에 폴더명으로 사용할 수 없는 문자가 있습니다.")
    return root / species_id


def find_media_tool(name):
    bundled = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parents[1])) / "media_tools" / f"{name}.exe"
    return str(bundled) if bundled.is_file() else shutil.which(name)


def copy_database_media(source, target_dir, stem):
    source, target_dir = Path(source), Path(target_dir)
    target = target_dir / f"{stem}{source.suffix.lower()}"
    if source.resolve() == target.resolve():
        return source
    target_dir.mkdir(parents=True, exist_ok=True)
    if target.exists():
        if filecmp.cmp(source, target, shallow=False):
            return target
        raise FileExistsError(f"같은 DB 번호의 다른 파일이 이미 있습니다: {target}\nDB 번호 또는 저장 위치를 확인해 주세요.")
    shutil.copy2(source, target)
    return target
