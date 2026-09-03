import json
import shutil
import tempfile
import urllib.request
from pathlib import Path
from typing import Any

from config.paths import DATA_DIR
from services.species_master_service import SpeciesMasterService

GITHUB_RAW_BASE = "https://raw.githubusercontent.com/jin123346/BioFlow/main"
REMOTE_VERSION_URL = f"{GITHUB_RAW_BASE}/data/version.json"
REMOTE_MASTER_URL = f"{GITHUB_RAW_BASE}/data/master/species_master.xlsx"
REMOTE_INSTITUTION_URL = f"{GITHUB_RAW_BASE}/data/institution/institution.json"

LOCAL_REMOTE_VERSION_FILE = DATA_DIR / "meta" / "remote_data_version.json"

# ponytail: 원격 버전 확인은 앱 시작 직후 QTimer로 미뤄서 동기 호출(짧은 타임아웃)로 처리.
# 파일이 커지거나 체감 지연이 생기면 QThread로 옮기는 걸 고려.


def _fetch_bytes(url: str, timeout: int = 5) -> bytes:
    request = urllib.request.Request(url, headers={"User-Agent": "BioFlow-update-check/1.0"})
    with urllib.request.urlopen(request, timeout=timeout) as response:
        return response.read()


def _read_local_remote_version() -> str:
    try:
        with LOCAL_REMOTE_VERSION_FILE.open("r", encoding="utf-8") as f:
            return str(json.load(f).get("version", ""))
    except Exception:
        return ""


def _write_local_remote_version(version: str):
    LOCAL_REMOTE_VERSION_FILE.parent.mkdir(parents=True, exist_ok=True)
    with LOCAL_REMOTE_VERSION_FILE.open("w", encoding="utf-8") as f:
        json.dump({"version": version}, f, ensure_ascii=False, indent=2)


def check_for_update() -> dict[str, Any] | None:
    """원격 data/version.json을 조회해서 로컬에 반영된 버전과 다르면 그 정보를 반환.
    네트워크 실패/버전 동일 시 None (조용히 넘어감 - 오프라인이어도 앱 시작에 지장 없음)."""
    try:
        remote = json.loads(_fetch_bytes(REMOTE_VERSION_URL).decode("utf-8"))
    except Exception:
        return None

    remote_version = str(remote.get("version", ""))
    if not remote_version or remote_version == _read_local_remote_version():
        return None
    return remote


def apply_update(remote_info: dict[str, Any]) -> None:
    """원격 species_master.xlsx / institution.json을 내려받아 로컬에 반영."""
    with tempfile.TemporaryDirectory() as tmp_dir:
        master_tmp = Path(tmp_dir) / "species_master.xlsx"
        master_tmp.write_bytes(_fetch_bytes(REMOTE_MASTER_URL, timeout=30))
        service = SpeciesMasterService()
        df = service.load_excel_file(str(master_tmp))
        service.save_master_to_local(df, str(master_tmp))

        institution_tmp = Path(tmp_dir) / "institution.json"
        institution_tmp.write_bytes(_fetch_bytes(REMOTE_INSTITUTION_URL, timeout=15))
        institution_dir = DATA_DIR / "institution"
        institution_dir.mkdir(parents=True, exist_ok=True)
        shutil.copy2(institution_tmp, institution_dir / "institution.json")

    _write_local_remote_version(str(remote_info.get("version", "")))
