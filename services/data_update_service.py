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

# 로컬에 설치된 버전 = 지금 앱이 쓰는 data/version.json (최초 실행 시 exe에 번들된 값이
# 그대로 시드되고, apply_update 성공 시 여기로 새 버전을 기록함) -> 이거 하나만 기준으로 비교.
# ponytail: 별도 "마지막으로 적용한 버전" 마커 파일을 따로 뒀었는데, 그 파일은 업데이트를
# 한 번도 적용 안 하면 끝까지 안 생겨서 매번 "새 버전 있음"으로 잘못 뜨는 버그가 있었음
# (로컬 data/version.json엔 이미 최신 버전이 있어도 비교 대상이 아니었음) -> 제거하고 이걸로 통일.
LOCAL_VERSION_FILE = DATA_DIR / "version.json"

# ponytail: 원격 버전 확인은 앱 시작 직후 QTimer로 미뤄서 동기 호출(짧은 타임아웃)로 처리.
# 파일이 커지거나 체감 지연이 생기면 QThread로 옮기는 걸 고려.


def _fetch_bytes(url: str, timeout: int = 5) -> bytes:
    request = urllib.request.Request(url, headers={"User-Agent": "BioFlow-update-check/1.0"})
    with urllib.request.urlopen(request, timeout=timeout) as response:
        return response.read()


def _read_local_version() -> str:
    try:
        with LOCAL_VERSION_FILE.open("r", encoding="utf-8") as f:
            return str(json.load(f).get("version", ""))
    except Exception:
        return ""


def _write_local_version(version: str):
    LOCAL_VERSION_FILE.parent.mkdir(parents=True, exist_ok=True)
    with LOCAL_VERSION_FILE.open("w", encoding="utf-8") as f:
        json.dump({"version": version}, f, ensure_ascii=False, indent=2)


def check_for_update() -> dict[str, Any] | None:
    """원격 data/version.json을 조회해서 로컬 data/version.json과 다르면 그 정보를 반환.
    네트워크 실패/버전 동일 시 None (조용히 넘어감 - 오프라인이어도 앱 시작에 지장 없음)."""
    try:
        remote = json.loads(_fetch_bytes(REMOTE_VERSION_URL).decode("utf-8"))
    except Exception:
        return None

    remote_version = str(remote.get("version", ""))
    if not remote_version or remote_version == _read_local_version():
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

    _write_local_version(str(remote_info.get("version", "")))
