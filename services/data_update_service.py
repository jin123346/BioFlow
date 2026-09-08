import json
import logging
import tempfile
import time
import urllib.request
from pathlib import Path
from typing import Any

from config.paths import DATA_DIR
from services.species_master_service import SpeciesMasterService

logger = logging.getLogger("bioflow.data_update")

GITHUB_RAW_BASE = "https://raw.githubusercontent.com/jin123346/BioFlow/main"
REMOTE_VERSION_URL = f"{GITHUB_RAW_BASE}/data/version.json"
REMOTE_MASTER_URL = f"{GITHUB_RAW_BASE}/data/master/species_master.xlsx"
REMOTE_ORGANIZATIONS_URL = f"{GITHUB_RAW_BASE}/data/reference/organizations.json"
REMOTE_PROJECTS_URL = f"{GITHUB_RAW_BASE}/data/reference/upload_projects.json"
# 앱(exe) 자체의 새 릴리즈 확인용. data/version.json(종정보 마스터 버전)과는 별개 파일.
# 새 버전을 배포할 때 이 파일의 version/download_url을 같이 올려야 알림이 뜸.
REMOTE_APP_VERSION_URL = f"{GITHUB_RAW_BASE}/app_version.json"
DEFAULT_DOWNLOAD_URL = "https://github.com/jin123346/BioFlow/releases/latest"

# 로컬에 설치된 버전 = 지금 앱이 쓰는 data/version.json (최초 실행 시 exe에 번들된 값이
# 그대로 시드되고, apply_update 성공 시 여기로 새 버전을 기록함) -> 이거 하나만 기준으로 비교.
# ponytail: 별도 "마지막으로 적용한 버전" 마커 파일을 따로 뒀었는데, 그 파일은 업데이트를
# 한 번도 적용 안 하면 끝까지 안 생겨서 매번 "새 버전 있음"으로 잘못 뜨는 버그가 있었음
# (로컬 data/version.json엔 이미 최신 버전이 있어도 비교 대상이 아니었음) -> 제거하고 이걸로 통일.
LOCAL_VERSION_FILE = DATA_DIR / "version.json"

# ponytail: 원격 버전 확인은 앱 시작 직후 QTimer로 미뤄서 동기 호출(짧은 타임아웃)로 처리.
# 파일이 커지거나 체감 지연이 생기면 QThread로 옮기는 걸 고려.


def _fetch_bytes(url: str, timeout: int = 5, bust_cache: bool = False) -> bytes:
    # raw.githubusercontent.com은 CDN(Fastly) 캐시를 태워서, push 직후 몇 분간은 예전
    # 내용이 그대로 나올 수 있음 -> 버전 체크처럼 "항상 최신이어야 하는" 요청만 매번 다른
    # URL(쿼리스트링)로 만들어 캐시를 우회함
    if bust_cache:
        url = f"{url}?_={int(time.time())}"
    request = urllib.request.Request(
        url, headers={"User-Agent": "BioFlow-update-check/1.0", "Cache-Control": "no-cache"}
    )
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
    네트워크 실패 시 None (조용히 넘어감 - 오프라인이어도 앱 시작에 지장 없음). 실패/결과는
    로그로 남겨서, "진짜 업데이트가 없는 건지" "확인 자체가 실패한 건지" 나중에 구분 가능."""
    try:
        remote = json.loads(_fetch_bytes(REMOTE_VERSION_URL, bust_cache=True).decode("utf-8"))
    except Exception as exc:
        logger.info("check_for_update: 원격 version.json 조회 실패 - %s", exc)
        return None

    remote_version = str(remote.get("version", ""))
    local_version = _read_local_version()
    logger.info("check_for_update: remote=%r local=%r", remote_version, local_version)
    if not remote_version or remote_version == local_version:
        return None
    return remote


def check_for_app_update(local_version: str) -> dict[str, Any] | None:
    """앱(exe) 자체의 새 릴리즈가 있는지 원격 app_version.json으로 확인.
    exe는 실행 중인 자기 자신을 못 바꿔치기하므로, 여기선 새 버전 유무와 다운로드 페이지
    URL만 반환 -- 실제 설치는 사용자가 그 페이지에서 직접 받아서 함."""
    try:
        remote = json.loads(_fetch_bytes(REMOTE_APP_VERSION_URL, bust_cache=True).decode("utf-8"))
    except Exception as exc:
        logger.info("check_for_app_update: 원격 app_version.json 조회 실패 - %s", exc)
        return None

    remote_version = str(remote.get("version", ""))
    logger.info("check_for_app_update: remote=%r local=%r", remote_version, local_version)
    if not remote_version or not local_version or remote_version == local_version:
        return None
    remote.setdefault("download_url", DEFAULT_DOWNLOAD_URL)
    return remote


def apply_update(remote_info: dict[str, Any]) -> None:
    """원격 species_master.xlsx / organizations.json을 내려받아 로컬에 그대로 덮어쓰고,
    사업 목록(upload_projects.json)은 사용자가 직접 추가했을 수 있어 덮어쓰지 않고 병합."""
    with tempfile.TemporaryDirectory() as tmp_dir:
        master_tmp = Path(tmp_dir) / "species_master.xlsx"
        master_tmp.write_bytes(_fetch_bytes(REMOTE_MASTER_URL, timeout=30))
        service = SpeciesMasterService()
        df = service.load_excel_file(str(master_tmp))
        service.save_master_to_local(df, str(master_tmp))

        try:
            organizations_bytes = _fetch_bytes(REMOTE_ORGANIZATIONS_URL, timeout=15)
            organizations_file = DATA_DIR / "reference" / "organizations.json"
            organizations_file.parent.mkdir(parents=True, exist_ok=True)
            organizations_file.write_bytes(organizations_bytes)
        except Exception as exc:
            # 기관 목록도 부가 항목이라 실패해도 마스터 업데이트 자체는 그대로 성공 처리
            logger.info("apply_update: 기관 목록 업데이트 건너뜀 - %s", exc)

        try:
            remote_projects = json.loads(_fetch_bytes(REMOTE_PROJECTS_URL, timeout=15).decode("utf-8"))
            _merge_remote_projects(remote_projects)
        except Exception as exc:
            # 사업 목록은 부가 항목이라 실패해도 마스터/기관 업데이트 자체는 그대로 성공 처리
            logger.info("apply_update: 사업 목록 업데이트 건너뜀 - %s", exc)

    _write_local_version(str(remote_info.get("version", "")))


def _merge_remote_projects(remote_projects: list[dict[str, Any]]) -> None:
    projects_file = DATA_DIR / "reference" / "upload_projects.json"
    try:
        with projects_file.open("r", encoding="utf-8") as f:
            local_projects = json.load(f)
        if not isinstance(local_projects, list):
            local_projects = []
    except Exception:
        local_projects = []

    existing_keys = {(p.get("사업년도"), p.get("순번")) for p in local_projects}
    added = 0
    for project in remote_projects:
        key = (project.get("사업년도"), project.get("순번"))
        if key in existing_keys:
            continue
        local_projects.append(project)
        existing_keys.add(key)
        added += 1

    if added:
        projects_file.parent.mkdir(parents=True, exist_ok=True)
        with projects_file.open("w", encoding="utf-8") as f:
            json.dump(local_projects, f, ensure_ascii=False, indent=2)
    logger.info("apply_update: 사업 목록 %d건 새로 병합", added)
