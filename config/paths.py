import shutil
import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
FROZEN = bool(getattr(sys, "frozen", False))
EXECUTABLE_DIR = Path(sys.executable).resolve().parent if FROZEN else BASE_DIR

# PyInstaller로 빌드한 exe는 Program Files 등 쓰기 금지 위치에 설치될 수 있어
# 항상 내 문서\BioFlow에 데이터를 두고 씀. exe에 번들된 data/ 폴더는 최초 실행 시에만
# 시드로 복사(ensure_data_seeded, app.py에서 앱 시작 전 호출)하고 이후엔 읽지 않음.
APP_DIR = Path.home() / "Documents" / "BioFlow"
BUNDLED_DATA_DIR = Path(getattr(sys, "_MEIPASS", BASE_DIR)) / "data"

# ZIP을 푼 위치마다 작성 중 데이터와 사용자 설정을 따로 보관한다. 기준자료는 APP_DIR에서
# 버전 간 공유하지만, 임시저장은 다른 BioFlow EXE가 함께 읽지 않도록 실행 파일 옆에 둔다.
INSTANCE_DIR = (EXECUTABLE_DIR / "user_data") if FROZEN else BASE_DIR
DRAFT_DIR = INSTANCE_DIR / "drafts" if FROZEN else BASE_DIR / "data" / "reference"
SETTINGS_DIR = INSTANCE_DIR / "settings" if FROZEN else BASE_DIR / "data" / "reference"

# 아이콘 등 앱과 함께 배포되는 고정 리소스 (사용자 데이터 아님, 읽기 전용)
ASSETS_DIR = Path(getattr(sys, "_MEIPASS", BASE_DIR)) / "assets"

# 지금 실행 중인 앱(exe) 자체의 버전 (빌드 시 BioFlow.spec이 datas로 같이 번들함).
# data/version.json(종정보 마스터 버전)과는 별개.
APP_VERSION_FILE = Path(getattr(sys, "_MEIPASS", BASE_DIR)) / "VERSION"


def get_app_version() -> str:
    try:
        return APP_VERSION_FILE.read_text(encoding="utf-8").strip()
    except Exception:
        return ""

DATA_DIR = (APP_DIR / "data") if FROZEN else (BASE_DIR / "data")
CACHE_DIR = INSTANCE_DIR / "cache" if FROZEN else BASE_DIR / "cache"
LOG_DIR = INSTANCE_DIR / "logs" if FROZEN else BASE_DIR / "logs"

MASTER_DIR = DATA_DIR / 'master'
META_DIR= DATA_DIR / 'meta'
HISTORY_DIR= MASTER_DIR / 'history'

MASTER_FILE_PATH= MASTER_DIR / 'species_master.xlsx'
MASTER_DB_PATH = MASTER_DIR / 'species_master.sqlite'
LOCAL_VERSION_FILE_PATH= META_DIR / 'local_version.json'
LOCAL_VERSION_HISTORY_FILE_PATH = META_DIR / "local_version_history.json"
MASTER_CACHE_FILE = CACHE_DIR / "master.pkl"

# 새 사용자에게 넘기면 안 되는 로컬 작성 상태. 빌드는 기준자료 허용 목록만 포함하지만,
# 예전 배포본이나 수동 패키징에도 사용자 상태가 섞이지 않도록 시드 단계에서 한 번 더 막는다.
_SEED_EXCLUDE_NAMES = {
    "new_data_draft.json",
    "new_specimen_draft.json",
    "excel_column_mappings.json",
    "sql_media_mappings.json",
    "local_version.json",
    "local_version_history.json",
    "remote_data_version.json",
}
_SEED_EXCLUDE_DIRS = {"history", "cache", "logs"}


def ensure_data_seeded():
    """최초 실행(내 문서\\BioFlow\\data가 아직 없음)이면 exe에 번들된 data/를 복사해둠."""
    if not FROZEN or not BUNDLED_DATA_DIR.exists():
        return
    if DATA_DIR.exists():
        # Existing installs retain user data while receiving newly bundled references.
        for relative in (Path("classification"), Path("master") / "species_master.sqlite"):
            source = BUNDLED_DATA_DIR / relative
            target = DATA_DIR / relative
            if source.is_dir() and not target.exists():
                shutil.copytree(source, target)
            elif source.is_file() and not target.exists():
                target.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(source, target)
        return

    def _skip_user_state(_src_dir: str, names: list[str]) -> set[str]:
        return {
            name for name in names
            if name in _SEED_EXCLUDE_NAMES or name in _SEED_EXCLUDE_DIRS
        }

    DATA_DIR.parent.mkdir(parents=True, exist_ok=True)
    shutil.copytree(BUNDLED_DATA_DIR, DATA_DIR, ignore=_skip_user_state)
