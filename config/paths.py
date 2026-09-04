import shutil
import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
FROZEN = bool(getattr(sys, "frozen", False))

# PyInstaller로 빌드한 exe는 Program Files 등 쓰기 금지 위치에 설치될 수 있어
# 항상 내 문서\BioFlow에 데이터를 두고 씀. exe에 번들된 data/ 폴더는 최초 실행 시에만
# 시드로 복사(ensure_data_seeded, app.py에서 앱 시작 전 호출)하고 이후엔 읽지 않음.
APP_DIR = Path.home() / "Documents" / "BioFlow"
BUNDLED_DATA_DIR = Path(getattr(sys, "_MEIPASS", BASE_DIR)) / "data"

# 아이콘 등 앱과 함께 배포되는 고정 리소스 (사용자 데이터 아님, 읽기 전용)
ASSETS_DIR = Path(getattr(sys, "_MEIPASS", BASE_DIR)) / "assets"

DATA_DIR = (APP_DIR / "data") if FROZEN else (BASE_DIR / "data")
CACHE_DIR = (APP_DIR / "cache") if FROZEN else (BASE_DIR / "cache")
LOG_DIR = (APP_DIR / "logs") if FROZEN else (BASE_DIR / "logs")

MASTER_DIR = DATA_DIR / 'master'
META_DIR= DATA_DIR / 'meta'
HISTORY_DIR= MASTER_DIR / 'history'

MASTER_FILE_PATH= MASTER_DIR / 'species_master.xlsx'
LOCAL_VERSION_FILE_PATH= META_DIR / 'local_version.json'
LOCAL_VERSION_HISTORY_FILE_PATH = META_DIR / "local_version_history.json"
MASTER_CACHE_FILE = CACHE_DIR / "master.pkl"

# 새 사용자에게 넘길 필요 없는(작성 중이던 테스트 데이터) 파일은 시드에서 제외
_SEED_EXCLUDE_NAMES = {"new_data_draft.json", "new_specimen_draft.json"}


def ensure_data_seeded():
    """최초 실행(내 문서\\BioFlow\\data가 아직 없음)이면 exe에 번들된 data/를 복사해둠."""
    if not FROZEN or DATA_DIR.exists() or not BUNDLED_DATA_DIR.exists():
        return

    def _skip_drafts(_src_dir: str, names: list[str]) -> set[str]:
        return {name for name in names if name in _SEED_EXCLUDE_NAMES}

    DATA_DIR.parent.mkdir(parents=True, exist_ok=True)
    shutil.copytree(BUNDLED_DATA_DIR, DATA_DIR, ignore=_skip_drafts)