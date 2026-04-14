from pathlib import Path
from datetime import datetime

BASE_DIR= Path(__file__).resolve().parent.parent
APP_DIR = Path.home() / "Documents" / "BioFlow"

DATA_DIR = BASE_DIR / 'data'
MASTER_DIR = DATA_DIR / 'master'
META_DIR= DATA_DIR / 'meta'
HISTORY_DIR= MASTER_DIR / 'history'
CACHE_DIR=BASE_DIR / 'cache'

MASTER_FILE_PATH= MASTER_DIR / 'species_master.xlsx'
LOCAL_VERSION_FILE_PATH= META_DIR / 'local_version.json'
LOCAL_VERSION_HISTORY_FILE_PATH = META_DIR / "local_version_history.json"
MASTER_CACHE_FILE = CACHE_DIR / f"master_{datetime.now().strftime('%Y%m%d')}.pkl"