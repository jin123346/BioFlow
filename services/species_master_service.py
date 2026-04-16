import json
import shutil
from datetime import datetime
from pathlib import Path

import pandas as pd

from config.paths import (
    MASTER_FILE_PATH,
    MASTER_CACHE_FILE,
    MASTER_DIR,
    META_DIR,
    HISTORY_DIR,
    CACHE_DIR,
    LOCAL_VERSION_FILE_PATH,
    LOCAL_VERSION_HISTORY_FILE_PATH,
)
from utils.text_utils import normalize_text

class SpeciesMasterService:
    def load_excel_file(self, file_path: str) -> pd.DataFrame:
        df = pd.read_excel(file_path, header=1)

        clean_columns = []
        for col in df.columns:
            if isinstance(col, (int, float)):
                continue
            clean_columns.append(str(col).strip().replace("\n", " "))

        df = df.iloc[:, :len(clean_columns)]
        df.columns = clean_columns
        return df

    def load_master_from_excel(self) -> pd.DataFrame:
        if not MASTER_FILE_PATH.exists():
            raise FileNotFoundError(f"마스터 파일이 없습니다: {MASTER_FILE_PATH}")

        df = pd.read_excel(MASTER_FILE_PATH, header=1)

        CACHE_DIR.mkdir(parents=True, exist_ok=True)
        df.to_pickle(MASTER_CACHE_FILE)
        return df

    def load_last_master(self) -> pd.DataFrame:
        if MASTER_CACHE_FILE.exists():
            return pd.read_pickle(MASTER_CACHE_FILE)

        if MASTER_FILE_PATH.exists():
            return self.load_master_from_excel()

        raise FileNotFoundError("저장된 종정보 파일이 없습니다.")

    def reload_master(self) -> pd.DataFrame:
        if not MASTER_FILE_PATH.exists():
            raise FileNotFoundError("재로딩할 원본 파일이 없습니다.")
        return self.load_master_from_excel()

    def save_master_to_local(self, df: pd.DataFrame, source_file_path: str) -> dict:
        if df is None or df.empty:
            raise ValueError("저장할 데이터가 없습니다.")

        if not source_file_path:
            raise ValueError("저장할 파일 경로가 없습니다.")

        MASTER_DIR.mkdir(parents=True, exist_ok=True)
        META_DIR.mkdir(parents=True, exist_ok=True)
        HISTORY_DIR.mkdir(parents=True, exist_ok=True)
        CACHE_DIR.mkdir(parents=True, exist_ok=True)

        shutil.copy2(source_file_path, MASTER_FILE_PATH)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_file_path = HISTORY_DIR / f"species_master_backup_{timestamp}.xlsx"
        shutil.copy2(source_file_path, backup_file_path)

        df.to_pickle(MASTER_CACHE_FILE)

        version_data = {
            "version": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "source": "local_upload",
            "update_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "file_name": Path(source_file_path).name,
            "saved_path": str(MASTER_FILE_PATH),
            "backup_path": str(backup_file_path),
        }

        with open(LOCAL_VERSION_FILE_PATH, "w", encoding="utf-8") as f:
            json.dump(version_data, f, ensure_ascii=False, indent=4)

        history_data = []
        if LOCAL_VERSION_HISTORY_FILE_PATH.exists():
            try:
                with open(LOCAL_VERSION_HISTORY_FILE_PATH, "r", encoding="utf-8") as f:
                    history_data = json.load(f)
                if not isinstance(history_data, list):
                    history_data = []
            except Exception:
                history_data = []

        history_data.append(version_data)

        with open(LOCAL_VERSION_HISTORY_FILE_PATH, "w", encoding="utf-8") as f:
            json.dump(history_data, f, ensure_ascii=False, indent=4)

        return {
            "master_path": str(MASTER_FILE_PATH),
            "backup_path": str(backup_file_path),
            "version_path": str(LOCAL_VERSION_FILE_PATH),
        }