"""SQLite-backed species reference with atomic Excel import and indexed lookup."""
from contextlib import closing
from datetime import datetime
import json, os, shutil, sqlite3
from pathlib import Path
import pandas as pd
from config.paths import HISTORY_DIR, LOCAL_VERSION_FILE_PATH, LOCAL_VERSION_HISTORY_FILE_PATH, MASTER_DB_PATH, MASTER_DIR
from utils.text_utils import normalize_text

SCHEMA_VERSION = 1
LOOKUP_COLUMNS = ("종학명정보ID", "종고유ID", "국명", "학명", "속명국문명", "종소명국문명", "분류체계명(국문)", "분류체계명(라틴)")

class SpeciesDatabaseError(Exception): pass
def quote(name): return '"' + str(name).replace('"', '""') + '"'
def cell_text(value):
    if value is None or pd.isna(value): return ""
    if isinstance(value, float) and value.is_integer(): return str(int(value))
    return str(value).strip()

class SpeciesMasterService:
    def __init__(self, path=None): self.path = Path(path) if path is not None else MASTER_DB_PATH

    @staticmethod
    def master_fingerprint():
        try:
            stat = MASTER_DB_PATH.stat(); return stat.st_mtime_ns, stat.st_size
        except OSError: return None

    def exists(self): return self.path.is_file()

    def connect(self, path=None, writable=False):
        target = Path(path) if path is not None else self.path
        if not writable and not target.is_file(): raise FileNotFoundError(f"종정보 DB가 없습니다: {target}")
        try:
            if writable: return sqlite3.connect(target)
            db = sqlite3.connect(target.resolve().as_uri() + "?mode=ro", uri=True)
            db.execute("PRAGMA query_only=ON"); db.execute("PRAGMA cache_size=-1024")
            return db
        except sqlite3.Error: raise SpeciesDatabaseError("종정보 DB를 열 수 없습니다.") from None

    def columns(self, db=None):
        if db is None:
            with closing(self.connect()) as connection: return self.columns(connection)
        return [row[1] for row in db.execute("PRAGMA table_info(species)") if row[1] != "_row"]

    def validate_database(self, path=None, expected_rows=None, expected_sha256=None):
        target = Path(path) if path is not None else self.path
        if expected_sha256:
            import hashlib
            if hashlib.sha256(target.read_bytes()).hexdigest().lower() != str(expected_sha256).lower():
                raise SpeciesDatabaseError("종정보 DB 해시가 일치하지 않습니다.")
        try:
            with closing(self.connect(target)) as db:
                if db.execute("PRAGMA integrity_check").fetchone()[0] != "ok": raise SpeciesDatabaseError("종정보 DB 무결성 검사에 실패했습니다.")
                if db.execute("PRAGMA user_version").fetchone()[0] != SCHEMA_VERSION: raise SpeciesDatabaseError("지원하지 않는 종정보 DB 버전입니다.")
                missing = set(LOOKUP_COLUMNS) - set(self.columns(db))
                if missing: raise SpeciesDatabaseError(f"종정보 DB 필수 컬럼이 없습니다: {', '.join(sorted(missing))}")
                count = db.execute("SELECT COUNT(*) FROM species").fetchone()[0]
                if expected_rows is not None and count != int(expected_rows): raise SpeciesDatabaseError("종정보 DB 행 수가 배포 정보와 일치하지 않습니다.")
                duplicate = db.execute(f"SELECT {quote('종고유ID')} FROM species WHERE {quote('종고유ID')} NOT IN ('','0') GROUP BY {quote('종고유ID')} HAVING COUNT(*)>1 LIMIT 1").fetchone()
                if duplicate: raise SpeciesDatabaseError(f"중복 종고유ID가 있습니다: {duplicate[0]}")
                return count
        except sqlite3.Error as exc: raise SpeciesDatabaseError(f"종정보 DB 구조가 올바르지 않습니다: {exc}") from None

    def load_excel_file(self, file_path):
        df = pd.read_excel(file_path, header=1)
        valid = [column for column in df.columns if not isinstance(column, (int, float))]
        df = df.iloc[:, :len(valid)].copy(); df.columns = [str(column).strip().replace("\n", " ") for column in valid]
        missing = set(LOOKUP_COLUMNS) - set(df.columns)
        if missing: raise ValueError(f"필수 종정보 컬럼이 없습니다: {', '.join(sorted(missing))}")
        return df

    def build_database(self, dataframe, destination):
        if dataframe is None or dataframe.empty: raise ValueError("저장할 종정보가 없습니다.")
        columns = [str(column).strip().replace("\n", " ") for column in dataframe.columns]
        if len(columns) != len(set(columns)): raise ValueError("종정보 컬럼명이 중복됩니다.")
        missing = set(LOOKUP_COLUMNS) - set(columns)
        if missing: raise ValueError(f"필수 종정보 컬럼이 없습니다: {', '.join(sorted(missing))}")
        destination = Path(destination); destination.parent.mkdir(parents=True, exist_ok=True)
        if destination.exists(): destination.unlink()
        with closing(self.connect(destination, writable=True)) as db:
            db.execute("PRAGMA journal_mode=OFF"); db.execute("PRAGMA synchronous=OFF")
            db.execute(f"CREATE TABLE species (_row INTEGER PRIMARY KEY,{','.join(f'{quote(c)} TEXT NOT NULL' for c in columns)})")
            insert = f"INSERT INTO species ({','.join(map(quote, columns))}) VALUES ({','.join('?' for _ in columns)})"
            batch = []
            for values in dataframe.itertuples(index=False, name=None):
                batch.append(tuple(cell_text(value) for value in values))
                if len(batch) == 1000: db.executemany(insert, batch); batch.clear()
            if batch: db.executemany(insert, batch)
            db.execute(f"CREATE INDEX ix_species_korean ON species({quote('국명')})")
            db.execute(f"CREATE INDEX ix_species_scientific ON species({quote('학명')})")
            db.execute("CREATE TABLE metadata (key TEXT PRIMARY KEY,value TEXT NOT NULL) WITHOUT ROWID")
            db.executemany("INSERT INTO metadata VALUES (?,?)", [("created_at", datetime.now().isoformat(timespec="seconds")), ("row_count", str(len(dataframe)))])
            db.execute(f"PRAGMA user_version={SCHEMA_VERSION}"); db.commit(); db.execute("VACUUM")
        self.validate_database(destination, expected_rows=len(dataframe)); return destination

    def install_database(self, source, expected_rows=None, expected_sha256=None, keep_backup=True):
        source = Path(source); self.validate_database(source, expected_rows, expected_sha256)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        temporary = self.path.with_suffix(".installing.sqlite"); shutil.copy2(source, temporary)
        self.validate_database(temporary, expected_rows, expected_sha256)
        backup = self.path.with_suffix(".sqlite.bak")
        if self.path.exists() and keep_backup: shutil.copy2(self.path, backup)
        try:
            os.replace(temporary, self.path)
            self.validate_database(expected_rows=expected_rows, expected_sha256=expected_sha256)
        except Exception:
            if keep_backup and backup.is_file():
                shutil.copy2(backup, self.path)
            raise
        finally:
            temporary.unlink(missing_ok=True)
        return backup if backup.exists() else None

    def save_master_to_local(self, df, source_file_path=""):
        MASTER_DIR.mkdir(parents=True, exist_ok=True); HISTORY_DIR.mkdir(parents=True, exist_ok=True)
        temporary = MASTER_DIR / "species_master.importing.sqlite"; self.build_database(df, temporary)
        backup = self.install_database(temporary)
        if temporary.exists(): temporary.unlink()
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        version_data = {"version": timestamp, "source": "local_upload", "update_at": timestamp, "file_name": Path(source_file_path).name if source_file_path else "", "saved_path": str(self.path), "backup_path": str(backup or "")}
        LOCAL_VERSION_FILE_PATH.parent.mkdir(parents=True, exist_ok=True)
        LOCAL_VERSION_FILE_PATH.write_text(json.dumps(version_data, ensure_ascii=False, indent=2), encoding="utf-8")
        try:
            history = json.loads(LOCAL_VERSION_HISTORY_FILE_PATH.read_text(encoding="utf-8")); history = history if isinstance(history, list) else []
        except Exception: history = []
        history.append(version_data); LOCAL_VERSION_HISTORY_FILE_PATH.write_text(json.dumps(history, ensure_ascii=False, indent=2), encoding="utf-8")
        return {"master_path": str(self.path), "backup_path": str(backup or ""), "version_path": str(LOCAL_VERSION_FILE_PATH)}

    def load_last_master(self):
        with closing(self.connect()) as db: return pd.read_sql_query("SELECT * FROM species ORDER BY _row", db).drop(columns=["_row"])
    def reload_master(self): return self.load_last_master()
    def preview(self, limit=200):
        with closing(self.connect()) as db:
            total = db.execute("SELECT COUNT(*) FROM species").fetchone()[0]
            frame = pd.read_sql_query("SELECT * FROM species ORDER BY _row LIMIT ?", db, params=(int(limit),)).drop(columns=["_row"])
            return frame, total

    def find_matches(self, korean="", scientific_key="", prefer_scientific=False):
        korean, scientific_key = normalize_text(korean), normalize_text(scientific_key).lower()
        if not korean and not scientific_key: return []
        with closing(self.connect()) as db:
            db.row_factory = sqlite3.Row
            if prefer_scientific or not korean:
                rows = [dict(row) for row in db.execute(f"SELECT * FROM species WHERE lower({quote('학명')}) LIKE ? ORDER BY _row LIMIT 50", (scientific_key + "%",)).fetchall()]
                return [row for row in rows if " ".join(normalize_text(row.get("학명")).split()[:2]).lower() == scientific_key]
            return [dict(row) for row in db.execute(f"SELECT * FROM species WHERE {quote('국명')}=? ORDER BY _row LIMIT 50", (korean,)).fetchall()]

    SEARCH_COLUMNS = {"학명": ["학명"], "국명": ["국명"], "분류군": ["분류체계명(국문)", "분류체계명(라틴)"]}
    def search_database(self, filters, limit=None):
        clauses, params = [], []
        for search_type, keyword in filters.items():
            keyword = normalize_text(keyword).lower()
            if not keyword: continue
            fields = self.SEARCH_COLUMNS.get(search_type, [])
            clauses.append("(" + " OR ".join(f"lower({quote(f)}) LIKE ?" for f in fields) + ")"); params.extend([f"%{keyword}%"] * len(fields))
        if not clauses: return pd.DataFrame()
        sql = "SELECT * FROM species WHERE " + " AND ".join(clauses) + " ORDER BY _row"
        if limit is not None: sql += " LIMIT ?"; params.append(int(limit))
        with closing(self.connect()) as db: return pd.read_sql_query(sql, db, params=params).drop(columns=["_row"])

    def search(self, df, filters):
        if df is None: return pd.DataFrame()
        mask, has_condition = pd.Series(True, index=df.index), False
        for search_type, keyword in filters.items():
            keyword = normalize_text(keyword).lower()
            if not keyword: continue
            has_condition = True; group = pd.Series(False, index=df.index)
            for field in [f for f in self.SEARCH_COLUMNS.get(search_type, []) if f in df.columns]: group |= df[field].apply(lambda value: keyword in normalize_text(value).lower())
            mask &= group
        return df[mask] if has_condition else df.iloc[0:0]
