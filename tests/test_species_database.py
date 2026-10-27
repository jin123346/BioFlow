import hashlib
import sqlite3
from pathlib import Path

import pandas as pd
import pytest

from services.species_master_service import LOOKUP_COLUMNS, SpeciesDatabaseError, SpeciesMasterService


def frame(*, korean="호랑이", scientific="Panthera tigris"):
    values = {column: "" for column in LOOKUP_COLUMNS}
    values.update({"종학명정보ID": "11", "종고유ID": "SP1", "국명": korean,
                   "학명": scientific, "속명국문명": "표범속", "종소명국문명": "호랑이"})
    return pd.DataFrame([values])


def test_build_lookup_preview_search_and_read_only(tmp_path):
    path = tmp_path / "species.sqlite"
    service = SpeciesMasterService(path)
    service.build_database(frame(), path)
    assert service.validate_database(expected_rows=1) == 1
    assert service.find_matches("호랑이")[0]["종고유ID"] == "SP1"
    assert service.find_matches("", "panthera tigris", True)[0]["국명"] == "호랑이"
    assert len(service.search_database({"분류군": "", "학명": "tigris"})) == 1
    preview, total = service.preview()
    assert total == 1 and list(preview["국명"]) == ["호랑이"]
    with service.connect() as db, pytest.raises(sqlite3.OperationalError):
        db.execute("DELETE FROM species")


def test_atomic_install_rejects_bad_file_and_preserves_current(tmp_path):
    active = tmp_path / "active.sqlite"
    service = SpeciesMasterService(active)
    service.build_database(frame(), active)
    before = active.read_bytes()
    bad = tmp_path / "bad.sqlite"
    bad.write_bytes(b"not sqlite")
    with pytest.raises(SpeciesDatabaseError):
        service.install_database(bad)
    assert active.read_bytes() == before


def test_atomic_install_validates_hash_rows_and_keeps_backup(tmp_path):
    active, incoming = tmp_path / "active.sqlite", tmp_path / "incoming.sqlite"
    service = SpeciesMasterService(active)
    service.build_database(frame(), active)
    SpeciesMasterService(incoming).build_database(frame(korean="여우", scientific="Vulpes vulpes"), incoming)
    digest = hashlib.sha256(incoming.read_bytes()).hexdigest()
    backup = service.install_database(incoming, expected_rows=1, expected_sha256=digest)
    assert backup and backup.exists()
    assert service.find_matches("여우")[0]["학명"] == "Vulpes vulpes"
    assert SpeciesMasterService(backup).find_matches("호랑이")[0]["학명"] == "Panthera tigris"
    with pytest.raises(SpeciesDatabaseError):
        service.install_database(incoming, expected_rows=2, expected_sha256=digest)


def test_duplicate_real_species_id_is_rejected_but_zero_is_allowed(tmp_path):
    duplicate = pd.concat([frame(), frame()], ignore_index=True)
    service = SpeciesMasterService(tmp_path / "dup.sqlite")
    with pytest.raises(SpeciesDatabaseError, match="중복 종고유ID"):
        service.build_database(duplicate, service.path)
    zeros = duplicate.copy()
    zeros["종고유ID"] = "0"
    service.build_database(zeros, service.path)
    assert service.validate_database() == 2
