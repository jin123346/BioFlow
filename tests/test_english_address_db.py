import sqlite3
from contextlib import closing
from pathlib import Path

import pytest

from services.english_address_service import AddressDatabaseError, EnglishAddressService, english_fields
from services.vworld_geocoding_service import VWorldAddress
from tools.build_address_db import decode_cells, normalize_code, rows, split_district


@pytest.fixture
def reference(tmp_path):
    path = tmp_path / "reference.sqlite"
    with sqlite3.connect(path) as db:
        db.execute("CREATE TABLE admin_names(code TEXT PRIMARY KEY,sido TEXT,city TEXT,district TEXT,town TEXT,ri TEXT) WITHOUT ROWID")
        db.execute("INSERT INTO admin_names VALUES(?,?,?,?,?,?)", ("0012345678", "Gyeonggi-do", "Anyang-si", "Dongan-gu", "Gwanyang-dong", ""))
        db.execute("PRAGMA user_version=1")
    return EnglishAddressService(path)


def test_lookup_exact_and_read_only(reference):
    reference.validate()
    result = reference.lookup("0012345678")
    assert (result.city, result.district) == ("Anyang-si", "Dongan-gu")
    assert reference.lookup("9999999999") is None
    assert reference.lookup(12345678) is None
    with closing(reference.connect()) as db:
        with pytest.raises(sqlite3.OperationalError):
            db.execute("DELETE FROM admin_names")
        plan = db.execute("EXPLAIN QUERY PLAN SELECT * FROM admin_names WHERE code=?", ("0012345678",)).fetchone()
        assert "PRIMARY KEY" in plan[-1]


def test_missing_corrupt_and_schema_failures(tmp_path):
    service = EnglishAddressService(tmp_path / "missing.sqlite")
    with pytest.raises(AddressDatabaseError):
        service.validate()
    assert not service.path.exists()
    service.path.write_text("not a database")
    with pytest.raises(AddressDatabaseError):
        service.validate()


def test_ambiguous_code_is_not_selected(tmp_path):
    path = tmp_path / "ambiguous.sqlite"
    with sqlite3.connect(path) as db:
        db.execute("CREATE TABLE admin_names(code TEXT,sido TEXT,city TEXT,district TEXT,town TEXT,ri TEXT,PRIMARY KEY(code,sido,city,district,town,ri)) WITHOUT ROWID")
        db.executemany("INSERT INTO admin_names VALUES(?,?,?,?,?,?)", [
            ("0012345678", "Province", "", "County", "Town", "First-ri"),
            ("0012345678", "Province", "", "County", "Town", "Second-ri"),
        ])
        db.execute("PRAGMA user_version=1")
    with pytest.raises(AddressDatabaseError, match="서로 다른 영문명"):
        EnglishAddressService(path).lookup("0012345678")


def test_builder_parsing_and_preservation():
    assert normalize_code("0012345678") == "0012345678"
    assert normalize_code("4.1173102E+9") == "4117310200"
    with pytest.raises(ValueError):
        normalize_code("123.5")
    assert split_district("Dongan-gu, Anyang-si") == ("Anyang-si", "Dongan-gu")
    assert split_district("Namhae-gun") == ("", "Namhae-gun")
    with pytest.raises(ValueError):
        split_district("Unexpected, Two")
    import io
    xml = b'<worksheet><row r="1"><c r="B1" t="s"><v>0</v></c><c r="A1"><v>1</v></c></row></worksheet>'
    parsed = list(rows(io.BytesIO(xml)))
    assert decode_cells(parsed[0], ["0012345678"], {"B"}) == {"B": "0012345678"}


def test_english_mapping_uses_only_structural_names(reference):
    result = reference.lookup("0012345678")
    address = VWorldAddress({"level5": "1590"}, "", "37", "127")
    fields = english_fields(address, result)
    assert fields["관찰위치 시군구 영문"] == "Dongan-gu, Anyang-si"
    assert fields["관찰위치 상세 영문"] == "1590"
    address.structure["level5"] = "개인 현장 기록"
    assert english_fields(address, result)["관찰위치 상세 영문"] == ""


def test_mountain_parcel_and_ri_cover_the_same_detail(reference):
    from services.address_mapping_service import map_address
    address = VWorldAddress(
        {"level1":"경상북도", "level2":"안동시", "level4L":"풍천면",
         "level4LC":"0012345678", "level5":"산128임", "detail":""},
        "경상북도 안동시 풍천면 광덕리 산 128", "36.5264", "128.52435")
    korean, metadata = map_address(address)
    assert korean["관찰위치 상세 국문"] == "광덕리 산 128"
    assert metadata["ri_korean"] == "광덕리"
    english = english_fields(address, reference.lookup("0012345678"))
    assert english["관찰위치 상세 영문"] == "San 128"


def test_build_all_sheets_deduplicates_and_rejects_conflicts(tmp_path):
    import json
    import zipfile
    from tools.build_address_db import build
    source = tmp_path / "source.xlsx"
    destination = tmp_path / "ref.sqlite"
    report = tmp_path / "report.json"
    def workbook(second_name):
        strings = ['법정동코드','영문시도명','영문시군구명','영문읍면동명','영문리명',
                   '0012345678','Province','City-si','Town-dong',second_name]
        with zipfile.ZipFile(source, 'w') as z:
            z.writestr('xl/sharedStrings.xml', '<sst>' + ''.join(f'<si><t>{s}</t></si>' for s in strings) + '</sst>')
            z.writestr('xl/workbook.xml', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="one" r:id="rId1"/><sheet name="two" r:id="rId2"/></sheets></workbook>')
            z.writestr('xl/_rels/workbook.xml.rels', '<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Target="worksheets/sheet2.xml"/></Relationships>')
            for sheet in (1,2):
                header = '<row r="1">' + ''.join(f'<c r="{col}1" t="s"><v>{i}</v></c>' for i,col in enumerate('BCDEF')) + '</row>'
                values = [5,6,7,8 if sheet == 1 else 9]
                row = '<row r="2">' + ''.join(f'<c r="{col}2" t="s"><v>{i}</v></c>' for col,i in zip('BCDE',values)) + '</row>'
                z.writestr(f'xl/worksheets/sheet{sheet}.xml','<worksheet>'+header+row+'</worksheet>')
    workbook('Town-dong')
    build(source,destination,report)
    summary = json.loads(report.read_text(encoding='utf-8'))
    assert summary['source_rows'] == 2
    assert summary['unique_codes'] == 1
    assert summary['duplicate_rows_removed'] == 1
    assert len(summary['sheets']) == 2
    previous = destination.read_bytes()
    workbook('Different-dong')
    with pytest.raises(ValueError, match='conflicting codes'):
        build(source,destination,report)
    assert destination.read_bytes() == previous
    assert len(json.loads(report.read_text(encoding='utf-8'))['conflicts']) == 1


def test_apply_english_only_to_matching_korean_and_preserve_manual(reference):
    from services.address_mapping_service import apply_address
    result = reference.lookup('0012345678')
    address = VWorldAddress({'level1':'경기도','level2':'안양시 동안구','level4L':'관양동','level5':'1590'},'', '37','127')
    updated = apply_address({'관찰위치 시도 영문':'Manual'}, address, reference=result)
    assert updated['관찰위치 시도 영문'] == 'Manual'
    assert updated['관찰위치 시군구 영문'] == 'Dongan-gu, Anyang-si'
    mismatch = apply_address({'관찰위치 시도 국문':'서울특별시'}, address, reference=result)
    assert '관찰위치 시군구 영문' not in mismatch
    overwrite = apply_address({'관찰위치 시도 영문':'Manual'}, address, reference=result, overwrite=True)
    assert overwrite['관찰위치 시도 영문'] == 'Gyeonggi-do'
