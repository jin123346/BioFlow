"""Build a compact, audited internal reference from all regional workbook sheets.

Only the six administrative columns are decoded. Row XML is streamed from ZIP;
neither building records nor the expanded multi-GB workbook are retained.
"""
import argparse
from collections import Counter, defaultdict
from contextlib import closing
from decimal import Decimal, InvalidOperation
import hashlib
import json
from pathlib import Path
import re
import sqlite3
import xml.etree.ElementTree as ET
import zipfile

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
HEADERS = ["법정동코드", "영문시도명", "영문시군구명", "영문읍면동명", "영문리명"]
ROW = re.compile(rb"<row\b[^>]*>.*?</row>", re.S)
CELL = re.compile(rb'<c\b([^>]*\br="([A-Z]+)[0-9]+"[^>]*)>(.*?)</c>', re.S)
VALUE = re.compile(rb"<v>(.*?)</v>", re.S)
ADMIN_CELL = re.compile(rb'<c r="([B-F])[0-9]+"([^>]*)>(.*?)</c>', re.S)


def rows(stream):
    pending = b""
    while True:
        block = stream.read(1024 * 1024)
        if not block:
            if b"<row " in pending:
                raise ValueError("Incomplete worksheet row")
            break
        pending += block
        end = pending.rfind(b"</row>")
        if end < 0:
            continue
        end += len(b"</row>")
        for match in ROW.finditer(pending[:end]):
            yield match.group()
        pending = pending[end:]


def decode_cells(row, strings, selected=None):
    result = {}
    for match in CELL.finditer(row):
        attrs, column, content = match.groups()
        column = column.decode()
        if selected is not None and column not in selected:
            continue
        if b"<f" in content:
            raise ValueError("Formula in reference columns")
        value = VALUE.search(content)
        if b't="inlineStr"' in attrs:
            value = "".join(ET.fromstring(b"<root>" + content + b"</root>").itertext())
        elif value:
            value = value.group(1).decode()
            if b't="s"' in attrs:
                value = strings[int(value)]
            else:
                value = ET.fromstring("<v>" + value + "</v>").text or ""
        else:
            value = ""
        result[column] = value.strip()
    return result


def normalize_code(value):
    if re.fullmatch(r"[0-9]{10}", value):
        return value
    try:
        number = Decimal(value)
        if not number.is_finite() or number != number.to_integral_value():
            raise ValueError()
        code = format(number, "f").split(".")[0]
        if re.fullmatch(r"[0-9]{1,10}", code):
            return code.zfill(10)
    except InvalidOperation:
        pass
    raise ValueError(f"Invalid legal-dong code: {value!r}")


def split_district(value):
    parts = [p.strip() for p in value.split(",")]
    if len(parts) == 2 and parts[0].endswith("-gu") and parts[1].endswith("-si"):
        return parts[1], parts[0]
    if len(parts) != 1:
        raise ValueError(f"Unrecognized city/district combination: {value!r}")
    return "", value


def build(source, output, report_path, retain_conflicts=False):
    variants = defaultdict(Counter)
    examples = {}
    decoded = {}
    report = {"source": source.name, "source_bytes": source.stat().st_size, "sheets": [], "source_rows": 0}
    with zipfile.ZipFile(source) as archive:
        strings = ["".join(e.itertext()) for e in ET.fromstring(archive.read("xl/sharedStrings.xml"))]
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        rels = {e.attrib["Id"]: e.attrib["Target"] for e in ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))}
        for sheet in workbook.find(NS + "sheets"):
            name = sheet.attrib["name"]
            target = rels[sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]]
            path = target.lstrip("/") if target.startswith("/") else "xl/" + target
            count, sheet_codes = 0, set()
            with archive.open(path) as stream:
                iterator = rows(stream)
                header = decode_cells(next(iterator), strings)
                columns = {field: col for col, field in header.items() if field in HEADERS}
                if set(columns) != set(HEADERS):
                    raise ValueError(f"Unexpected columns: {name}: {header}")
                for row_number, raw in enumerate(iterator, 2):
                    # Source administrative cells repeat across thousands of buildings.
                    # Cache only their compact XML values, never building columns.
                    raw_key = tuple(match.groups() for match in ADMIN_CELL.finditer(raw)) if list(columns.values()) == list("BCDEF") else None
                    cached = decoded.get(raw_key) if raw_key else None
                    if cached is None:
                        values = decode_cells(raw, strings, set(columns.values()))
                        if not any(values.values()):
                            continue
                        code, sido, district, town, ri = [values.get(columns[field], "") for field in HEADERS]
                        code = normalize_code(code)
                        if not sido or not town:
                            raise ValueError(f"Missing administrative name: {name}:{row_number}")
                        city, district = split_district(district)
                        names = (sido, city, district, town, ri)
                        if raw_key:
                            decoded[raw_key] = (code, names)
                    else:
                        code, names = cached
                    variants[code][names] += 1
                    examples.setdefault((code, names), (name, row_number))
                    sheet_codes.add(code)
                    count += 1
            report["sheets"].append({"name": name, "rows": count, "unique_codes": len(sheet_codes)})
            report["source_rows"] += count
            print(f"{name}: {count:,} rows / {len(sheet_codes):,} codes", flush=True)
    conflicts = []
    for code, names in sorted(variants.items()):
        if len(names) > 1:
            conflicts.append({"code": code, "variants": [
                {"names": combo, "count": count, "first_source": examples[(code, combo)]}
                for combo, count in sorted(names.items())]})
    report.update(unique_codes=len(variants), unique_combinations=sum(map(len, variants.values())), conflicts=conflicts)
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    if conflicts and not retain_conflicts:
        raise ValueError(f"{len(conflicts)} conflicting codes; report saved, database NOT replaced")
    report["conflict_policy"] = "retain_all_variants_and_block_automatic_mapping"
    report["automatically_mappable_codes"] = len(variants) - len(conflicts)
    output.parent.mkdir(parents=True, exist_ok=True)
    temporary = output.with_suffix(".building.sqlite")
    with closing(sqlite3.connect(temporary)) as db:
        db.execute("DROP TABLE IF EXISTS admin_names")
        db.execute("CREATE TABLE admin_names (code TEXT NOT NULL, sido TEXT NOT NULL, city TEXT NOT NULL, district TEXT NOT NULL, town TEXT NOT NULL, ri TEXT NOT NULL, PRIMARY KEY(code,sido,city,district,town,ri)) WITHOUT ROWID")
        db.executemany("INSERT INTO admin_names VALUES (?,?,?,?,?,?)", [(code, *combo) for code, names in sorted(variants.items()) for combo in sorted(names)])
        db.execute("PRAGMA user_version=1")
        db.commit()
        db.execute("VACUUM")
        assert db.execute("PRAGMA integrity_check").fetchone()[0] == "ok"
    temporary.replace(output)
    report["database_bytes"] = output.stat().st_size
    report["duplicate_rows_removed"] = report["source_rows"] - report["unique_combinations"]
    report["database_sha256"] = hashlib.sha256(output.read_bytes()).hexdigest()
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({k:v for k,v in report.items() if k not in {"sheets", "conflicts"}}, ensure_ascii=False), flush=True)


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--source", type=Path, default=Path("data/address/영문주소정리_new.xlsx"))
    parser.add_argument("--output", type=Path, default=Path("data/address/english_admin.sqlite"))
    parser.add_argument("--report", type=Path, default=Path("docs/address-db-build.json"))
    parser.add_argument("--retain-conflicts", action="store_true", help="Retain all conflicting variants; runtime lookup refuses ambiguous codes")
    args = parser.parse_args()
    build(args.source, args.output, args.report, args.retain_conflicts)
