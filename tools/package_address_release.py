"""Package and compare isolated pre/post PyInstaller builds."""
import argparse
import hashlib
import io
import json
from pathlib import Path
import zipfile


def measure_build(folder, archive):
    files = sorted(p for p in folder.rglob("*") if p.is_file())
    archive.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(archive, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        for path in files:
            z.write(path, str(Path(folder.name) / path.relative_to(folder)))
    return {"directory": str(folder), "file_count": len(files),
            "uncompressed_bytes": sum(p.stat().st_size for p in files),
            "exe_bytes": (folder / "BioFlow.exe").stat().st_size,
            "zip": str(archive), "zip_bytes": archive.stat().st_size}


def package(baseline, optimized, report):
    db = optimized / "data/address/english_admin.sqlite"
    if not db.is_file():
        raise ValueError("Missing external address database")
    if list(optimized.rglob("*new.xlsx")) or list(optimized.rglob("영문주소_정리.xlsx")):
        raise ValueError("Address source workbook must not be distributed")
    if len(list(optimized.rglob("english_admin.sqlite"))) != 1:
        raise ValueError("Expected one external database copy")
    db_zip = io.BytesIO()
    with zipfile.ZipFile(db_zip, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        z.write(db, "data/address/english_admin.sqlite")
    data = {"baseline": measure_build(baseline, baseline.parent / (baseline.name + '.zip')),
            "optimized": measure_build(optimized, optimized.parent / (optimized.name + '.zip')),
            "database_bytes": db.stat().st_size, "database_zip_bytes": len(db_zip.getvalue()),
            "database_sha256": hashlib.sha256(db.read_bytes()).hexdigest()}
    data["delta_uncompressed_bytes"] = data["optimized"]["uncompressed_bytes"] - data["baseline"]["uncompressed_bytes"]
    data["delta_zip_bytes"] = data["optimized"]["zip_bytes"] - data["baseline"]["zip_bytes"]
    report.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(data, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("baseline", type=Path)
    parser.add_argument("optimized", type=Path)
    parser.add_argument("--report", type=Path, default=Path("docs/address-db-packaging.json"))
    args = parser.parse_args()
    package(args.baseline, args.optimized, args.report)
