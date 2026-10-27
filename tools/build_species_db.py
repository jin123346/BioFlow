"""Convert an administrator-supplied species workbook to the release SQLite."""
import argparse, hashlib, json, sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from services.species_master_service import SpeciesMasterService

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--source", type=Path, default=Path("data/master/species_master.xlsx"))
    parser.add_argument("--output", type=Path, default=Path("data/master/species_master.sqlite"))
    parser.add_argument("--report", type=Path, default=Path("docs/species-db-build.json"))
    args = parser.parse_args()
    service = SpeciesMasterService(args.output)
    frame = service.load_excel_file(str(args.source))
    service.build_database(frame, args.output)
    count = service.validate_database()
    report = {"source": args.source.name, "source_bytes": args.source.stat().st_size,
              "rows": count, "columns": len(service.columns()), "database_bytes": args.output.stat().st_size,
              "sha256": hashlib.sha256(args.output.read_bytes()).hexdigest(), "schema_version": 1}
    args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False))
