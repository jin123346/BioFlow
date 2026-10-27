import tempfile
import unittest
from pathlib import Path

from services.coordinate_service import apply_coordinates, parse_coordinates
from services.sorting_service import renumber_display_rows, sort_records


class SortingFeatureTests(unittest.TestCase):
    def test_stable_sort_keeps_empty_last_and_all_fields_together(self):
        rows = [
            {"국명": "참새", "파일경로": "c.jpg", "데이터베이스번호": "DB3"},
            {"국명": "까치", "파일경로": "a.jpg", "데이터베이스번호": "DB1"},
            {"국명": "", "파일경로": "blank.jpg", "데이터베이스번호": "DB0"},
            {"국명": "까치", "파일경로": "b.jpg", "데이터베이스번호": "DB2"},
        ]
        sorted_rows = sort_records(rows, "국명")
        self.assertEqual([row["파일경로"] for row in sorted_rows], ["a.jpg", "b.jpg", "c.jpg", "blank.jpg"])
        self.assertEqual([row["데이터베이스번호"] for row in sorted_rows], ["DB1", "DB2", "DB3", "DB0"])

    def test_display_renumber_does_not_change_database_or_project_sequence(self):
        rows = [{"NO.": 9, "사업별 순번": "77", "데이터베이스번호": "KEEP"}]
        result = renumber_display_rows(rows)
        self.assertEqual(result[0], {"NO.": 1, "사업별 순번": "77", "데이터베이스번호": "KEEP"})


class CoordinateFeatureTests(unittest.TestCase):
    def test_validation_and_selected_empty_only_application(self):
        latitude, longitude = parse_coordinates("37.5", "127.0")
        records, changed = apply_coordinates(
            [{"위도": "", "경도": ""}, {"위도": "35", "경도": "129"}, {"위도": "", "경도": ""}],
            [0, 1], latitude, longitude, only_empty=True,
        )
        self.assertEqual(changed, [0])
        self.assertEqual((records[0]["위도"], records[0]["경도"]), ("37.5", "127"))
        self.assertEqual((records[1]["위도"], records[1]["경도"]), ("35", "129"))

    def test_invalid_coordinates_are_rejected(self):
        for latitude, longitude in (("north", "127"), ("91", "127"), ("37", "181")):
            with self.subTest(latitude=latitude, longitude=longitude), self.assertRaises(ValueError):
                parse_coordinates(latitude, longitude)


if __name__ == "__main__":
    unittest.main()
