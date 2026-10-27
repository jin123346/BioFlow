from contextlib import ExitStack
import json
import os
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import MagicMock, patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
from PySide6.QtWidgets import QApplication
from PIL import Image

from ui.tabs.upload_tab import ROW_COLUMNS, SPECIMEN_ROW_COLUMNS, UploadTab


class UploadWorkflowTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        self.stack = ExitStack()
        self.addCleanup(self.stack.close)
        self.callback_errors = []
        self.stack.enter_context(patch.object(sys, "excepthook", lambda kind, value, tb: self.callback_errors.append(str(value))))
        # Build real widgets/signals without loading or saving the user's drafts.
        for method in ("_ensure_reference_files", "load_reference_data", "load_species_master_lookup",
                       "load_draft", "load_specimen_draft", "schedule_autosave", "schedule_specimen_autosave"):
            self.stack.enter_context(patch.object(UploadTab, method))
        self.editor = UploadTab()
        self.addCleanup(self.editor.close)

    def seed(self, specimen):
        editor = self.editor
        table = editor.specimen_data_table if specimen else editor.data_table
        columns = SPECIMEN_ROW_COLUMNS if specimen else ROW_COLUMNS
        widgets = editor.specimen_form_widgets if specimen else editor.form_widgets
        write = editor.write_specimen_data_row if specimen else editor.write_data_row
        editor.autosave_suspended = editor.specimen_autosave_suspended = True
        for row, order in enumerate((1, 1, 2, 1, 1)):
            table.insertRow(row)
            write(row, {"NO.": row + 1, "사업년도": 2026, "순번": order, "사업별 순번": row + 1,
                        "기관코드": "TEST", "분류코드": "BI", "미디어타입": "이미지", "미디어코드": "IM",
                        "데이터베이스번호": f"OLD{row}"})
        editor.autosave_suspended = editor.specimen_autosave_suspended = False
        table.selectRow(1)
        widgets["사업별 순번"].setText("100")
        widgets["사업별 순번"].textEdited.emit("100")
        widgets["사업별 순번"].editingFinished.emit()
        return table, columns, widgets

    def test_real_row_selection_and_save_preserve_project_numbering(self):
        for specimen in (False, True):
            with self.subTest(specimen=specimen):
                table, columns, _ = self.seed(specimen)
                for _ in range(3):
                    table.selectRow(4)
                    table.selectRow(1)
                    self.editor.table_to_records(table, columns)
                values = lambda name: [table.item(row, columns.index(name)).text() for row in range(5)]
                self.assertEqual(values("사업별 순번"), ["1", "100", "3", "101", "102"])
                self.assertEqual(values("NO."), ["1", "2", "3", "4", "5"])
                self.assertEqual(values("순번"), ["1", "1", "2", "1", "1"])
                numbers = values("데이터베이스번호")
                self.assertEqual(numbers[2], "OLD2")
                for row, sequence in ((1, 100), (3, 101), (4, 102)):
                    self.assertTrue(numbers[row].endswith(f"{sequence:05d}"))
        self.assertEqual(self.callback_errors, [])

    def test_new_input_uses_next_number_and_does_not_overwrite_selected_row(self):
        for specimen in (False, True):
            with self.subTest(specimen=specimen):
                table, columns, widgets = self.seed(specimen)
                before = [[table.item(row, col).text() for col in range(len(columns))] for row in range(5)]
                (self.editor.clear_specimen_form if specimen else self.editor.clear_form)()
                self.assertEqual(table.selectedIndexes(), [])
                self.assertEqual(widgets["사업별 순번"].text(), "103")
                last = self.editor._specimen_last_selected_row if specimen else self.editor._last_selected_row
                self.assertIsNone(last)
                after = [[table.item(row, col).text() for col in range(len(columns))] for row in range(5)]
                self.assertEqual(before, after)
        self.assertEqual(self.callback_errors, [])

    def test_specimen_database_number_omits_sp_media_code(self):
        values = {
            "기관코드": "NSMK",
            "분류코드": "PL",
            "사업년도": 2026,
            "순번": 7,
            "사업별 순번": 5,
        }

        generated = self.editor.prepare_new_project_number(values, specimen=True)

        self.assertEqual(generated["데이터베이스번호"], "NSMK-PL-260700005")

    def test_observation_database_number_keeps_media_code(self):
        values = {
            "기관코드": "NSMK",
            "분류코드": "PL",
            "미디어타입": "이미지",
            "사업년도": 2026,
            "순번": 7,
            "사업별 순번": 5,
        }

        generated = self.editor.prepare_new_project_number(values, specimen=False)

        self.assertEqual(generated["데이터베이스번호"], "NSMK-OI-PL-260700005")

    def test_observation_missing_media_type_defaults_to_information_ob(self):
        generated = self.editor.prepare_new_project_number(
            {
                "기관코드": "NSMK", "분류코드": "PL", "미디어타입": "",
                "미디어코드": "", "사업년도": 2026, "순번": 7, "사업별 순번": 5,
            },
            specimen=False,
        )

        self.assertEqual(generated["미디어타입"], "정보")
        self.assertEqual(generated["미디어코드"], "OB")
        self.assertEqual(generated["데이터베이스번호"], "NSMK-OB-PL-260700005")

    def test_video_identifier_and_filename_follow_media_type_and_project_sequence(self):
        values = {
            "기관코드": "NSMK",
            "분류코드": "PL",
            "미디어타입": "동영상",
            "사업년도": 2026,
            "순번": 7,
            "사업별 순번": 5,
            "원본경로": r"C:\original\clip.MP4",
        }

        generated = self.editor.prepare_new_project_number(values, specimen=False)

        self.assertEqual(generated["미디어코드"], "OM")
        self.assertEqual(generated["데이터베이스번호"], "NSMK-OM-PL-260700005")
        self.assertEqual(generated["동영상 파일명"], "NSMK-OM-PL-260700005.mp4")

        generated.update({"미디어타입": "이미지", "사업별 순번": "9"})
        regenerated = self.editor.prepare_new_project_number(generated, specimen=False)
        self.assertEqual(regenerated["미디어코드"], "OI")
        self.assertEqual(regenerated["데이터베이스번호"], "NSMK-OI-PL-260700009")

    def test_database_check_repairs_well_formed_but_stale_media_identifier(self):
        table = self.editor.data_table
        table.insertRow(0)
        self.editor.write_data_row(0, {
            "NO.": 1,
            "기관코드": "NSMK",
            "분류코드": "PL",
            "미디어타입": "동영상",
            "미디어코드": "OI",
            "사업년도": 2026,
            "순번": 7,
            "사업별 순번": 5,
            "원본경로": r"C:\original\clip.mp4",
            "데이터베이스번호": "NSMK-OI-PL-260700005",
        })
        table.clearSelection()

        with patch.object(self.editor, "show_database_number_repair_log") as result:
            self.editor.repair_invalid_database_numbers()

        self.assertEqual(
            table.item(0, ROW_COLUMNS.index("데이터베이스번호")).text(),
            "NSMK-OM-PL-260700005",
        )
        self.assertEqual(table.item(0, ROW_COLUMNS.index("미디어코드")).text(), "OM")
        self.assertEqual(
            table.item(0, ROW_COLUMNS.index("동영상 파일명")).text(),
            "NSMK-OM-PL-260700005.mp4",
        )
        self.assertEqual(len(result.call_args.args[2]), 1)

    def test_observation_project_sequence_is_counted_per_media_type(self):
        table = self.editor.data_table
        for row, (media_type, sequence) in enumerate((("이미지", 1), ("동영상", 1), ("이미지", 2))):
            table.insertRow(row)
            self.editor.write_data_row(row, {
                "NO.": row + 1,
                "미디어타입": media_type,
                "사업년도": 2026,
                "순번": 7,
                "사업별 순번": sequence,
            })
        table.clearSelection()
        self.editor._last_selected_row = None
        self.editor.set_form_value("사업년도", 2026)
        self.editor.set_form_value("순번", 7)

        self.editor.set_form_value("미디어타입", "이미지")
        self.assertEqual(self.editor.project_sequence_for_form(False, new=True), 3)
        self.editor.set_form_value("미디어타입", "동영상")
        self.assertEqual(self.editor.project_sequence_for_form(False, new=True), 2)
        self.editor.set_form_value("미디어타입", "사운드")
        self.assertEqual(self.editor.project_sequence_for_form(False, new=True), 1)

    def test_renumber_project_sequences_does_not_cross_media_types(self):
        table = self.editor.data_table
        rows = (("이미지", 10), ("동영상", 50), ("이미지", 20), ("동영상", 60))
        for row, (media_type, sequence) in enumerate(rows):
            table.insertRow(row)
            self.editor.write_data_row(row, {
                "NO.": row + 1,
                "기관코드": "NSMK",
                "분류코드": "PL",
                "미디어타입": media_type,
                "사업년도": 2026,
                "순번": 7,
                "사업별 순번": sequence,
            })

        self.editor.renumber_project_sequences_below(table, ROW_COLUMNS, 0, False)

        values = [table.item(row, ROW_COLUMNS.index("사업별 순번")).text() for row in range(4)]
        self.assertEqual(values, ["10", "50", "11", "60"])

    def test_project_sequence_commit_survives_down_and_right_navigation_order(self):
        table = self.editor.data_table
        sequence_column = ROW_COLUMNS.index("사업별 순번")
        for row in range(2):
            table.insertRow(row)
            self.editor.write_data_row(row, {
                "NO.": row + 1,
                "기관코드": "NSMK",
                "분류코드": "PL",
                "미디어타입": "이미지",
                "사업년도": 2026,
                "순번": 7,
                "사업별 순번": row + 1,
            })

        # 아래 방향키는 선택 행 변경이 셀 편집 완료보다 먼저 전달될 수 있다.
        table.setCurrentCell(0, sequence_column)
        self.editor.set_form_value("사업별 순번", "1")
        self.editor._form_dirty = False
        table.setCurrentCell(1, sequence_column)
        table.item(0, sequence_column).setText("10")
        self.assertEqual(table.item(0, sequence_column).text(), "10")

        # 오른쪽 이동처럼 같은 행에 머무는 경우도 동일하게 유지되어야 한다.
        table.setCurrentCell(0, sequence_column)
        table.item(0, sequence_column).setText("20")
        table.setCurrentCell(0, sequence_column + 1)
        self.assertEqual(table.item(0, sequence_column).text(), "20")
        self.assertEqual(self.callback_errors, [])

    def test_standard_address_master_wins_over_reverse_geocoder_english(self):
        self.assertEqual(
            self.editor.preferred_address_english(
                "sido", "경상북도", "North Gyeongsang"
            ),
            "Gyeongsangbuk-do",
        )

    def test_reverse_geocoder_english_is_fallback_for_unknown_address(self):
        self.assertEqual(
            self.editor.preferred_address_english(
                "sido", "마스터에 없는 지역", "Provider English Name"
            ),
            "Provider English Name",
        )

    def test_reverse_address_promotes_metropolitan_city_to_sido(self):
        parsed = self.editor.parse_reverse_geocode_address({
            "city": "대구광역시",
            "borough": "동구",
            "suburb": "공산동",
        })

        self.assertEqual(parsed, {
            "sido": "대구광역시",
            "sigungu": "동구",
            "eupmyeondong": "공산동",
            "detail": "",
        })

    def test_reverse_address_keeps_gyeongbuk_city_hierarchy(self):
        parsed = self.editor.parse_reverse_geocode_address({
            "state": "경상북도",
            "city": "구미시",
            "town": "산동읍",
        })

        self.assertEqual(parsed, {
            "sido": "경상북도",
            "sigungu": "구미시",
            "eupmyeondong": "산동읍",
            "detail": "",
        })

    def test_reverse_address_reads_gunwi_from_municipality(self):
        parsed = self.editor.parse_reverse_geocode_address({
            "state": "대구광역시",
            "municipality": "군위군",
            "town": "효령면",
        })

        self.assertEqual(parsed, {
            "sido": "경상북도",
            "sigungu": "군위군",
            "eupmyeondong": "효령면",
            "detail": "",
        })

    def test_reverse_address_merges_gunwi_hierarchy_and_ri(self):
        parsed = self.editor.parse_reverse_geocode_address({
            "village": "노행리",
            "town": "효령면",
            "county": "군위군",
            "city": "대구광역시",
        })
        self.assertEqual(parsed, {
            "sido": "경상북도",
            "sigungu": "군위군",
            "eupmyeondong": "효령면",
            "detail": "노행리",
        })

    def test_reverse_address_prefers_suncheon_city_over_conflicting_gurye_county(self):
        parsed = self.editor.parse_reverse_geocode_address({
            "state": "전남광주통합특별시",
            "city": "순천시",
            "county": "구례군",
            "village": "내구리",
        })
        self.assertEqual(parsed, {
            "sido": "전라남도",
            "sigungu": "순천시",
            "eupmyeondong": "",
            "detail": "내구리",
        })

    def test_reverse_geocode_merges_broader_ri_when_detailed_rural_result_omits_it(self):
        responses = []
        for address in (
            {"road": "노행길", "town": "효령면", "county": "군위군", "city": "대구광역시"},
            {"village": "노행리", "city": "대구광역시"},
        ):
            response = MagicMock()
            response.__enter__.return_value.read.return_value = json.dumps(
                {"address": address}, ensure_ascii=False
            ).encode("utf-8")
            responses.append(response)

        with patch("ui.tabs.upload_tab.urllib.request.urlopen", side_effect=responses) as urlopen, \
             patch("ui.tabs.upload_tab.time.sleep"):
            parsed = self.editor.reverse_geocode(36.20215, 128.6234513, "ko")

        self.assertEqual(urlopen.call_count, 2)
        self.assertEqual(parsed, {
            "sido": "경상북도",
            "sigungu": "군위군",
            "eupmyeondong": "효령면",
            "detail": "노행리",
        })

    def test_integrated_region_restores_former_gwangju_for_five_districts(self):
        for district in ("동구", "서구", "남구", "북구", "광산구"):
            with self.subTest(district=district):
                parsed = self.editor.parse_reverse_geocode_address({
                    "state": "전남광주통합특별시",
                    "borough": district,
                    "suburb": "테스트동",
                })
                self.assertEqual(parsed["sido"], "광주광역시")
                self.assertEqual(parsed["sigungu"], district)

    def test_integrated_region_restores_former_jeonnam_for_other_cities_and_counties(self):
        for sigungu in ("목포시", "여수시", "담양군"):
            with self.subTest(sigungu=sigungu):
                parsed = self.editor.parse_reverse_geocode_address({
                    "state": "전남광주통합특별시",
                    "municipality": sigungu,
                    "town": "테스트면",
                })
                self.assertEqual(parsed["sido"], "전라남도")
                self.assertEqual(parsed["sigungu"], sigungu)

    def test_reverse_address_supports_alternate_integrated_region_name(self):
        parsed = self.editor.parse_reverse_geocode_address({
            "state": "광주전남통합특별시",
            "borough": "광산구",
        })
        self.assertEqual(parsed["sido"], "광주광역시")

    def test_coordinate_lookup_uses_one_korean_request_then_address_master(self):
        self.editor.set_form_value("위도", "36.1195")
        self.editor.set_form_value("경도", "128.3446")
        result = {"sido": "경상북도", "sigungu": "구미시", "eupmyeondong": "산동읍"}

        with patch.object(self.editor, "reverse_geocode", return_value=result) as lookup:
            self.editor.fill_address_from_latlon_nominatim()

        lookup.assert_called_once_with(36.1195, 128.3446, "ko")
        self.assertEqual(self.editor.get_form_value("관찰위치 시도 국문"), "경상북도")
        self.assertEqual(self.editor.get_form_value("관찰위치 시도 영문"), "Gyeongsangbuk-do")
        self.assertEqual(self.editor.get_form_value("관찰위치 시군구 국문"), "구미시")

    def test_preview_navigation_selects_matching_table_row(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            paths = []
            for index in range(2):
                path = root / f"image-{index}.png"
                Image.new("RGB", (8, 8), "red").save(path)
                paths.append(path)
                row = self.editor.data_table.rowCount()
                self.editor.data_table.insertRow(row)
                self.editor.write_data_row(row, {"NO.": row + 1, "파일경로": str(path)})
            self.editor.data_table.selectRow(0)

            preview_path, _, current, total, message = self.editor.preview_navigation_result(False, 1)

            self.assertEqual(Path(preview_path), paths[1])
            self.assertEqual((current, total, message), (2, 2, ""))
            self.assertEqual(self.editor.current_data_row(), 1)

    def test_original_path_immediately_updates_inline_preview(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "preview.png"
            Image.new("RGB", (20, 10), "blue").save(path)

            self.editor.set_form_value("파일경로", "")
            self.editor.set_form_value("원본경로", str(path))
            self.app.processEvents()

            pixmap = self.editor.media_thumbnail_label.pixmap()
            self.assertIsNotNone(pixmap)
            self.assertFalse(pixmap.isNull())

    def test_media_folder_import_preserves_user_start_per_media_type(self):
        self.editor.set_form_value("사업별 순번", "10")
        def apply_duration(_path, media_type, _probe=None):
            if media_type == "동영상":
                self.editor.set_form_value("동영상 재생시간", "0:19")

        files = [
            (Path("image-1.jpg"), "이미지"),
            (Path("video-1.mp4"), "동영상"),
            (Path("image-2.jpg"), "이미지"),
        ]
        with patch("ui.tabs.upload_tab.QFileDialog.getExistingDirectory", return_value="media"), \
             patch.object(self.editor, "scan_media_folder", return_value=files), \
             patch.object(self.editor, "apply_location_from_media"), \
             patch.object(self.editor, "apply_duration_from_media", side_effect=apply_duration), \
             patch.object(self.editor, "apply_date_from_media"), \
             patch.object(self.editor, "probe_ffprobe_format", return_value=None), \
             patch("ui.tabs.upload_tab.QMessageBox.information"):
            self.editor.import_media_folder()

        sequences = [
            self.editor.data_table.item(row, ROW_COLUMNS.index("사업별 순번")).text()
            for row in range(3)
        ]
        self.assertEqual(sequences, ["10", "10", "11"])
        durations = [
            self.editor.data_table.item(row, ROW_COLUMNS.index("동영상 재생시간")).text()
            for row in range(3)
        ]
        self.assertEqual(durations, ["", "0:19", ""])


if __name__ == "__main__":
    unittest.main()
