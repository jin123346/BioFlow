from contextlib import ExitStack
import os
import sys
import unittest
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
from PySide6.QtWidgets import QApplication

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


if __name__ == "__main__":
    unittest.main()
