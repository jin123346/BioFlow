import json
import os
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication

from services.classification_service import ClassificationService
from ui.tabs.species_master_tab import SpeciesMasterTab


def sample_records():
    return [
        {
            "clssc_sstem_info_id": "10",
            "taxon_phlm_sn": "6",
            "taxon_kingdom_sn": "1",
            "taxonomy_name": "동물계 > 절지동물문 > 곤충강 > 나비목 > 호랑나비과",
            "taxonomy_latin": "Animalia > Arthropoda > Insecta > Lepidoptera > Papilionidae",
        },
        {
            "clssc_sstem_info_id": "20",
            "taxon_phlm_sn": "2",
            "taxon_kingdom_sn": "2",
            "taxonomy_name": "식물계 > 관속식물문 > 목련강 > 장미목 > 장미과",
            "taxonomy_latin": "Plantae > Tracheophyta > Magnoliopsida > Rosales > Rosaceae",
        },
    ]


def test_classification_records_can_be_browsed_and_searched(tmp_path):
    path = tmp_path / "master.json"
    path.write_text(json.dumps(sample_records(), ensure_ascii=False), encoding="utf-8")

    with patch("services.classification_service.CLASSIFICATION_FILE", path):
        records = ClassificationService.records()
        assert len(records) == 2
        assert ClassificationService.search_records(records, "호랑나비") == [records[0]]
        assert ClassificationService.search_records(records, "rosaceae") == [records[1]]
        assert ClassificationService.search_records(records, "20") == [records[1]]


def test_species_master_page_has_lazy_classification_tab():
    app = QApplication.instance() or QApplication([])
    records = sample_records()
    with patch("ui.tabs.species_master_tab.QTimer.singleShot"), \
         patch.object(ClassificationService, "records", return_value=records):
        widget = SpeciesMasterTab()
        try:
            assert "분류체계(clssfy)" in [
                widget.sub_tabs.tabText(index) for index in range(widget.sub_tabs.count())
            ]
            assert widget.classification_records is None
            widget.sub_tabs.setCurrentIndex(widget.classification_tab_index)
            assert widget.classification_table.rowCount() == 2
            widget.classification_search_input.setText("Papilionidae")
            widget.search_classification()
            assert widget.classification_table.rowCount() == 1
            assert widget.classification_table.item(0, 0).text() == "10"
        finally:
            widget.close()
            app.processEvents()
