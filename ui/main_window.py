import json
import shutil
from datetime import datetime
from pathlib import Path
import pandas as pd
from PySide6.QtWidgets import ( QMainWindow, QTableWidget, QTableWidgetItem,QTabWidget, QMessageBox)
from ui.tabs.species_master_tab import SpeciesMasterTab
from ui.tabs.species_review_tab import SpeciesReviewTab
from ui.tabs.upload_tab import UploadTab
from ui.tabs.validataion_tab import ValidationTab
from ui.tabs.sql_tab import SqlTab
from ui.tabs.gbif_tab import GbifTab
from config.paths import MASTER_FILE_PATH, LOCAL_VERSION_FILE_PATH, MASTER_DIR, META_DIR ,HISTORY_DIR




# 앱의 가장 바깥 창(메인창)
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.df= None  #엑셀 데이터프레임 저장할 변수

        
        self.setWindowTitle("BioFlow") #창 제목 설정
        self.resize(1200,800) #창 크기 설정

        self.init_ui()

    def init_ui(self):
        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)

        self.master_tab = SpeciesMasterTab()
        self.upload_tab = UploadTab()
        self.species_review_tab = SpeciesReviewTab(self.upload_tab)
        self.validation_tab = ValidationTab()
        self.sql_tab = SqlTab()
        self.gbif_tab = GbifTab()


        self.tabs.addTab(self.master_tab, "종정보 마스터")
        self.tabs.addTab(self.upload_tab, "신규 데이터")
        self.tabs.addTab(self.species_review_tab, "신규/업데이트 종 정리")
        self.tabs.addTab(self.validation_tab, "비교 결과")
        self.tabs.addTab(self.sql_tab, "SQL 생성")
        self.tabs.addTab(self.gbif_tab, "GBIF 데이터")

    def closeEvent(self, event):
        has_rows = self.upload_tab.data_table.rowCount() > 0
        has_form_content = self.upload_tab.has_form_content(self.upload_tab.collect_form_values())
        has_specimen_rows = self.upload_tab.specimen_data_table.rowCount() > 0
        has_specimen_form_content = self.upload_tab.has_specimen_form_content(
            self.upload_tab.collect_specimen_form_values()
        )
        if not has_rows and not has_form_content and not has_specimen_rows and not has_specimen_form_content:
            event.accept()
            return

        reply = QMessageBox.question(
            self,
            "종료 확인",
            "작성 중인 신규 데이터가 있습니다. 임시저장할까요?",
            QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel,
            QMessageBox.Yes,
        )
        if reply == QMessageBox.Cancel:
            event.ignore()
            return
        if reply == QMessageBox.Yes:
            self.upload_tab.autosave_draft()
            self.upload_tab.autosave_specimen_draft()
        event.accept()
