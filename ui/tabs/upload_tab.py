
from __future__ import annotations

import re
import shutil
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

import pandas as pd
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from PySide6.QtCore import QAbstractTableModel, QModelIndex, QObject, Qt, Signal
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSplitter,
    QTableView,
    QTabWidget,
    QVBoxLayout,
    QWidget,
    QComboBox,
)
from services.upload_input_service import UploadInputService
from services.upload_to_excel import UploadToExcelService
from models.pandas_table_model import PandasTableModel

# =========================
# 설정 / 상수
# =========================
OBSERVATION_MEDIA_CODE= "OB"
SPECIMEN_MEDIA_CODE = "SP"

BASIC_SHEET_NAME="정보입력"
IMAGE_SHEET_NAME="이미지"
MOVIE_SHEET_NAME="동영상"
SOUND_SHEET_NAME="사운드"

SPECIES_LOOKUP_SHEET_NAME = "Sheet3"
OUTPUT_TEMPLATE_SHEET_NAME="Sheet1"

INPUT_HEADER_ROW = 1
OUTPUT_GUIDE_ROW = 1
OUTPUT_HEADER_ROW = 2
OUTPUT_DATA_START_ROW = 3

TAXON_CODE_RULES = [
    ("식물계", "PL"),
    ("Plantae", "PL"),
    ("곤충강", "IN"),
    ("Insecta", "IN"),
    ("조강", "BI"),
    ("Aves", "BI"),
    ("조기강", "PI"),
    ("Actinopterygii", "PI"),
    ("포유강", "MM"),
    ("Mammalia", "MM"),
] #버섯, 절지동무류, 양서류, 파충류, 어류, 해양생물, 등 기타 생물군은 추후 추가 예정





class UploadTab(QWidget):
    def __init__(self):
        super().__init__()
        print("upload 1")

        self.input_df = None
        print("upload 2")
        self.upload_input_service = UploadInputService()
        print("upload 3")
        self.table_model = PandasTableModel()
        print("upload 4")
        self.init_ui()
        print("upload 5")
        
    def init_ui(self):
        print("ui 1")
        layout = QVBoxLayout()
        self.setLayout(layout)
        
        print("ui 2")
        title = QLabel("NARIS 업로드 파일 생성")
        layout.addWidget(title)
        
        print("ui 3")
        file_layout = QHBoxLayout()
        
        self.file_path_edit=QLineEdit()
        self.file_path_edit.setPlaceholderText("업로드할 엑셀 파일을 선택하세요")
        
        print("ui 4")
        self.select_file_button =  QPushButton("파일 선택")
        self.select_file_button.clicked.connect(self.select_input_file)
        
        file_layout.addWidget(self.file_path_edit)
        file_layout.addWidget(self.select_file_button)  
        layout.addLayout(file_layout)
        
        print("ui 5")
        sheet_layout = QHBoxLayout()
        self.sheet_combo = QComboBox()
        print("ui 6")
        self.sheet_combo.setPlaceholderText("시트 선택")
        
        print("ui 7")
        self.load_button = QPushButton("시트 로드")
        self.load_button.clicked.connect(self.load_selected_sheet)
        
        sheet_layout.addWidget(QLabel("시트 선택:"))
        sheet_layout.addWidget(self.sheet_combo)
        sheet_layout.addWidget(self.load_button)
        
        layout.addLayout(sheet_layout)
        
        print("ui 8-1")

        self.table_view = QTableView()
        print("ui 8-2")

        self.table_view.setModel(self.table_model)
        print("ui 8-3")

        layout.addWidget(self.table_view)
        
        print("ui 9")
        
        
    def select_input_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "업로드할 엑셀 파일 선택", "", "Excel Files (*.xlsx *.xls);;All Files (*)")
        
        if not file_path:
            return
        
        self.file_path_edit.setText(file_path)
        
        try:
            sheet_names = self.upload_input_service.get_sheet_names(file_path)
            self.sheet_combo.clear()
            self.sheet_combo.addItems(sheet_names)
            
            #자주 쓰는 시트가 있으면 기본 선택
            priority_sheets = ["정보입력", "이미지", "동영상", "사운드"]
            for sheet in priority_sheets:
                idx = self.sheet_combo.findText(sheet)
                if idx >= 0:
                    self.sheet_combo.setCurrentIndex(idx)
                    break
        except Exception as e:
            QMessageBox.critical(self, "오류", f"파일을 읽는 중 오류가 발생했습니다: {str(e)}")
            
    def load_selected_sheet(self):
        file_path = self.file_path_edit.text().strip()
        sheet_name = self.sheet_combo.currentText()
        if not file_path or not sheet_name:
            QMessageBox.warning(self, "경고", "파일과 시트를 모두 선택해주세요.")
            return 
        
        self.input_df = self.upload_input_service.load_sheet(file_path, sheet_name)
        self.table_model.set_dataframe(self.input_df)
        self.table_view.resizeColumnsToContents()