
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
from utils.text_utils import normalize_text

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
        
       
        sheet_layout = QHBoxLayout()
        self.sheet_combo = QComboBox()
      
        self.sheet_combo.setPlaceholderText("시트 선택")
        
        
        self.load_button = QPushButton("시트 로드")
        self.load_button.clicked.connect(self.load_selected_sheet)
        
        sheet_layout.addWidget(QLabel("시트 선택:"))
        sheet_layout.addWidget(self.sheet_combo)
        sheet_layout.addWidget(self.load_button)
        
        layout.addLayout(sheet_layout)
        
        self.match_button=QPushButton("종 고유ID 찾기")
        self.match_button.clicked.connect(self.match_species_id)
        layout.addWidget(self.match_button)
        


        self.table_view = QTableView()
   

        self.table_view.setModel(self.table_model)
     

        layout.addWidget(self.table_view)


    def match_species_id(self):
        if self.input_df is None or self.input_df.empty:
            QMessageBox.warning(self,"경고","먼저 시트를 로드해주세요")
            return
        
        main_window=self.window()
        species_master_df = main_window.master_tab.df
        
        if species_master_df is None or species_master_df.empty:
            QMessageBox.warning(self,"경고","종정보 master가 로드되지 않았습니다.")
            return 
        
        df = self.input_df.copy()
        master  = species_master_df.copy()
       
        master["국명"] = master["국명"].apply(normalize_text)
        master["국명"] = master["국명"].apply(normalize_text)
        
        df["국명"] = df["국명 (잘라내서 붙이기 안됨 복사해서 붙이기는 허용)"].apply(normalize_text)
        df["학명"] = df["학명 (국명입력시 자동생성)"].apply(normalize_text)
        df["matched_species_id"]=""
        df["matched_by"]=""
        df["matched_name"]=""
        df["matched_sci_name"]=""
        df["matched_status"]=""
        for idx, row in df.iterrows():
            existing_id = normalize_text(row.get("종고유ID",""))
            kor_name = normalize_text(row.get("국명 (잘라내서 붙이기 안됨 복사해서 붙이기는 허용)",""))
            sci_name= normalize_text(row.get("학명 (국며입력시 자동생성)",""))
            
            #1.id 검증
            if existing_id:
                matched = master[master["종학명정보ID"]==existing_id]
                if not matched.empty:
                    df.at[idx,"matched_species_id"]= matched.iloc[0].get("종학명정보ID", "")
                    df.at[idx,"matched_by"]="기존Id"
                    df.at[idx,"matched_name"]= matched.iloc[0].get("국명", "")
                    df.at[idx,"matched_sci_name"]= matched.iloc[0].get("학명", "")
                    df.at[idx,"matched_status"]="매칭 성공"
                else:
                    df.at[idx,"matched_status"]="매칭 실패 - ID 없음"
            
            #2. 국명 매칭
            if  kor_name:
                matched = master[master["국명"]==kor_name]
                if not matched.empty:
                    df.at[idx,"matched_species_id"]= matched.iloc[0].get("종학명정보ID", "")
                    df.at[idx,"matched_by"]="국명"
                    df.at[idx,"matched_name"]= matched.iloc[0].get("국명", "")
                    df.at[idx,"matched_sci_name"]= matched.iloc[0].get("학명", "")
                    df.at[idx,"matched_status"]="매칭 성공"
                    continue
            #3. 학명 매칭
            if sci_name:
                matched = master[master["학명"]==sci_name]
                if not matched.empty:
                    df.at[idx,"matched_species_id"]= matched.iloc[0].get("종학명정보ID", "")
                    df.at[idx,"matched_by"]="학명"
                    df.at[idx,"matched_name"]= matched.iloc[0].get("국명", "")
                    df.at[idx,"matched_sci_name"]= matched.iloc[0].get("학명", "")
                    df.at[idx,"matched_status"]="매칭 성공"
                    continue
                
            # 4. 실패
            df.at[idx, "match_status"] = "미매칭"
        
        self.input_df = df
        self.table_model.set_dataframe(df)
        self.table_view.resizeColumnsToContents()
        
        QMessageBox.information(self,"매칭 완료","종 고유ID 매칭이 완료되었습니다. 매칭 결과를 확인해주세요.")
        
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