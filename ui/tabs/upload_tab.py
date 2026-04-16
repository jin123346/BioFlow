
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
)

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

        layout = QVBoxLayout()
        self.setLayout(layout)

        layout.addWidget(QLabel("신규 데이터 업로드 탭"))