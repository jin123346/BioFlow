
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
from PySide6.QtCore import QAbstractTableModel, QModelIndex, QObject, Qt, Signal
from config.paths import MASTER_FILE_PATH, LOCAL_VERSION_FILE_PATH, MASTER_DIR, META_DIR
import pandas as pd
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
import re
import shutil
from utils.excel_utils import safe_int
from utils.text_utils import normalize_scientific_name, build_db_no
from utils.taxon_utils import detect_taxon_code
from models.match_result import MatchResult
from pathlib import Path
from typing import Optional

class UploadToExcelService(QObject):
    progress= Signal(str)
    
    def __init__(self):
        super().__init__()
        
        
    def load_input_sheet(self,path: str, sheet_name:str = BASIC_SHEET_NAME) -> pd.DataFrame:
        self.progress.emit(f"입력 시트 로드 : {Path(path).name} - {sheet_name}")
        df= pd.read_excel(path, sheet_name=sheet_name, header=INPUT_HEADER_ROW-1 dtype=object)
        df = df.dropna(how="all") # 전체 행이 비어있는 경우 제거
        return df
    
    def load_spe