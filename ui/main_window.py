import json
import shutil
from datetime import datetime
from pathlib import Path
import pandas as pd
from PySide6.QtWidgets import ( QMainWindow, QTableWidget, QTableWidgetItem,QTabWidget)
from ui.tabs.species_master_tab import SpeciesMasterTab
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

        print("A")
        self.master_tab = SpeciesMasterTab()
        print("B")
        self.upload_tab = UploadTab()
        print("C")
        self.validation_tab = ValidationTab()
        print("D")
        self.sql_tab = SqlTab()
        print("E")
        self.gbif_tab = GbifTab()
        print("F")

        self.tabs.addTab(self.master_tab, "종정보 마스터")
        self.tabs.addTab(self.upload_tab, "신규 데이터")
        self.tabs.addTab(self.validation_tab, "비교 결과")
        self.tabs.addTab(self.sql_tab, "SQL 생성")
        self.tabs.addTab(self.gbif_tab, "GBIF 데이터")

             
            