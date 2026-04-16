from pathlib import Path

import pandas as pd
from PySide6.QtCore import QSettings, QTimer
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from config.paths import MASTER_FILE_PATH
from services.species_master_service import SpeciesMasterService
from services.species_compare_service import SpeciesCompareService
from utils.excel_utils import find_target_sheet_name
from utils.text_utils import normalize_text
from config.constants import TARGET_SHEET_NAME


class SpeciesMasterTab(QWidget):
    def __init__(self):
        super().__init__()
        self.df = None
        self.update_files = []
        self.df_candidates = None
        self.df_compare_result = None
        self.status_label = None
        self.settings = QSettings("Bioflow", "BioflowApp")

        self.master_service = SpeciesMasterService()
        self.compare_service = SpeciesCompareService()
        
        self.init_ui()
        QTimer.singleShot(100, self.load_last_species_master)

    def init_ui(self):
        main_layout = QVBoxLayout()
        self.setLayout(main_layout)

        title_label = QLabel("종정보 마스터 관리")
        main_layout.addWidget(title_label)

        file_layout = QHBoxLayout()

        self.master_file_path = QLineEdit()
        self.master_file_path.setPlaceholderText("기존 종정보 엑셀 파일 경로")

        self.select_file_button = QPushButton("파일선택")
        self.select_file_button.clicked.connect(self.select_master_file)

        file_layout.addWidget(QLabel("종정보 마스터"))
        file_layout.addWidget(self.master_file_path)
        file_layout.addWidget(self.select_file_button)
        main_layout.addLayout(file_layout)

        self.status_label = QLabel("대기 중")
        self.status_label.setStyleSheet("color: blue;")
        main_layout.addWidget(self.status_label)

        self.load_button = QPushButton("엑셀 로드")
        self.load_button.clicked.connect(self.load_excel)
        main_layout.addWidget(self.load_button)

        self.preview_table = QTableWidget()
        main_layout.addWidget(self.preview_table)

        self.save_button = QPushButton("로컬 저장")
        self.save_button.clicked.connect(self.save_master_to_local)
        main_layout.addWidget(self.save_button)

        update_file_layout = QHBoxLayout()

        self.select_update_files_button = QPushButton("업데이트 파일 여러개 선택")
        self.select_update_files_button.clicked.connect(self.select_update_files)

        self.extract_candidates_button = QPushButton("업데이트 후보 추출")
        self.extract_candidates_button.clicked.connect(self.extract_update_candidates)

        self.compare_button = QPushButton("신규/업로드 분류")
        self.compare_button.clicked.connect(self.compare_with_master)

        self.save_compare_result_button = QPushButton("업데이트 엑셀 저장")
        self.save_compare_result_button.clicked.connect(self.save_compare_result_to_excel)

        update_file_layout.addWidget(self.select_update_files_button)
        update_file_layout.addWidget(self.extract_candidates_button)
        update_file_layout.addWidget(self.compare_button)
        update_file_layout.addWidget(self.save_compare_result_button)
        main_layout.addLayout(update_file_layout)

        self.update_files_list = QListWidget()
        main_layout.addWidget(self.update_files_list)

        self.result_table = QTableWidget()
        main_layout.addWidget(self.result_table)

        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setPlaceholderText("작업 로그가 여기에 표시됩니다.")
        main_layout.addWidget(self.log_text)

    def append_log(self, message: str):
        self.log_text.append(message)

    def select_master_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "종정보 엑셀 파일 선택",
            "",
            "Excel Files (*.xlsx *.xls);;All Files (*)"
        )

        if file_path:
            self.master_file_path.setText(file_path)
            self.append_log(f"파일 선택 완료 : {file_path}")
        else:
            self.append_log("파일 선택이 취소되었습니다.")

    def load_excel(self):
        file_path = self.master_file_path.text().strip()

        if not file_path:
            self.append_log("파일을 먼저 선택해 주세요")
            return

        try:
            df = self.master_service.load_excel_file(file_path)
            self.df = df

            self.append_log(f"엑셀 로드 성공 (행 : {len(df)} , 컬럼: {len(df.columns)})")
            self.append_log(f"컬럼 목록 : {list(df.columns)}")
            self.append_log(f"엑셀 로드 완료 (행: {len(df)})")

            self.show_dataframe_preview(df)

        except Exception as e:
            self.append_log(f"엑셀 로드 실패 : {str(e)}")

    def show_dataframe_preview(self, df: pd.DataFrame):
        preview_df = df.head(20)

        self.preview_table.clear()
        self.preview_table.setRowCount(len(preview_df))
        self.preview_table.setColumnCount(len(preview_df.columns))
        self.preview_table.setHorizontalHeaderLabels([str(col) for col in preview_df.columns])

        for row_idx in range(len(preview_df)):
            for col_idx in range(len(preview_df.columns)):
                value = preview_df.iloc[row_idx, col_idx]
                display_value = "" if pd.isna(value) else str(value)
                item = QTableWidgetItem(display_value)
                self.preview_table.setItem(row_idx, col_idx, item)

        self.preview_table.resizeColumnsToContents()
        self.append_log("미리보기 테이블 표시 완료(상위 20행)")

    def save_master_to_local(self):
        if self.df is None:
            self.append_log("저장할 데이터가 없습니다. 먼저 엑셀을 로드해 주세요.")
            return

        source_file_path = self.master_file_path.text().strip()

        if not source_file_path:
            self.append_log("저장할 파일 경로가 없습니다. 먼저 엑셀을 로드해 주세요.")
            return

        try:
            result = self.master_service.save_master_to_local(self.df, source_file_path)
            self.append_log(f"로컬 저장 성공 : {result['master_path']}")
            self.append_log(f"백업 저장 성공 : {result['backup_path']}")
            self.append_log(f"버전 정보 저장 성공 : {result['version_path']}")
        except Exception as e:
            self.append_log(f"로컬 저장 실패 : {str(e)}")

    def load_last_species_master(self):
        self.append_log("종정보 마스터 확인 중...")
        self.status_label.setText("종정보 마스터 로딩 중...")
        self.status_label.setStyleSheet("color: green;")
        QApplication.processEvents()

        try:
            df = self.master_service.load_last_master()
            self.df = df

            self.master_file_path.setText(str(MASTER_FILE_PATH))
            self.append_log("종정보 마스터 로드 성공")
            self.show_dataframe_preview(df)
            self.status_label.setText("로드 완료")
            self.status_label.setStyleSheet("color: red;")

        except Exception as e:
            self.status_label.setText("로드 실패")
            self.append_log(f"자동 로드 실패: {str(e)}")

    def reload_master(self):
        try:
            df = self.master_service.reload_master()
            self.df = df
            self.show_dataframe_preview(df)
            self.append_log("재로딩 완료")
        except Exception as e:
            self.append_log(f"재로딩 실패: {str(e)}")

    # 아래 4개는 다음 단계에서 species_compare_service로 분리
    def select_update_files(self):
        file_paths, _ = QFileDialog.getOpenFileNames(
            self,
            "업데이트용 엑셀 파일 선택",
            "",
            "Excel Files (*.xlsx *.xls);;All Files (*)"
        )

        if file_paths:
            self.update_files = file_paths
            self.update_files_list.clear()

            for path in file_paths:
                self.update_files_list.addItem(path)
            self.append_log(f"업데이트 파일 선택 완료 : {len(file_paths)}개 파일")
        else:
            self.append_log("업데이트 파일 선택이 취소되었습니다.")

    def extract_update_candidates(self):
        if not self.update_files:
            self.append_log("업데이트할 파일을 먼저 선택해 주세요.")
            return

        all_rows = []

        for file_path in self.update_files:
            try:
                actual_sheet_name = find_target_sheet_name(file_path, TARGET_SHEET_NAME)

                if not actual_sheet_name:
                    self.append_log(f"대상 시트를 찾지 못했습니다: {Path(file_path).name}")
                    continue

                df_sheet = pd.read_excel(file_path, sheet_name=actual_sheet_name)
                df_sheet.columns = [str(col).strip().replace("\n", " ") for col in df_sheet.columns]

                df_sheet["source_file"] = Path(file_path).name
                df_sheet["source_path"] = file_path
                df_sheet["update_type"] = range(2, len(df_sheet) + 2)

                all_rows.append(df_sheet)
                self.append_log(f"업데이트 후보 추출 성공 : {file_path} (행: {len(df_sheet)})")
            except Exception as e:
                self.append_log(f"업데이트 후보 추출 실패 : {file_path} - {str(e)}")

        if not all_rows:
            self.append_log("업데이트 후보로 추출된 데이터가 없습니다.")
            return

        self.df_candidates = pd.concat(all_rows, ignore_index=True)
        self.append_log(f"업데이트 후보 통합 완료 : 총 {len(self.df_candidates)} 행")
        self.show_dataframe_preview(self.df_candidates)

    def show_result_preview(self, df: pd.DataFrame):
        self.result_table.clear()
        self.result_table.setRowCount(len(df))
        self.result_table.setColumnCount(len(df.columns))
        self.result_table.setHorizontalHeaderLabels([str(col) for col in df.columns])

        for row_idx in range(len(df)):
            for col_idx in range(len(df.columns)):
                value = df.iloc[row_idx, col_idx]
                display_value = "" if pd.isna(value) else str(value)

                item = QTableWidgetItem(display_value)
                self.result_table.setItem(row_idx, col_idx, item)

        self.result_table.resizeColumnsToContents()
        self.append_log("비교 결과 테이블 표시 완료")

    def save_compare_result_to_excel(self):
        if self.df_compare_result is None or self.df_compare_result.empty:
            self.append_log("저장할 비교 결과가 없습니다. 먼저 신규/업데이트 분류를 실행해 주세요.")
            return

        from datetime import datetime

        default_name = f"species_compare_result_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "비교 결과 저장",
            default_name,
            "Excel Files (*.xlsx)"
        )

        if not file_path:
            self.append_log("비교 결과 저장이 취소되었습니다.")
            return

        try:
            save_df = self.df_compare_result.copy()

            if "status" in save_df.columns:
                status_order = {
                    "신규": 0,
                    "업데이트": 1,
                    "업데이트(중복)": 2,
                    "검토필요": 3,
                    "식별정보부족": 4,
                }
                save_df["_status_order"] = save_df["status"].map(status_order).fillna(99)
                save_df = save_df.sort_values(by=["_status_order"]).drop(columns=["_status_order"])

            save_df.to_excel(file_path, index=False)
            self.append_log(f"비교 결과 저장 완료: {file_path}")

        except Exception as e:
            self.append_log(f"비교 결과 저장 실패: {str(e)}")

    def compare_with_master(self):
        if self.df is None:
            self.append_log("마스터 없음")
            return

        if self.df_candidates is None:
            self.append_log("후보 없음")
            return

        result_df = self.compare_service.compare(self.df, self.df_candidates)

        self.df_compare_result = result_df

        self.append_log("비교 완료")
        self.show_result_preview(result_df.head(50))





# def normalize_sheet_name(name: str) -> str:
#     if name is None:
#         return ""

#     text = str(name)
#     text = text.replace("\n", " ")
#     text = text.replace("\xa0", " ")   # 특수 공백 제거
#     text = " ".join(text.split())      # 연속 공백 정리 + 앞뒤 공백 제거
#     return text.strip()


# TARGET_SHEET_NAME = "국명이없거나 학명정보가잘못된것 기재 "
# TARGET_SHEET_NAME_NORMALIZED = normalize_sheet_name(TARGET_SHEET_NAME)

# # 앱의 가장 바깥 창(메인창)
# class SpeciesMasterTab(QWidget):
#     def __init__(self):
#         super().__init__()
#         self.df = None
#         self.update_files = [] # 업데이트용 엑셀 파일 경로 목록
#         self.df_candidates= None #업데이트 시트 통합 결과
#         self.df_compare_result=None #비교 결과 데이터프레임
#         self.status_label= None
#         self.settings = QSettings("Bioflow", "BioflowApp")

#         self.init_ui()
#         QTimer.singleShot(100, self.load_last_species_master)
        
#         #self.load_last_species_master()

#     def init_ui(self):
#         main_layout = QVBoxLayout()
#         self.setLayout(main_layout)

#         title_label = QLabel("종정보 마스터 관리")
#         main_layout.addWidget(title_label)


#         #종정보 마스터 파일 선택 영역
#         file_layout = QHBoxLayout()

#         self.master_file_path = QLineEdit()
#         self.master_file_path.setPlaceholderText("기존 종정보 엑셀 파일 경로")


#         self.select_file_button = QPushButton("파일선택")
#         self.select_file_button.clicked.connect(self.select_master_file)
        

#         file_layout.addWidget(QLabel("종정보 마스터"))
#         file_layout.addWidget(self.master_file_path)
#         file_layout.addWidget(self.select_file_button)

#         main_layout.addLayout(file_layout)

#         self.status_label =QLabel("대기 중")
#         self.status_label.setStyleSheet("color: blue;")
#         main_layout.addWidget(self.status_label)
#         self.setLayout(main_layout)

#         self.load_button = QPushButton("엑셀 로드")
#         self.load_button.clicked.connect(self.load_excel)
#         main_layout.addWidget(self.load_button)

#         self.preview_table= QTableWidget()
#         main_layout.addWidget(self.preview_table)
        
#         self.save_button=QPushButton("로컬 저장")
#         self.save_button.clicked.connect(self.save_master_to_local)
#         main_layout.addWidget(self.save_button)
        
#         #업데이트 파일 선택 영역
#         update_file_layout = QHBoxLayout()
        
#         self.select_update_files_button=QPushButton("업데이트 파일 여러개 선택")
#         self.select_update_files_button.clicked.connect(self.select_update_files)
        
#         self.extract_candidates_button= QPushButton("업데이트 후보 추출")
#         self.extract_candidates_button.clicked.connect(self.extract_update_candidates)
        
#         self.compare_button= QPushButton("신규/업로드 분류")
#         self.compare_button.clicked.connect(self.compare_with_master)

#         self.save_compare_result_button = QPushButton("업데이트 엑셀 저장")
#         self.save_compare_result_button.clicked.connect(self.save_compare_result_to_excel)
        


#         update_file_layout.addWidget(self.select_update_files_button)
#         update_file_layout.addWidget(self.extract_candidates_button)
#         update_file_layout.addWidget(self.compare_button)
#         update_file_layout.addWidget(self.save_compare_result_button)
#         main_layout.addLayout(update_file_layout)
        
#         self.update_files_list = QListWidget()
#         main_layout.addWidget(self.update_files_list)
        
#         #비교결과 테이블
#         self.result_table= QTableWidget()
#         main_layout.addWidget(self.result_table)
        

#         self.log_text = QTextEdit()
#         self.log_text.setReadOnly(True)
#         self.log_text.setPlaceholderText("작업 로그가 여기에 표시됩니다.")
#         main_layout.addWidget(self.log_text)

#     def select_master_file(self):
#         file_path, _ = QFileDialog.getOpenFileName(
#             self,
#             "종정보 엑셀 파일 선택",
#             "",
#             "Excel Files (*.xlsx *.xls);;All Files (*)"
#         )

#         if file_path:
#             self.master_file_path.setText(file_path)
#             self.append_log(f"파일 선택 완료 : {file_path}")
#         else:
#             self.append_log("파일 선택이 취소되었습니다.")

#     def append_log(self,message: str):
#         self.log_text.append(message)

#     def load_excel(self):
#         file_path = self.master_file_path.text()

#         if not file_path:
#             self.append_log("파일을 먼저 선택해 주세요")
#             return
        
#         try:
#             df = pd.read_excel(file_path, header=1)

#             self.append_log(f"엑셀 로드 성공 (행 : {len(df)} , 컬럼: {len(df.columns)})")

#             #컬럼 출력
#             self.append_log(f"컬럼 목록 : {list(df.columns)}")

#             clean_columns = []
#             for col in df.columns:
#                 if isinstance(col, (int, float)):
#                     continue
#                 clean_columns.append(str(col).strip().replace("\n", " "))

#             df = df.iloc[:, :len(clean_columns)]
#             df.columns = clean_columns

#             self.df = df
#             self.append_log(f"엑셀 로드 완료 (행: {len(df)})")

#             self.show_dataframe_preview(df)

#         except Exception as e:
#             self.append_log(f"엑셀 로드 실패 : {str(e)}")

    
#     #엑셀 테이블 뷰
#     def show_dataframe_preview(self, df: pd.DataFrame):
#         preview_df = df.head(20)

#         self.preview_table.clear()
#         self.preview_table.setRowCount(len(preview_df))
#         self.preview_table.setColumnCount(len(preview_df.columns))
#         self.preview_table.setHorizontalHeaderLabels([str(col) for col in preview_df.columns])

#         for row_idx in range(len(preview_df)):
#             for col_idx in range(len(preview_df.columns)):
#                 value= preview_df.iloc[row_idx, col_idx]

#                 if pd.isna(value):
#                     display_value=""
#                 else:
#                     display_value=str(value)
#                 item = QTableWidgetItem(display_value)
#                 self.preview_table.setItem(row_idx,col_idx,item)

#         self.preview_table.resizeColumnsToContents()
#         self.append_log("미리보기 테이블 표시 완료(상위 20행)")
        
#     def save_master_to_local(self):
#         if not hasattr(self, 'df'):
#             self.append_log("저장할 데이터가 없습니다. 먼저 엑셀을 로드해 주세요.")
#             return  
        
#         source_file_path = self.master_file_path.text().strip()
        
#         if not source_file_path:
#             self.append_log("저장할 파일 경로가 없습니다. 먼저 엑셀을 로드해 주세요.")
#             return  
        
#         try:
#             #폴더 생성
#             MASTER_DIR.mkdir(parents=True, exist_ok=True)
#             META_DIR.mkdir(parents=True, exist_ok=True)
#             HISTORY_DIR.mkdir(parents=True, exist_ok=True)
#             CACHE_DIR.mkdir(parents=True, exist_ok=True)
            
#             #1.원본파일 복사 저장
#             shutil.copy2(source_file_path, MASTER_FILE_PATH)
            
#             #1-1. 백업본저장
#             timestamp= datetime.now().strftime("%Y%m%d_%H%M%S")
#             backup_file_path = HISTORY_DIR/f"species_master_backup_{timestamp}.xlsx"
#             shutil.copy2(source_file_path, backup_file_path)
            

#             self.df.to_pickle(MASTER_CACHE_FILE)
            
            
#             #2.버전 정보 저장
#             version_data = {
#                 "version": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
#                 "source" : "local_upload",
#                 "update_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
#                 "file_name": Path(source_file_path).name,
#                 "saved_path": str(MASTER_FILE_PATH),
#                 "backup_path": str(backup_file_path)
#             }
            
#             with open(LOCAL_VERSION_FILE_PATH, 'w', encoding='utf-8') as f:
#                 json.dump(version_data, f,ensure_ascii=False, indent=4)
                
#             # 3. 히스토리 누적 저장
#             history_data = []

#             if LOCAL_VERSION_HISTORY_FILE_PATH.exists():
#                 try:
#                     with open(LOCAL_VERSION_HISTORY_FILE_PATH, "r", encoding="utf-8") as f:
#                         history_data = json.load(f)

#                     if not isinstance(history_data, list):
#                         history_data = []
#                 except Exception:
#                     history_data = []

#             history_data.append(version_data)

#             with open(LOCAL_VERSION_HISTORY_FILE_PATH, "w", encoding="utf-8") as f:
#                 json.dump(history_data, f, ensure_ascii=False, indent=4)

#             self.append_log(f"로컬 저장 성공 : {MASTER_FILE_PATH}")
#             self.append_log(f"백업 저장 성공 : {backup_file_path}")    
#             self.append_log(f"버전 정보 저장 성공 : {LOCAL_VERSION_FILE_PATH}")     
#         except Exception as e:
#             self.append_log(f"로컬 저장 실패 : {str(e)}")
        
    
#     def select_update_files(self):
#         file_paths, _ = QFileDialog.getOpenFileNames(
#             self,
#             "업데이트용 엑셀 파일 선택",
#             "",
#             "Excel Files (*.xlsx *.xls);;All Files (*)"
#         )
        

#         if file_paths:
#             self.update_files = file_paths
#             self.update_files_list.clear()
            
#             for path in file_paths:
#                 self.update_files_list.addItem(path)
#             self.append_log(f"업데이트 파일 선택 완료 : {len(file_paths)}개 파일")
#         else:
#             self.append_log("업데이트 파일 선택이 취소되었습니다.") 
            
#     def extract_update_candidates(self):
#         if not self.update_files:
#             self.append_log("업데이트할 파일을 먼저 선택해 주세요.")
#             return
        
#         all_rows = []
        
#         for file_path in self.update_files:
#             try:
#                 actual_sheet_name = find_target_sheet_name(file_path, TARGET_SHEET_NAME)

#                 if not actual_sheet_name:
#                     self.append_log(f"대상 시트를 찾지 못했습니다: {Path(file_path).name}")
#                     continue

#                 df_sheet = pd.read_excel(file_path, sheet_name=actual_sheet_name)
#                 df_sheet.columns = [str(col).strip().replace("\n", " ") for col in df_sheet.columns]

#                 df_sheet["source_file"] = Path(file_path).name
#                 df_sheet["source_path"] = file_path
#                 df_sheet["update_type"] = range(2,len(df_sheet)+2) #업데이트 시트 내에서 행 번호로 임시 업데이트 유형 지정 (1,2는 헤더이므로 2부터 시작)
                
#                 all_rows.append(df_sheet)
#                 self.append_log(f"업데이트 후보 추출 성공 : {file_path} (행: {len(df_sheet)})")
#             except Exception as e:
#                 self.append_log(f"업데이트 후보 추출 실패 : {file_path} - {str(e)}")
                
#         if not all_rows:
#             self.append_log("업데이트 후보로 추출된 데이터가 없습니다.")
#             return
    
#         self.df_candidates = pd.concat(all_rows, ignore_index=True)
#         self.append_log(f"업데이트 후보 통합 완료 : 총 {len(self.df_candidates)} 행")
#         self.show_dataframe_preview(self.df_candidates)
        
#     def show_result_preview(self,df: pd.DataFrame):
#         self.result_table.clear()
#         self.result_table.setRowCount(len(df))
#         self.result_table.setColumnCount(len(df.columns))
#         self.result_table.setHorizontalHeaderLabels([str(col) for col in df.columns])
        
#         for row_idx in range(len(df)):
#             for col_idx in range(len(df.columns)):
#                 value= df.iloc[row_idx, col_idx]
#                 display_value= "" if pd.isna(value) else str(value)
                
#                 item = QTableWidgetItem(display_value)
#                 self.result_table.setItem(row_idx,col_idx,item)
                
#         self.result_table.resizeColumnsToContents()
#         self.append_log("비교 결과 테이블 표시 완료")
        
#     #업데이트 데이터 엑셀파일 저장로직
#     def save_compare_result_to_excel(self):
#         if self.df_compare_result is None or self.df_compare_result.empty:
#             self.append_log("저장할 비교 결과가 없습니다. 먼저 신규/업데이트 분류를 실행해 주세요.")
#             return

#         default_name = f"species_compare_result_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

#         file_path, _ = QFileDialog.getSaveFileName(
#             self,
#             "비교 결과 저장",
#             default_name,
#             "Excel Files (*.xlsx)"
#         )

#         if not file_path:
#             self.append_log("비교 결과 저장이 취소되었습니다.")
#             return

#         try:
#             save_df = self.df_compare_result.copy()

#             # 원하면 상태별 정렬
#             if "status" in save_df.columns:
#                 status_order = {
#                     "신규": 0,
#                     "업데이트": 1,
#                     "업데이트(중복)": 2,
#                     "검토필요": 3,
#                     "식별정보부족": 4,
#                 }
#                 save_df["_status_order"] = save_df["status"].map(status_order).fillna(99)
#                 save_df = save_df.sort_values(by=["_status_order"]).drop(columns=["_status_order"])

#             save_df.to_excel(file_path, index=False)

#             self.append_log(f"비교 결과 저장 완료: {file_path}")

#         except Exception as e:
#             self.append_log(f"비교 결과 저장 실패: {str(e)}")

        
#     def compare_with_master(self):
#         if self.df is None:
#             self.append_log("종정보 마스터를 먼저 로드해 주세요.")
#             return
#         if self.df_candidates is None:
#             self.append_log("업데이트 후보를 먼저 추출해 주세요.")
#             return
        
#         master_df  = self.df
#         candidate_df  = self.df_candidates.copy()
        
#         required_master_cols = [
#             "종학명정보ID",
#             "분류체계정보ID",
#             "분류체계명(라틴)",
#             "분류체계명(국문)",
#             "학명",
#             "국명",
#         ]
        
        
#         if "학명" not in candidate_df.columns or "국명" not in candidate_df.columns:
#             self.append_log("업데이트 시트에 '학명' 또는 '국명' 컬럼이 없습니다.")
#             return
        
#         master_df["학명"] = master_df["학명"].apply(normalize_text)
#         master_df["국명"] = master_df["국명"].apply(normalize_text)
        
#         result_rows=[]
        
#         for _, row in candidate_df.iterrows():
#             cand_scnm = normalize_text(row.get("학명"))
#             cand_cnnm = normalize_text(row.get("국명"))
            
#             matched_by_scnm = master_df[master_df["학명"] == cand_scnm] if cand_scnm else pd.DataFrame()
#             matched_by_cnnm = master_df[master_df["국명"] == cand_cnnm] if cand_cnnm else pd.DataFrame()
            
#             matched_row = None
#             status = ""
#             match_basis = ""

            
#             if not cand_cnnm and not cand_scnm:
#                 status = "식별정보부족"
#                 match_basis = "학명/국명 모두 없음"

#             elif cand_scnm and cand_cnnm:
#                 exact_match = master_df[
#                     (master_df["학명"] == cand_scnm) &
#                     (master_df["국명"] == cand_cnnm)
#                 ]

#                 if len(exact_match) == 1:
#                     matched_row = exact_match.iloc[0]
#                     status = "업데이트"
#                     match_basis = "학명+국명 일치"

#                 elif len(exact_match) > 1:
#                     matched_row = exact_match.iloc[0]
#                     status = "업데이트(중복)"
#                     match_basis = "학명+국명 일치(마스터 내 중복)"

#                 else:
#                     scnm_cnt = len(matched_by_scnm)
#                     cnnm_cnt = len(matched_by_cnnm)

#                     if scnm_cnt == 0 and cnnm_cnt == 0:
#                         status = "신규"
#                         match_basis = "학명/국명 모두 미매칭"

#                     elif scnm_cnt == 1 and cnnm_cnt == 0:
#                         matched_row = matched_by_scnm.iloc[0]
#                         status = "검토필요"
#                         match_basis = "학명만 일치"

#                     elif scnm_cnt == 0 and cnnm_cnt == 1:
#                         matched_row = matched_by_cnnm.iloc[0]
#                         status = "검토필요"
#                         match_basis = "국명만 일치"

#                     elif scnm_cnt == 1 and cnnm_cnt == 1:
#                         scnm_row = matched_by_scnm.iloc[0]
#                         cnnm_row = matched_by_cnnm.iloc[0]

#                         if scnm_row["species_essnt_info_id"] == cnnm_row["species_essnt_info_id"]:
#                             matched_row = scnm_row
#                             status = "업데이트"
#                             match_basis = "학명/국명 동일 개체 매칭"
#                         else:
#                             status = "검토필요"
#                             match_basis = "학명 일치 종과 국명 일치 종이 서로 다름"

#                     elif scnm_cnt > 1 and cnnm_cnt == 0:
#                         matched_row = matched_by_scnm.iloc[0]
#                         status = "업데이트(중복)"
#                         match_basis = "학명 일치(마스터 내 중복)"

#                     elif scnm_cnt == 0 and cnnm_cnt > 1:
#                         matched_row = matched_by_cnnm.iloc[0]
#                         status = "업데이트(중복)"
#                         match_basis = "국명 일치(마스터 내 중복)"

#                     else:
#                         status = "검토필요"
#                         match_basis = "학명/국명 부분 일치 또는 중복 혼재"

#             elif cand_scnm and not cand_cnnm:
#                 if len(matched_by_scnm) == 1:
#                     matched_row = matched_by_scnm.iloc[0]
#                     status = "업데이트"
#                     match_basis = "학명 일치"
#                 elif len(matched_by_scnm) > 1:
#                     matched_row = matched_by_scnm.iloc[0]
#                     status = "업데이트(중복)"
#                     match_basis = "학명 일치(마스터 내 중복)"
#                 else:
#                     status = "신규"
#                     match_basis = "학명 미매칭"

#             elif cand_cnnm and not cand_scnm:
#                 if len(matched_by_cnnm) == 1:
#                     matched_row = matched_by_cnnm.iloc[0]
#                     status = "업데이트"
#                     match_basis = "국명 일치"
#                 elif len(matched_by_cnnm) > 1:
#                     matched_row = matched_by_cnnm.iloc[0]
#                     status = "업데이트(중복)"
#                     match_basis = "국명 일치(마스터 내 중복)"
#                 else:
#                     status = "신규"
#                     match_basis = "국명 미매칭"
            
#             result_row = row.to_dict()
#             result_row["status"] = status
#             result_row["match_basis"] = match_basis
            
#             if matched_row is not None:
#                 result_row["matched_specs_essnt_info_id"] = matched_row.get("종학명정보ID", "")
#                 result_row["matched_clssc_sstem_info_id"] = matched_row.get("분류체계정보ID", "")
#                 result_row["matched_clssc_name_latin"] = matched_row.get("분류체계명(라틴)", "")
#                 result_row["matched_clssc_name_kor"] = matched_row.get("분류체계명(국문)", "")
#                 result_row["matched_scnm"] = matched_row.get("학명", "")
#                 result_row["matched_cnnm"] = matched_row.get("국명", "")
#             else:
#                 result_row["matched_specs_essnt_info_id"] = ""
#                 result_row["matched_clssc_sstem_info_id"] = ""
#                 result_row["matched_clssc_name_latin"] = ""
#                 result_row["matched_clssc_name_kor"] = ""
#                 result_row["matched_scnm"] = ""
#                 result_row["matched_cnnm"] = ""
            
#             result_rows.append(result_row)

#         self.df_compare_result = pd.DataFrame(result_rows)
#         self.append_log(
#             f"비교 완료 - 신규: {(self.df_compare_result['status'] == '신규').sum()}건, "
#             f"업데이트: {(self.df_compare_result['status'] == '업데이트').sum()}건, "
#             f"검토필요: {(self.df_compare_result['status'] == '검토필요').sum()}건, "
#             f"식별정보부족: {(self.df_compare_result['status'] == '식별정보부족').sum()}건"
#         )

#         self.show_result_preview(self.df_compare_result.head(50))
#         reply = QMessageBox.question(
#             self,
#             "비교 결과 저장",
#             "비교가 완료되었습니다. 결과를 엑셀로 저장하시겠습니까?",
#             QMessageBox.Yes | QMessageBox.No,
#             QMessageBox.Yes
#         )

#         if reply == QMessageBox.Yes:
#             self.save_compare_result_to_excel()

#     #기존 종정보 업로드 파일 로드
#     def _load_master_from_excel(self):
#         self.append_log("원본 엑셀 로딩 중...")
#         self.df = pd.read_excel(MASTER_FILE_PATH, header=1)

#         CACHE_DIR.mkdir(parents=True, exist_ok=True)
#         self.df.to_pickle(MASTER_CACHE_FILE)

#         self.master_file_path.setText(str(MASTER_FILE_PATH))
#         self.append_log("원본 엑셀 로드 및 캐시 갱신 완료")
#         self.show_dataframe_preview(self.df)
#         self.status_label.setText("로드 완료")

#     #앱 시작시 자동로드
#     def load_last_species_master(self):
#         self.append_log("종정보 마스터 확인 중...")
#         self.status_label.setText("종정보 마스터 로딩 중...")
#         self.status_label.setStyleSheet("color: green;")
#         QApplication.processEvents()

#         try:
#             if MASTER_CACHE_FILE.exists():
#                 self.append_log("캐시 로딩 중...")
#                 self.df = pd.read_pickle(MASTER_CACHE_FILE)
#                 self.master_file_path.setText(str(MASTER_FILE_PATH))
#                 self.append_log(f"캐시 로드 성공: {MASTER_CACHE_FILE}")
#                 self.show_dataframe_preview(self.df)
#                 self.status_label.setText("로드 완료")
#                 self.status_label.setStyleSheet("color: red;")
#                 return

#             if MASTER_FILE_PATH.exists():
#                 self._load_master_from_excel()
#                 self.status_label.setText("로드 완료")
#                 self.status_label.setStyleSheet("color: red;")
#                 return

#             self.append_log("저장된 종정보 파일 없음")

#         except Exception as e:
#             self.status_label.setText("로드 실패")
#             self.append_log(f"자동 로드 실패: {str(e)}")

#     #다시 불러오기
#     def reload_master(self):
#         try:
#             if not MASTER_FILE_PATH.exists():
#                 self.append_log("재로딩할 원본 파일이 없습니다.")
#                 return

#             self._load_master_from_excel()
#             self.append_log("재로딩 완료")

#         except Exception as e:
#             self.append_log(f"재로딩 실패: {str(e)}")
            

# def normalize_text(value):
#         if pd.isna(value):
#             return ""
#         text = str(value).strip()
#         return "" if text.lower() == "nan" else text
    
# def find_target_sheet_name(file_path: str, target_sheet_name: str) -> str | None:
#     excel_file = pd.ExcelFile(file_path)
#     normalized_target = normalize_sheet_name(target_sheet_name)

#     for sheet_name in excel_file.sheet_names:
#         if normalize_sheet_name(sheet_name) == normalized_target:
#             return sheet_name
#     return None


