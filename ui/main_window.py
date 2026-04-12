from PySide6.QtWidgets import (
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QTextEdit,
    QFileDialog,
    QTableWidget,
    QTableWidgetItem,
)
import pandas as pd

# 앱의 가장 바깥 창(메인창)
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        
        self.setWindowTitle("BioFlow") #창 제목 설정
        self.resize(1000,700) #창 크기 설정

        self.init_ui()

    def init_ui(self):
        central_widget = QWidget() #메인 창 안에 들어갈 실제 화면영역
        self.setCentralWidget(central_widget)

        main_layout= QVBoxLayout() # 위에서 아래로 차곡차곡 쌓는 레이아웃
        central_widget.setLayout(main_layout)

        title_label = QLabel("BioFlow - 생물자원 데이터 자동화 도구")  #텍스트 보여주는 위젯
        main_layout.addWidget(title_label)

        #종정보 마스터 파일 선택 영역
        file_layout = QHBoxLayout()

        self.master_file_path = QLineEdit()
        self.master_file_path.setPlaceholderText("기존 종정보 엑셀 파일 경로")


        self.select_file_button = QPushButton("파일선택")
        self.select_file_button.clicked.connect(self.select_master_file)

        file_layout.addWidget(QLabel("종정보 마스터"))
        file_layout.addWidget(self.master_file_path)
        file_layout.addWidget(self.select_file_button)

        main_layout.addLayout(file_layout)

        self.load_button = QPushButton("엑셀 로드")
        self.load_button.clicked.connect(self.load_excel)
        main_layout.addWidget(self.load_button)

        self.preview_table= QTableWidget()
        main_layout.addWidget(self.preview_table)

        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setPlaceholderText("작업 로그가 여기에 표시됩니다.")
        main_layout.addWidget(self.log_text)

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

    def append_log(self,message: str):
        self.log_text.append(message)

    def load_excel(self):
        file_path = self.master_file_path.text()

        if not file_path:
            self.append_log("파일을 먼저 선택해 주세요")
            return
        
        try:
            df = pd.read_excel(file_path, header=1)

            self.append_log(f"엑셀 로드 성공 (행 : {len(df)} , 컬럼: {len(df.columns)})")

            #컬럼 출력
            self.append_log(f"컬럼 목록 : {list(df.columns)}")

            clean_columns = []
            for col in df.columns:
                if isinstance(col, (int, float)):
                    continue
                clean_columns.append(str(col).strip().replace("\n", " "))

            df = df.iloc[:, :len(clean_columns)]
            df.columns = clean_columns

            self.df = df
            self.append_log(f"엑셀 로드 완료 (행: {len(df)})")

            self.show_dataframe_preview(df)

        except Exception as e:
            self.append_log(f"엑셀 로드 실패 : {str(e)}")

    
    #엑셀 테이블 뷰
    def show_dataframe_preview(self, df: pd.DataFrame):
        preview_df = df.head(20)

        self.preview_table.clear()
        self.preview_table.setRowCount(len(preview_df))
        self.preview_table.setColumnCount(len(preview_df.columns))
        self.preview_table.setHorizontalHeaderLabels([str(col) for col in preview_df.columns])

        for row_idx in range(len(preview_df)):
            for col_idx in range(len(preview_df.columns)):
                value= preview_df.iloc[row_idx, col_idx]

                if pd.isna(value):
                    display_value=""
                else:
                    display_value=str(value)
                item = QTableWidgetItem(display_value)
                self.preview_table.setItem(row_idx,col_idx,item)

        self.preview_table.resizeColumnsToContents()
        self.append_log("미리보기 테이블 표시 완료(상위 20행)")