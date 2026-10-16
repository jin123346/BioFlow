from PySide6.QtGui import QIcon
from PySide6.QtWidgets import (
    QHBoxLayout, QLabel, QListWidget, QMainWindow, QStackedWidget,
    QTableWidget, QTableWidgetItem, QMessageBox, QVBoxLayout, QWidget,
)
from ui.tabs.dashboard_tab import DashboardTab
# 미구현 탭은 구현 완료 후 다시 활성화합니다.
# from ui.tabs.validataion_tab import ValidationTab
# from ui.tabs.gbif_tab import GbifTab
from config.paths import ASSETS_DIR




# 앱의 가장 바깥 창(메인창)
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.df= None  #엑셀 데이터프레임 저장할 변수

        
        self.setWindowTitle("BioFlow") #창 제목 설정
        self.setWindowIcon(QIcon(str(ASSETS_DIR / "favicon" / "icon_256.png"))) #창/작업표시줄 아이콘
        self.resize(1200,800) #창 크기 설정

        self.init_ui()

    def init_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root_layout = QHBoxLayout(central)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)

        sidebar = QWidget()
        sidebar.setObjectName("Sidebar")
        sidebar.setFixedWidth(200)
        sidebar_layout = QVBoxLayout(sidebar)
        sidebar_layout.setContentsMargins(0, 0, 0, 0)
        sidebar_layout.setSpacing(0)

        logo = QLabel("BioFlow")
        logo.setObjectName("SidebarLogo")
        sidebar_layout.addWidget(logo)

        self.nav_list = QListWidget()
        self.nav_list.setObjectName("SidebarNav")
        sidebar_layout.addWidget(self.nav_list)
        sidebar_layout.addStretch()
        root_layout.addWidget(sidebar)

        self.pages = QStackedWidget()
        root_layout.addWidget(self.pages, 1)

        self.dashboard_tab = DashboardTab()
        self._loaded_pages = {0: self.dashboard_tab}

        for page, label in [
            (self.dashboard_tab, "대시보드"),
            (QWidget(), "종정보 마스터"),
            (QWidget(), "신규 데이터"),
            (QWidget(), "신규/업데이트 종 정리"),
            (QWidget(), "SQL 생성"),
            # 미구현: 비교 결과, GBIF 데이터
        ]:
            self.pages.addWidget(page)
            self.nav_list.addItem(label)

        self.nav_list.currentRowChanged.connect(self.show_page)
        self.nav_list.setCurrentRow(0)

    def ensure_page(self, index):
        if index not in self._loaded_pages:
            if index == 1:
                from ui.tabs.species_master_tab import SpeciesMasterTab
                page = SpeciesMasterTab()
                page.master_changed.connect(self.refresh_species_consumers)
            elif index == 2:
                from ui.tabs.upload_tab import UploadTab
                page = UploadTab()
                if getattr(self, "_active_master_df", None) is not None:
                    page.load_species_master_lookup(self._active_master_df)
            elif index == 3:
                from ui.tabs.species_review_tab import SpeciesReviewTab
                page = SpeciesReviewTab(self.upload_tab)
            elif index == 4:
                from ui.tabs.sql_tab import SqlTab
                page = SqlTab()
            else:
                raise ValueError(index)
            placeholder = self.pages.widget(index)
            self.pages.removeWidget(placeholder)
            placeholder.deleteLater()
            self.pages.insertWidget(index, page)
            self._loaded_pages[index] = page
        return self._loaded_pages[index]

    def show_page(self, index):
        if index >= 0:
            self.pages.setCurrentWidget(self.ensure_page(index))

    def refresh_species_consumers(self, dataframe):
        self._active_master_df = dataframe
        editor = self._loaded_pages.get(2)
        if editor is not None:
            editor.load_species_master_lookup(dataframe)
        review = self._loaded_pages.get(3)
        if review is not None:
            review._master_df = dataframe

    @property
    def master_tab(self):
        return self.ensure_page(1)

    @property
    def upload_tab(self):
        return self.ensure_page(2)

    @property
    def species_review_tab(self):
        return self.ensure_page(3)

    def closeEvent(self, event):
        if 2 not in self._loaded_pages:
            event.accept()
            return
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
