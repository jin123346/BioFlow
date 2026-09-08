import json
import os
from pathlib import Path

from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QFrame,
    QGraphicsDropShadowEffect,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from config.paths import (
    DATA_DIR,
    HISTORY_DIR,
    LOCAL_VERSION_FILE_PATH,
    LOCAL_VERSION_HISTORY_FILE_PATH,
    get_app_version,
)
from services.data_update_service import LOCAL_VERSION_FILE

REFERENCE_DIR = DATA_DIR / "reference"
DRAFT_TARGETS = [
    ("관찰 정보 작성 (신규 데이터)", REFERENCE_DIR / "new_data_draft.json"),
    ("표본정보 작성 (신규 데이터)", REFERENCE_DIR / "new_specimen_draft.json"),
]


def read_json(path: Path, fallback):
    try:
        with path.open("r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return fallback


class DashboardTab(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()
        self.refresh()

    def init_ui(self):
        layout = QVBoxLayout()
        layout.setSpacing(14)
        self.setLayout(layout)

        header_layout = QHBoxLayout()
        title_label = QLabel("대시보드")
        title_label.setStyleSheet("font-size: 20px; font-weight: 700; color: #17203a;")
        header_layout.addWidget(title_label)
        header_layout.addStretch()
        refresh_button = QPushButton("새로고침")
        refresh_button.clicked.connect(self.refresh)
        header_layout.addWidget(refresh_button)
        layout.addLayout(header_layout)

        self.version_label = QLabel("")
        self.version_label.setStyleSheet(
            "background: #eef1f6; color: #5a6070; border-radius: 6px; padding: 8px 12px;"
        )
        layout.addWidget(self.version_label)

        # 가장 자주 확인할 정보라 카드 하나를 크게 잡고(스트레치 3), 그림자로 강조
        draft_group = QGroupBox("임시 저장 작업 현황")
        draft_layout = QVBoxLayout()
        draft_layout.setSpacing(16)
        draft_group.setLayout(draft_layout)
        shadow = QGraphicsDropShadowEffect(blurRadius=24, xOffset=0, yOffset=4)
        shadow.setColor(QColor(23, 32, 58, 40))
        draft_group.setGraphicsEffect(shadow)

        self.draft_labels = {}
        for i, (title, path) in enumerate(DRAFT_TARGETS):
            if i > 0:
                divider = QFrame()
                divider.setFrameShape(QFrame.HLine)
                divider.setStyleSheet("background: #eceff3; max-height: 1px; border: none;")
                draft_layout.addWidget(divider)

            label = QLabel("")
            label.setWordWrap(True)
            self.draft_labels[str(path)] = label
            draft_layout.addWidget(label)
        draft_layout.addStretch(1)
        layout.addWidget(draft_group, 3)

        # 이력/백업은 참고용이라 작게, 좌우로 나란히
        bottom_row = QHBoxLayout()

        history_group = QGroupBox("업데이트 이력 (최근 5건)")
        history_layout = QVBoxLayout()
        history_group.setLayout(history_layout)
        self.history_list = QListWidget()
        self.history_list.setMaximumHeight(110)
        self.history_list.setStyleSheet("font-size: 11px;")
        history_layout.addWidget(self.history_list)
        bottom_row.addWidget(history_group, 1)

        backup_group = QGroupBox("마스터 백업 파일 (최근 5개)")
        backup_layout = QVBoxLayout()
        backup_group.setLayout(backup_layout)
        self.backup_list = QListWidget()
        self.backup_list.setMaximumHeight(110)
        self.backup_list.setStyleSheet("font-size: 11px;")
        backup_layout.addWidget(self.backup_list)
        open_backup_folder_button = QPushButton("백업 폴더 열기")
        open_backup_folder_button.clicked.connect(lambda: self.open_folder(HISTORY_DIR))
        backup_layout.addWidget(open_backup_folder_button)
        bottom_row.addWidget(backup_group, 1)

        layout.addLayout(bottom_row, 1)

    def refresh(self):
        self.refresh_version_label()
        self.refresh_draft_status()
        self.refresh_history_list()
        self.refresh_backup_list()

    def refresh_version_label(self):
        data_version = read_json(LOCAL_VERSION_FILE, {}).get("version") or "알 수 없음"
        last_saved = read_json(LOCAL_VERSION_FILE_PATH, {}).get("update_at") or "저장 이력 없음"
        app_version = get_app_version() or "알 수 없음"
        self.version_label.setText(
            f"앱 버전: {app_version}   |   기준 데이터 버전: {data_version}   |   마지막 로컬 저장: {last_saved}"
        )

    def refresh_draft_status(self):
        for title, path in DRAFT_TARGETS:
            label = self.draft_labels[str(path)]
            draft = read_json(path, {}) if path.exists() else {}
            rows = draft.get("rows") or []
            has_content = bool(rows)

            dot_color = "#1f8f5f" if has_content else "#c3cad6"
            detail = (
                f"{len(rows)}건 작성 중 · 마지막 자동저장 {draft.get('updated_at', '알 수 없음')}"
                if has_content
                else "작성 중인 내용 없음"
            )
            label.setText(
                f"<span style='color:{dot_color}; font-size:20px;'>●</span>"
                f"&nbsp;&nbsp;<span style='font-size:16px; font-weight:700; color:#17203a;'>{title}</span>"
                f"<br>&nbsp;&nbsp;&nbsp;&nbsp;"
                f"<span style='font-size:12px; color:#5a6070;'>{detail}</span>"
            )

    def refresh_history_list(self):
        self.history_list.clear()
        history = read_json(LOCAL_VERSION_HISTORY_FILE_PATH, [])
        if not isinstance(history, list):
            history = []

        for entry in reversed(history[-5:]):
            update_at = entry.get("update_at", "")
            file_name = entry.get("file_name", "")
            source = entry.get("source", "")
            self.history_list.addItem(f"{update_at}  -  {file_name}  ({source})")

        if not history:
            self.history_list.addItem("업데이트 이력이 없습니다.")

    def refresh_backup_list(self):
        self.backup_list.clear()
        if not HISTORY_DIR.exists():
            self.backup_list.addItem("백업 파일이 없습니다.")
            return

        backups = sorted(HISTORY_DIR.glob("*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
        for path in backups[:5]:
            self.backup_list.addItem(f"{path.name}")

        if not backups:
            self.backup_list.addItem("백업 파일이 없습니다.")

    def open_folder(self, path: Path):
        path.mkdir(parents=True, exist_ok=True)
        os.startfile(str(path))
