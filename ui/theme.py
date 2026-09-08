from PySide6.QtCore import Qt
from PySide6.QtWidgets import QAbstractSpinBox, QComboBox, QSpinBox

_ORIGINAL_COMBO_WHEEL_EVENT = QComboBox.wheelEvent
_ORIGINAL_SPINBOX_WHEEL_EVENT = QAbstractSpinBox.wheelEvent
_ORIGINAL_COMBO_INIT = QComboBox.__init__
_ORIGINAL_SPINBOX_INIT = QSpinBox.__init__


def install_scroll_wheel_guard():
    """드롭다운/스핀박스 위에 마우스만 올려도 휠 스크롤에 값이 바뀌는 게 불편해서,
    포커스(클릭)된 상태에서만 휠 입력을 받도록 막음.

    QComboBox/QSpinBox 기본 focusPolicy가 Qt.WheelFocus라서, 실제(스폰터니어스) 휠
    이벤트는 wheelEvent()가 불리기 전에 이미 그 스크롤 자체로 포커스를 받아버림
    (hasFocus()로 체크해도 "지금 막 이 스크롤이 준 포커스"라 걸러지지 않음 -- 합성
    이벤트로 테스트했을 땐 이 부작용이 없어서 통과했지만 실제 마우스 휠에서는 뚫림).
    그래서 생성 시점에 focusPolicy를 StrongFocus(클릭/Tab만 포커스, 휠은 제외)로
    바꿔치기해서 애초에 휠만으로는 포커스를 못 받게 막고, wheelEvent도 같이 가드함."""

    def guarded_combo_wheel_event(self, event):
        if self.hasFocus():
            _ORIGINAL_COMBO_WHEEL_EVENT(self, event)
        else:
            event.ignore()

    def guarded_spinbox_wheel_event(self, event):
        if self.hasFocus():
            _ORIGINAL_SPINBOX_WHEEL_EVENT(self, event)
        else:
            event.ignore()

    def combo_init_no_wheel_focus(self, *args, **kwargs):
        _ORIGINAL_COMBO_INIT(self, *args, **kwargs)
        self.setFocusPolicy(Qt.StrongFocus)

    def spinbox_init_no_wheel_focus(self, *args, **kwargs):
        _ORIGINAL_SPINBOX_INIT(self, *args, **kwargs)
        self.setFocusPolicy(Qt.StrongFocus)

    QComboBox.wheelEvent = guarded_combo_wheel_event
    QAbstractSpinBox.wheelEvent = guarded_spinbox_wheel_event
    QComboBox.__init__ = combo_init_no_wheel_focus
    QSpinBox.__init__ = spinbox_init_no_wheel_focus


STYLE_SHEET = """
QWidget {
    font-family: 'Segoe UI', 'Malgun Gothic';
    font-size: 13px;
    color: #222831;
}
QMainWindow, QStackedWidget {
    background: #f4f6f9;
}
QPushButton {
    background: #ffffff;
    border: 1px solid #d8dde5;
    border-radius: 6px;
    padding: 6px 14px;
}
QPushButton:hover {
    background: #eef1f6;
    border-color: #c3cad6;
}
QPushButton:pressed {
    background: #e4e8ee;
}
QLineEdit, QTextEdit, QComboBox, QListWidget {
    background: #ffffff;
    border: 1px solid #d8dde5;
    border-radius: 6px;
    padding: 5px 8px;
    selection-background-color: #1f8f5f;
}
QLineEdit:focus, QComboBox:focus {
    border: 1px solid #1f8f5f;
}
QTableWidget, QTableView {
    background: #ffffff;
    alternate-background-color: #f7f9fb;
    gridline-color: #eceff3;
    border: 1px solid #d8dde5;
    border-radius: 6px;
}
QHeaderView::section {
    background: #f4f6f9;
    color: #5a6070;
    padding: 6px;
    border: none;
    border-bottom: 1px solid #d8dde5;
    font-weight: 600;
}
QTabWidget::pane {
    border: 1px solid #d8dde5;
    border-radius: 6px;
    top: -1px;
}
QTabBar::tab {
    background: transparent;
    padding: 8px 16px;
    color: #5a6070;
    border-bottom: 2px solid transparent;
}
QTabBar::tab:selected {
    color: #1f8f5f;
    border-bottom: 2px solid #1f8f5f;
    font-weight: 600;
}
QScrollBar:vertical {
    background: transparent;
    width: 10px;
}
QScrollBar::handle:vertical {
    background: #c3cad6;
    border-radius: 5px;
    min-height: 24px;
}
QScrollBar::add-line, QScrollBar::sub-line {
    height: 0px;
}

QGroupBox {
    background: #ffffff;
    border: 1px solid #d8dde5;
    border-radius: 10px;
    margin-top: 16px;
    padding: 16px 14px 14px 14px;
    font-weight: 700;
    font-size: 14px;
    color: #17203a;
}
QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top left;
    left: 14px;
    padding: 0 6px;
    color: #1f8f5f;
}

/* App sidebar navigation */
#Sidebar {
    background: #17203a;
}
#SidebarLogo {
    color: #ffffff;
    font-size: 16px;
    font-weight: 700;
    padding: 18px 16px;
}
#SidebarNav {
    background: transparent;
    border: none;
    padding: 4px 8px;
    outline: none;
}
#SidebarNav::item {
    color: #9aa5c0;
    border-radius: 8px;
    padding: 10px 12px;
    margin: 2px 0px;
}
#SidebarNav::item:hover {
    background: #232d4d;
    color: #ffffff;
}
#SidebarNav::item:selected {
    background: #1f8f5f;
    color: #ffffff;
    font-weight: 600;
}
"""
