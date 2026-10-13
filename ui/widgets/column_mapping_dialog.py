from __future__ import annotations

import re

import pandas as pd
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QHeaderView,
    QLabel,
    QMessageBox,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
)

IGNORE_LABEL = "(사용 안 함)"
UNMATCHED_COLOR = QColor("#fff3b0")


def header_key(name) -> str:
    """'위도\n(36.123****) 소수점으로입력' → '위도', '대륙/해양 국문' → '대륙해양국문' 처럼 비교용 키로 정규화."""
    name = re.split(r"[(（]", str(name), maxsplit=1)[0]
    return re.sub(r"[\s/]", "", name).lower().removesuffix("명")


def auto_map(excel_columns, columns, export_columns, saved: dict[str, str]) -> dict[str, str | None]:
    """엑셀 컬럼 → 프로그램 컬럼. ""=사용 안 함, None=매칭 못 함(사용자 확인 필요).
    우선순위: 저장된 매핑 > 프로그램 컬럼명 > 완성본 컬럼명."""
    lookup = {header_key(target): (source or "") for target, source in export_columns}
    lookup.update({header_key(column): column for column in columns})
    result: dict[str, str | None] = {}
    used: set[str] = set()
    for name in excel_columns:
        target = saved[name] if name in saved else lookup.get(header_key(name))
        if target and (target not in columns or target in used):
            target = None
        if target:
            used.add(target)
        result[name] = target
    return result


def apply_mapping(df: pd.DataFrame, mapping: dict[str, str | None], columns) -> pd.DataFrame:
    """columns 순서의 DataFrame으로 변환. 매핑 안 된 프로그램 컬럼은 빈칸(None)."""
    sources = {target: name for name, target in mapping.items() if target}
    return pd.DataFrame({column: (df[sources[column]] if column in sources else None) for column in columns}, index=df.index)


class ColumnMappingDialog(QDialog):
    """엑셀 컬럼마다 프로그램 컬럼을 고르는 창. 자동 매칭 결과를 미리 채우고, 못 맞춘 행은 노란색."""

    def __init__(self, df: pd.DataFrame, mapping: dict[str, str | None], columns, parent=None):
        super().__init__(parent)
        self.setWindowTitle("엑셀 컬럼 매핑")
        self.excel_columns = list(mapping)

        unmatched = sum(1 for target in mapping.values() if target is None)
        layout = QVBoxLayout(self)
        layout.addWidget(
            QLabel(
                f"엑셀 컬럼 {len(mapping)}개 중 {unmatched}개를 자동으로 맞추지 못했습니다(노란색).\n"
                "각 엑셀 컬럼을 넣을 프로그램 컬럼을 고르세요. 여기서 정한 매핑은 저장되어 다음부터 자동 적용됩니다."
            )
        )

        self.table = QTableWidget(len(mapping), 3)
        self.table.setHorizontalHeaderLabels(["엑셀 컬럼", "예시 값", "프로그램 컬럼"])
        self.table.verticalHeader().setVisible(False)
        self.combos: list[QComboBox] = []
        for row, name in enumerate(self.excel_columns):
            sample = df[name].dropna()
            sample_text = str(sample.iloc[0])[:40] if not sample.empty else ""
            name_item = QTableWidgetItem(str(name).replace("\n", " "))
            sample_item = QTableWidgetItem(sample_text)
            combo = QComboBox()
            combo.addItem(IGNORE_LABEL, "")
            for column in columns:
                combo.addItem(column, column)
            target = mapping[name]
            if target:
                combo.setCurrentIndex(combo.findData(target))
            else:
                name_item.setBackground(UNMATCHED_COLOR)
                sample_item.setBackground(UNMATCHED_COLOR)
            self.table.setItem(row, 0, name_item)
            self.table.setItem(row, 1, sample_item)
            self.table.setCellWidget(row, 2, combo)
            self.combos.append(combo)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        layout.addWidget(self.table, 1)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self.resize(760, 600)

    def mapping(self) -> dict[str, str]:
        return {name: combo.currentData() for name, combo in zip(self.excel_columns, self.combos)}

    def accept(self):
        targets = [target for target in self.mapping().values() if target]
        duplicates = sorted({target for target in targets if targets.count(target) > 1})
        if duplicates:
            QMessageBox.warning(self, "엑셀 컬럼 매핑", f"같은 프로그램 컬럼에 여러 엑셀 컬럼이 연결됐습니다: {', '.join(duplicates)}")
            return
        if not targets:
            QMessageBox.warning(self, "엑셀 컬럼 매핑", "연결된 컬럼이 하나도 없습니다.")
            return
        super().accept()


if __name__ == "__main__":
    cols = ["NO.", "종고유ID", "국명", "학명", "위도", "관찰 공개여부", "대륙/해양 국문", "관찰일"]
    export = [("관찰정보순번", None), ("관찰년월일", "관찰일")]
    excel = ["NO.", "종고유ID\n(국명입력시 자동생성)", "위도\n(36.123****) 소수점으로입력", "관찰\n공개여부",
             "대륙/해양 국문\n(드롭다운선택)", "관찰정보순번", "scientificName", "vernacularName"]
    m = auto_map(excel, cols, export, {"vernacularName": "국명"})
    assert m == {
        "NO.": "NO.", "종고유ID\n(국명입력시 자동생성)": "종고유ID", "위도\n(36.123****) 소수점으로입력": "위도",
        "관찰\n공개여부": "관찰 공개여부", "대륙/해양 국문\n(드롭다운선택)": "대륙/해양 국문",
        "관찰정보순번": "", "scientificName": None, "vernacularName": "국명",
    }, m
    out = apply_mapping(pd.DataFrame({name: [None] for name in excel} | {"NO.": [1], "vernacularName": ["독수리"]}), m, cols)
    assert list(out.columns) == cols and out.loc[0, "국명"] == "독수리" and out.loc[0, "학명"] is None
    print("ok")
