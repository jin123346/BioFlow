from PySide6.QtCore import QAbstractTableModel, QModelIndex, QObject, Qt, Signal
import pandas as pd

class PandasTableModel(QAbstractTableModel):
    def __init__(self, df: pd.DataFrame, parent=None):
        super().__init__(parent)
        self._df = df.copy()
        
    def rowCount(self, parent=QModelIndex()) -> int:
        return 0 if parent.isValid() else len(self._df)
    def columnCount(self, /, parent =QModelIndex()) -> int:
        return super().columnCount(parent)
    
    def data(self, index: QModelIndex, role: int = Qt.DisplayRole) :
        if not index.isValid():
            return None
        value= self.df.iat[index.row(), index.column() ]
        if role in (Qt.DisplayRole, Qt.EditRole):
            return "" if pd.isna(value) else str(value)
        return None
    
    def headerData(self,section:int, orientation:Qt.Orientation, role:int=Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal:
            return str(self._df.columns[section]) 
        return str(section + 1)
    
    def flags(self,index: QModelIndex):
        if not index.isValid():
            return Qt.ItemIsEnabled
        editable_columns = {"종고유ID", "국명", "학명", "매칭종정보ID", "분류코드", "매칭상태", "비고"}
        column_name = self._df.columns[index.column()]
        flags = Qt.ItemIsSelectable | Qt.ItemIsEnabled
        if column_name in editable_columns:
            flags |= Qt.ItemIsEditable
        return flags
    
    def setData(self, index: QModelIndex, value, role: int = Qt.EditRole):
        if role!=Qt.EditRole or not index.isValid():
            return False
        self._df.iat[index.row(), index.column()] = value
        self.dataChanged.emit(index, index)
        return True
    
    def set_dataframe(self,df: pd.DataFrame):
        self.beginResetModel()
        self._df = df.copy()
        self.endResetModel()
        
        
    def datafrane(self) -> pd.DataFrame:
        return self._df.copy()
    